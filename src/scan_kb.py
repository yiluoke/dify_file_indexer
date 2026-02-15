from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import hashlib
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import yaml
from tqdm import tqdm

from .extractors import extract_text_and_outline
from .redaction import redact_text
from .summarizer import make_extract_summary, extract_keywords
from .versioning import infer_version_key
from .classify import infer_system, infer_screen_id, infer_doc_type


@dataclasses.dataclass
class DocIndex:
    doc_id: str
    title: str
    path: str
    rel_path: str
    ext: str
    size_bytes: int
    updated_at: str  # ISO
    mtime_epoch: float
    sha1: str
    system: Optional[str]
    screen_id: Optional[str]
    doc_type: Optional[str]
    version_key: str
    headings: List[str]
    preview: str
    summary: str
    keywords: List[str]
    aliases: List[str]


@dataclasses.dataclass(frozen=True)
class ScanItem:
    """A scanned file.

    - path: actual file path to index
    - alias_from: if discovered via a Windows shortcut (.lnk), the shortcut path
    """

    path: Path
    alias_from: Optional[Path] = None


def sha1_file(path: Path, block_size: int = 1024 * 1024) -> str:
    h = hashlib.sha1()
    with path.open("rb") as f:
        while True:
            b = f.read(block_size)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def safe_relpath(path: Path, roots: List[Path]) -> str:
    # 가장 가까운 root로부터 상대경로
    for r in roots:
        try:
            return str(path.relative_to(r))
        except ValueError:
            continue
    return str(path)


def load_state(state_path: Path) -> Dict[str, Any]:
    if state_path.exists():
        try:
            return json.loads(state_path.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_state(state_path: Path, state: Dict[str, Any]) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def load_config(path: Path) -> Dict[str, Any]:
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def _norm_key(p: Path) -> str:
    # Windowsの大文字小文字差異や相対パス差異を吸収
    return os.path.normcase(os.path.abspath(str(p)))


def _is_within_any_root(target: Path, roots: List[Path]) -> bool:
    t = os.path.abspath(str(target))
    for r in roots:
        try:
            rr = os.path.abspath(str(r))
            if os.path.commonpath([t, rr]) == rr:
                return True
        except Exception:
            # ドライブが違う/UNC混在などでcommonpathが落ちる場合
            continue
    return False


def _resolve_lnk_windows(lnk_path: Path) -> Optional[Path]:
    """Resolve .lnk target path on Windows using COM.

    Returns None if not resolvable.
    """

    try:
        import win32com.client  # type: ignore
    except Exception:
        # pywin32未導入 or 非Windows
        return None

    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(str(lnk_path))
        target = getattr(sc, "Targetpath", "")
        if not target:
            return None
        return Path(target)
    except Exception:
        return None


def _resolve_lnk_chain_windows(lnk_path: Path, max_chain: int) -> Optional[Path]:
    cur = lnk_path
    for _ in range(max(1, int(max_chain))):
        if cur.suffix.lower() != ".lnk":
            return cur
        nxt = _resolve_lnk_windows(cur)
        if not nxt:
            return None
        cur = nxt
    return cur


def iter_files(
    roots: List[Path],
    include_ext: List[str],
    exclude_dirs: List[str],
    exclude_dir_keywords: List[str],
    exclude_path_regex: List[str],
    shortcut_cfg: Dict[str, Any],
) -> Iterable[ScanItem]:
    """Iterate files under roots.

    - Regular files: yield ScanItem(path=file)
    - Windows shortcuts (.lnk): resolve and yield ScanItem(path=target, alias_from=lnk)
      (and optionally follow directory targets as additional roots)
    """

    include_ext_l = {e.lower() for e in include_ext}
    exclude_dirs_l = {d.lower() for d in exclude_dirs}
    exclude_dir_kw_l = [k.lower() for k in exclude_dir_keywords if str(k).strip()]
    compiled = [re.compile(p) for p in exclude_path_regex]

    sc_enabled = bool(shortcut_cfg.get("enabled", False))
    sc_follow_dir = bool(shortcut_cfg.get("follow_dir_targets", True))
    sc_allow_outside = bool(shortcut_cfg.get("allow_outside_roots", False))
    sc_max_chain = int(shortcut_cfg.get("max_chain", 2))

    # BFSで「追加root（ショートカット先ディレクトリ）」にも対応
    queue: List[Path] = [r for r in roots]
    visited_dirs: set[str] = set()

    while queue:
        root = queue.pop(0)
        if not root.exists():
            continue

        root_key = _norm_key(root)
        if root_key in visited_dirs:
            continue
        visited_dirs.add(root_key)

        for dirpath, dirnames, filenames in os.walk(root):
            dpath = Path(dirpath)

            def _is_excluded_dirname(dn: str) -> bool:
                dnl = dn.lower()
                if dnl in exclude_dirs_l:
                    return True
                # keyword match (substring) e.g. 'old', 'backup', '削除'
                for kw in exclude_dir_kw_l:
                    if kw and kw in dnl:
                        return True
                return False

            # exclude dirs (in-place modify for os.walk)
            dirnames[:] = [dn for dn in dirnames if not _is_excluded_dirname(dn)]
            # exclude regex (dir)
            filtered_dirnames = []
            for dn in dirnames:
                full = str(dpath / dn)
                if any(rx.search(full) for rx in compiled):
                    continue
                filtered_dirnames.append(dn)
            dirnames[:] = filtered_dirnames

            for fn in filenames:
                p = dpath / fn
                p_str = str(p)
                # exclude regex (file)
                if any(rx.search(p_str) for rx in compiled):
                    continue

                # Windows shortcut (.lnk)
                if sc_enabled and p.suffix.lower() == ".lnk":
                    target = _resolve_lnk_chain_windows(p, sc_max_chain)
                    if not target:
                        continue
                    # allow policy: roots配下のみ（デフォルト）
                    if (not sc_allow_outside) and (not _is_within_any_root(target, roots)):
                        continue
                    if target.exists() and target.is_dir():
                        if sc_follow_dir:
                            # apply dir-name exclusions to shortcut directory targets too
                            if _is_excluded_dirname(target.name):
                                continue
                            t_str = str(target)
                            if any(rx.search(t_str) for rx in compiled):
                                continue
                            queue.append(target)
                        continue
                    if target.exists() and target.is_file() and target.suffix.lower() in include_ext_l:
                        yield ScanItem(path=target, alias_from=p)
                    continue

                # regular file
                if p.suffix.lower() not in include_ext_l:
                    continue
                yield ScanItem(path=p, alias_from=None)


def md_escape(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\r", "\n")


def build_markdown(doc: DocIndex) -> str:
    # YAML front matter for metadata
    fm = {
        "doc_id": doc.doc_id,
        "title": doc.title,
        "path": doc.path,
        "rel_path": doc.rel_path,
        "ext": doc.ext,
        "size_bytes": doc.size_bytes,
        "updated_at": doc.updated_at,
        "sha1": doc.sha1,
        "system": doc.system,
        "screen_id": doc.screen_id,
        "doc_type": doc.doc_type,
        "version_key": doc.version_key,
        "keywords": doc.keywords,
        "aliases": doc.aliases,
    }
    # NOTE: Difyのチャンクで見出しが壊れるケースがあるため、本文にも明示ラベルを入れる
    lines = []
    lines.append("---")
    lines.append(yaml.safe_dump(fm, allow_unicode=True, sort_keys=False).strip())
    lines.append("---\n")
    lines.append(f"# {md_escape(doc.title)}\n")
    lines.append("## PATH\n")
    lines.append(f"- {md_escape(doc.path)}\n")
    if doc.aliases:
        lines.append("## ALIASES (shortcuts / links)\n")
        for a in doc.aliases[:200]:
            lines.append(f"- {md_escape(a)}")
        lines.append("")
    lines.append("## METADATA\n")
    lines.append(f"- system: {doc.system or ''}")
    lines.append(f"- screen_id: {doc.screen_id or ''}")
    lines.append(f"- doc_type: {doc.doc_type or ''}")
    lines.append(f"- updated_at: {doc.updated_at}")
    lines.append(f"- version_key: {doc.version_key}")
    lines.append(f"- sha1: {doc.sha1}\n")
    if doc.headings:
        lines.append("## HEADINGS\n")
        for h in doc.headings[:80]:
            lines.append(f"- {md_escape(h)}")
        lines.append("")
    if doc.preview:
        lines.append("## PREVIEW (limited)\n")
        lines.append(md_escape(doc.preview))
        lines.append("")
    if doc.summary:
        lines.append("## SUMMARY\n")
        lines.append(md_escape(doc.summary))
        lines.append("")
    if doc.keywords:
        lines.append("## KEYWORDS\n")
        lines.append(", ".join(doc.keywords))
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def _read_front_matter(md_text: str) -> Optional[Dict[str, Any]]:
    m = re.match(r"^---\n(.*?)\n---\n", md_text, flags=re.DOTALL)
    if not m:
        return None
    try:
        return yaml.safe_load(m.group(1))
    except Exception:
        return None


def _upsert_aliases_in_existing_md(md_text: str, aliases: List[str]) -> str:
    if not aliases:
        return md_text

    # Update YAML front matter (add/replace aliases)
    fm = _read_front_matter(md_text)
    if fm is None:
        return md_text
    fm["aliases"] = aliases

    # rewrite front matter block
    dumped = yaml.safe_dump(fm, allow_unicode=True, sort_keys=False).strip()
    md_text = re.sub(r"^---\n.*?\n---\n", f"---\n{dumped}\n---\n", md_text, count=1, flags=re.DOTALL)

    # Update/insert ALIASES section
    alias_lines = ["## ALIASES (shortcuts / links)", ""] + [f"- {md_escape(a)}" for a in aliases[:200]] + [""]
    alias_block = "\n".join(alias_lines)

    if re.search(r"^## ALIASES \(shortcuts / links\)\s*$", md_text, flags=re.MULTILINE):
        # replace the section until the next '## ' header or EOF
        md_text = re.sub(
            r"^## ALIASES \(shortcuts / links\)\n.*?(?=^## \w|\Z)",
            alias_block + "\n",
            md_text,
            flags=re.DOTALL | re.MULTILINE,
        )
    else:
        # insert after PATH section (best effort)
        md_text = re.sub(
            r"(^## PATH\n\n- .*?\n)(\n)",
            r"\1\n" + alias_block + r"\n\2",
            md_text,
            count=1,
            flags=re.DOTALL | re.MULTILINE,
        )
    return md_text


def write_latest_map(out_dir: Path, docs: List[DocIndex], allow_fallback: bool) -> None:
    # group key
    groups: Dict[Tuple[str, str, str], List[DocIndex]] = {}
    for d in docs:
        system = d.system or ""
        screen = d.screen_id or ""
        dtype = d.doc_type or ""
        if system and screen and dtype:
            k = (system, screen, dtype)
        elif allow_fallback:
            # fallback hierarchy: system+dtype+title (screen unknown)
            if system and dtype:
                k = (system, "__NO_SCREEN__", dtype)
            elif system:
                k = (system, "__NO_SCREEN__", "__NO_TYPE__")
            else:
                k = ("__NO_SYSTEM__", "__NO_SCREEN__", "__NO_TYPE__")
        else:
            continue
        groups.setdefault(k, []).append(d)

    def pick_latest(items: List[DocIndex]) -> DocIndex:
        # version_key is a lexicographically sortable key by design
        return sorted(items, key=lambda x: (x.version_key, x.mtime_epoch), reverse=True)[0]

    lines = []
    lines.append("# latest_map (PoC)\n")
    lines.append("最新版の推定結果。根拠は `version_key` と `updated_at`。\n")
    for k in sorted(groups.keys()):
        latest = pick_latest(groups[k])
        system, screen, dtype = k
        lines.append(f"## {system} / {screen} / {dtype}")
        lines.append(f"- latest_title: {latest.title}")
        lines.append(f"- latest_path: {latest.path}")
        lines.append(f"- updated_at: {latest.updated_at}")
        lines.append(f"- version_key: {latest.version_key}\n")

    (out_dir / "latest_map.md").write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True, help="config.yml path")
    ap.add_argument("--out", required=True, help="output dir")
    ap.add_argument("--state", default=None, help="state.json path (default: <out>/state.json)")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    config_path = Path(args.config)
    out_dir = Path(args.out)
    out_docs = out_dir / "docs"
    out_docs.mkdir(parents=True, exist_ok=True)
    state_path = Path(args.state) if args.state else out_dir / "state.json"

    cfg = load_config(config_path)

    roots = [Path(p) for p in cfg.get("roots", [])]
    include_ext = cfg.get("include_ext", [])
    exclude_dirs = cfg.get("exclude_dirs", [])
    exclude_dir_keywords = cfg.get("exclude_dir_keywords", [])
    exclude_path_regex = cfg.get("exclude_path_regex", [])

    limits = {
        "max_extract_chars": int(cfg.get("max_extract_chars", 8000)),
        "max_headings": int(cfg.get("max_headings", 40)),
        "max_preview_paragraphs": int(cfg.get("max_preview_paragraphs", 12)),
        "max_preview_cells": int(cfg.get("max_preview_cells", 80)),
        "max_preview_slides": int(cfg.get("max_preview_slides", 30)),
    }

    state = load_state(state_path)
    # key -> {mtime,size,sha1,doc_id,path}
    prev = state.get("files", {})

    redact_cfg = cfg.get("redact", {"enabled": False})
    sys_cfg = cfg.get("system_from_path", {"enabled": True, "depth_from_root": 1})
    screen_rx = cfg.get("screen_id_regex", [])
    doc_type_rules = cfg.get("doc_type_rules", [])
    latest_cfg = cfg.get("latest_map", {"enabled": True, "allow_fallback_keys": True})
    shortcut_cfg = cfg.get("shortcuts", {"enabled": False})

    docs: List[DocIndex] = []
    new_state_files: Dict[str, Any] = {}

    items = list(iter_files(roots, include_ext, exclude_dirs, exclude_dir_keywords, exclude_path_regex, shortcut_cfg))

    # aggregate aliases by normalized target path
    alias_map: Dict[str, List[str]] = {}
    unique_targets: Dict[str, Path] = {}
    for it in items:
        k = _norm_key(it.path)
        unique_targets.setdefault(k, it.path)
        if it.alias_from is not None:
            alias_map.setdefault(k, []).append(str(it.alias_from))

    for k, p in tqdm(list(unique_targets.items()), desc="Scanning"):
        try:
            st = p.stat()
        except Exception:
            continue

        p_str = str(p)
        mtime = st.st_mtime

        aliases = sorted(set(alias_map.get(k, [])))

        # quick change detection by mtime+size (sha1は重いので、必要時だけ)
        prev_ent = prev.get(k) or prev.get(p_str)
        if prev_ent and prev_ent.get("mtime") == mtime and prev_ent.get("size") == st.st_size:
            # unchanged: keep state and (if needed) update aliases in existing md
            doc_id = prev_ent.get("doc_id") or hashlib.sha1((k).encode("utf-8", errors="ignore")).hexdigest()[:16]
            md_path = out_docs / f"{doc_id}.md"
            if (not args.dry_run) and md_path.exists() and aliases:
                try:
                    md_text = md_path.read_text(encoding="utf-8")
                    md_text2 = _upsert_aliases_in_existing_md(md_text, aliases)
                    if md_text2 != md_text:
                        md_path.write_text(md_text2, encoding="utf-8")
                except Exception:
                    pass

            # for latest_map: read metadata from existing md front matter
            if md_path.exists():
                try:
                    md_text = md_path.read_text(encoding="utf-8")
                    fm = _read_front_matter(md_text) or {}
                    docs.append(
                        DocIndex(
                            doc_id=str(fm.get("doc_id", doc_id)),
                            title=str(fm.get("title", Path(p_str).stem)),
                            path=str(fm.get("path", p_str)),
                            rel_path=str(fm.get("rel_path", safe_relpath(p, roots))),
                            ext=str(fm.get("ext", p.suffix.lower())),
                            size_bytes=int(fm.get("size_bytes", int(st.st_size))),
                            updated_at=str(fm.get("updated_at", dt.datetime.fromtimestamp(mtime).isoformat(timespec="seconds"))),
                            mtime_epoch=float(mtime),
                            sha1=str(fm.get("sha1", prev_ent.get("sha1", ""))),
                            system=fm.get("system"),
                            screen_id=fm.get("screen_id"),
                            doc_type=fm.get("doc_type"),
                            version_key=str(fm.get("version_key", infer_version_key(p.name, mtime))),
                            headings=[],
                            preview="",
                            summary="",
                            keywords=list(fm.get("keywords", [])) if isinstance(fm.get("keywords", []), list) else [],
                            aliases=aliases
                            or (list(fm.get("aliases", [])) if isinstance(fm.get("aliases", []), list) else []),
                        )
                    )
                except Exception:
                    pass

            new_state_files[k] = {
                "mtime": mtime,
                "size": st.st_size,
                "sha1": prev_ent.get("sha1", ""),
                "doc_id": doc_id,
                "path": p_str,
            }
            continue

        # compute sha1 only for changed files (still may be heavy; acceptable for PoC)
        try:
            file_sha1 = sha1_file(p)
        except Exception:
            file_sha1 = ""

        rel = safe_relpath(p, roots)
        title = p.stem

        extracted = extract_text_and_outline(p, limits)
        headings = extracted.get("headings", [])[: limits["max_headings"]]
        preview = extracted.get("preview", "")[: limits["max_extract_chars"]]

        if redact_cfg.get("enabled", False):
            preview = redact_text(preview, redact_cfg.get("patterns", []))
            headings = [redact_text(h, redact_cfg.get("patterns", [])) for h in headings]

        # classification
        alias_names = [Path(a).stem for a in aliases]
        classify_text = title + "\n" + "\n".join(alias_names) + "\n" + preview
        system = infer_system(p, roots, sys_cfg)
        screen_id = infer_screen_id(classify_text, screen_rx)
        doc_type = infer_doc_type(classify_text, doc_type_rules)

        version_key = infer_version_key(p.name, mtime)

        summary = make_extract_summary(preview, max_sentences=int(cfg.get("summary_sentences", 3)))
        keywords = extract_keywords(title + "\n" + "\n".join(alias_names) + "\n" + preview, topk=int(cfg.get("keywords_topk", 15)))

        doc_id = hashlib.sha1((k).encode("utf-8", errors="ignore")).hexdigest()[:16]
        updated_at = dt.datetime.fromtimestamp(mtime).isoformat(timespec="seconds")

        doc = DocIndex(
            doc_id=doc_id,
            title=title,
            path=p_str,
            rel_path=rel,
            ext=p.suffix.lower(),
            size_bytes=int(st.st_size),
            updated_at=updated_at,
            mtime_epoch=mtime,
            sha1=file_sha1,
            system=system,
            screen_id=screen_id,
            doc_type=doc_type,
            version_key=version_key,
            headings=headings,
            preview=preview,
            summary=summary,
            keywords=keywords,
            aliases=aliases,
        )
        docs.append(doc)

        new_state_files[k] = {"mtime": mtime, "size": st.st_size, "sha1": file_sha1, "doc_id": doc_id, "path": p_str}

        if not args.dry_run:
            md = build_markdown(doc)
            (out_docs / f"{doc.doc_id}.md").write_text(md, encoding="utf-8")

    # state save
    state_out = {"generated_at": dt.datetime.now().isoformat(timespec="seconds"), "files": new_state_files}
    if not args.dry_run:
        save_state(state_path, state_out)

    # latest map
    if latest_cfg.get("enabled", True):
        if not docs:
            # include docs from previous run by reading existing md? (PoC: skip)
            pass
        if not args.dry_run:
            write_latest_map(out_dir, docs, allow_fallback=bool(latest_cfg.get("allow_fallback_keys", True)))

    print(f"Done. docs={len(docs)} out={out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
