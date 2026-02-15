from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional


def infer_system(file_path: Path, roots: List[Path], cfg: Dict[str, Any]) -> Optional[str]:
    if not cfg.get("enabled", True):
        return None
    depth = int(cfg.get("depth_from_root", 1))
    # choose nearest root
    for r in roots:
        try:
            rel = file_path.relative_to(r)
            parts = list(rel.parts)
            if len(parts) >= depth:
                return parts[depth - 1]
        except Exception:
            continue
    return None


def infer_screen_id(text: str, regex_list: List[str]) -> Optional[str]:
    for pat in regex_list or []:
        try:
            rx = re.compile(pat)
            m = rx.search(text)
            if m:
                return m.group(1)
        except Exception:
            continue
    return None


def infer_doc_type(text: str, rules: List[Dict[str, Any]]) -> Optional[str]:
    t = text.lower()
    for r in rules or []:
        words = r.get("contains_any", [])
        for w in words:
            if (w or "").lower() in t:
                return r.get("doc_type")
    return None
