from __future__ import annotations

import math
import re
from typing import Dict, List, Tuple

# Very lightweight extractive summarizer (no ML, no external API)


_SENT_SPLIT_RX = re.compile(r"(?<=[。！？!?\.])\s+")


def _normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _sentences(text: str) -> List[str]:
    t = _normalize(text)
    if not t:
        return []
    sents = _SENT_SPLIT_RX.split(t)
    # fallback: split by newlines if punctuation missing
    if len(sents) <= 1:
        sents = [x.strip() for x in text.splitlines() if x.strip()]
    return [s.strip() for s in sents if len(s.strip()) >= 10]


def _tokens(text: str) -> List[str]:
    # Works for mixed JP/EN: extract Kanji/Kana sequences, latin words, numbers
    toks = re.findall(r"[\u3040-\u30ff\u4e00-\u9fff]{2,}|[A-Za-z]{3,}|[0-9]{2,}", text)
    return [t.lower() for t in toks]


def make_extract_summary(text: str, max_sentences: int = 3) -> str:
    sents = _sentences(text)
    if not sents:
        return ""
    # term frequency
    tf: Dict[str, int] = {}
    for tok in _tokens(text):
        tf[tok] = tf.get(tok, 0) + 1

    # sentence scoring
    scored: List[Tuple[float, int, str]] = []
    for i, s in enumerate(sents[:80]):  # cap
        toks = _tokens(s)
        if not toks:
            continue
        score = sum(math.log1p(tf.get(t, 0)) for t in toks) / (1.0 + len(toks) ** 0.5)
        scored.append((score, i, s))

    if not scored:
        return " / ".join(sents[:max_sentences])

    top = sorted(scored, key=lambda x: x[0], reverse=True)[:max_sentences]
    # keep original order
    top_sorted = sorted(top, key=lambda x: x[1])
    return " ".join([t[2] for t in top_sorted])


def extract_keywords(text: str, topk: int = 15) -> List[str]:
    tf: Dict[str, int] = {}
    for tok in _tokens(text):
        if len(tok) <= 2:
            continue
        tf[tok] = tf.get(tok, 0) + 1
    items = sorted(tf.items(), key=lambda x: (x[1], len(x[0])), reverse=True)[:topk]
    return [k for k, _ in items]
