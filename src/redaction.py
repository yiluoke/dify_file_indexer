from __future__ import annotations

import re
from typing import Any, Dict, List


def redact_text(text: str, patterns: List[Dict[str, Any]]) -> str:
    if not text or not patterns:
        return text
    out = text
    for p in patterns:
        try:
            rx = re.compile(p["regex"])
            repl = p.get("replace", "[REDACTED]")
            out = rx.sub(repl, out)
        except Exception:
            continue
    return out
