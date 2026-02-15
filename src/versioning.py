from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Tuple

# Produce lexicographically sortable key (higher is newer)
# priority order: explicit date in name > semantic version > revision > mtime
# key format: P{priority}-D{yyyymmdd}-V{semver}-R{rev}-M{mtime_int}

_RX_DATE1 = re.compile(r"((?:19|20)\d{2})[./_-]?(0[1-9]|1[0-2])[./_-]?(0[1-9]|[12]\d|3[01])")
_RX_DATE_JP = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月(0?[1-9]|[12]\d|3[01])日")
_RX_SEMVER = re.compile(r"(?i)\b(?:v|ver|version)[-_ ]?(\d+(?:\.\d+){0,3})\b")
_RX_REV = re.compile(r"(?i)\b(?:rev|r)[-_ ]?(\d{1,3})\b")


def _semver_tuple(v: str) -> Tuple[int, int, int, int]:
    parts = v.split(".")
    nums = [int(p) if p.isdigit() else 0 for p in parts][:4]
    while len(nums) < 4:
        nums.append(0)
    return tuple(nums)  # type: ignore


def infer_version_key(filename: str, mtime_epoch: float) -> str:
    name = filename

    date_int = 0
    m = _RX_DATE1.search(name)
    if m:
        y, mo, d = m.group(1), m.group(2), m.group(3)
        date_int = int(f"{y}{mo}{d}")
    else:
        mj = _RX_DATE_JP.search(name)
        if mj:
            y, mo, d = mj.group(1), int(mj.group(2)), int(mj.group(3))
            date_int = int(f"{y}{mo:02d}{d:02d}")

    semver = (0, 0, 0, 0)
    ms = _RX_SEMVER.search(name)
    if ms:
        semver = _semver_tuple(ms.group(1))

    rev = 0
    mr = _RX_REV.search(name)
    if mr:
        try:
            rev = int(mr.group(1))
        except ValueError:
            rev = 0

    mtime_int = int(mtime_epoch)

    # priority
    pr = 0
    if date_int:
        pr = 3
    elif semver != (0, 0, 0, 0):
        pr = 2
    elif rev:
        pr = 1
    else:
        pr = 0

    v_str = "".join(f"{x:03d}" for x in semver)  # fixed width
    key = f"P{pr}-D{date_int:08d}-V{v_str}-R{rev:03d}-M{mtime_int:010d}"
    return key
