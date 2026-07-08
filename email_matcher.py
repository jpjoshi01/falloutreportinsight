from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Iterable


def normalize_text(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").lower()).strip()


def split_keywords(raw_keywords: str | Iterable[str]) -> list[str]:
    if isinstance(raw_keywords, str):
        parts = re.split(r"[,;\n]+", raw_keywords)
    else:
        parts = list(raw_keywords)
    return [normalize_text(part) for part in parts if normalize_text(part)]


def flexible_match(text: str, keywords: str | Iterable[str], min_ratio: float = 0.72) -> tuple[bool, list[str]]:
    haystack = normalize_text(text)
    matched: list[str] = []
    for keyword in split_keywords(keywords):
        if keyword in haystack:
            matched.append(keyword)
            continue
        words = [word for word in keyword.split(" ") if len(word) > 2]
        if words and sum(1 for word in words if word in haystack) / len(words) >= 0.6:
            matched.append(keyword)
            continue
        if SequenceMatcher(None, keyword, haystack[: max(len(keyword) * 3, 80)]).ratio() >= min_ratio:
            matched.append(keyword)
    return bool(matched), matched
