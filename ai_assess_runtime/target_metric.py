"""事業目標として表示する割合・金額の正規化。"""

import re


def _display_number(raw: str) -> str:
    number = float(raw.replace(",", ""))
    return str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".")


def _normalize_percentage_target(value: object) -> str:
    """説明付きの単一割合を、意味を変えず表示値へ正規化する。"""
    text = str(value or "").strip().replace("％", "%")
    matches = list(re.finditer(
        r"(?<![\d.])([+\-−－]?)(\d+(?:\.\d+)?)\s*"
        # Models commonly write ranges as either ``20〜35%`` or
        # ``20%〜35%``.  Accept both spellings and canonicalize them below.
        r"(?:%\s*)?(?:[〜～~\-–—]\s*(\d+(?:\.\d+)?)\s*)?%",
        text,
    ))
    valid: list[tuple[str, str | None]] = []
    for match in matches:
        sign, low_raw, high_raw = match.group(1), match.group(2), match.group(3)
        low = float(low_raw)
        high = float(high_raw or low_raw)
        if not sign and 0 < low <= high <= 100:
            valid.append((low_raw, high_raw))
    if len({(low, high or "") for low, high in valid}) != 1:
        return text
    low_raw, high_raw = valid[0]
    low_text = _display_number(low_raw)
    high_text = _display_number(high_raw) if high_raw else ""
    return f"{low_text}〜{high_text}%" if high_text else f"{low_text}%"


def _percentage_range_parts(value: object) -> tuple[float, float] | None:
    text = _normalize_percentage_target(value)
    match = re.fullmatch(r"(\d+(?:\.\d+)?)(?:〜(\d+(?:\.\d+)?))?%", text)
    if not match:
        return None
    low = float(match.group(1))
    high = float(match.group(2) or match.group(1))
    return (low, high) if 0 < low <= high <= 100 else None


_MONEY_UNIT = r"兆円|億円|百万円|万円|千円|円"


def _monetary_target_parts(value: object) -> tuple[float, float, str] | None:
    """単一の円貨目標または同一単位の金額レンジを返す。"""
    text = re.sub(r"\s+", "", str(value or "")).replace("，", ",")
    patterns = (
        rf"(\d[\d,]*(?:\.\d+)?)({_MONEY_UNIT})[〜～~\-–—](\d[\d,]*(?:\.\d+)?)\2",
        rf"(\d[\d,]*(?:\.\d+)?)(?:[〜～~\-–—](\d[\d,]*(?:\.\d+)?))?({_MONEY_UNIT})",
    )
    candidates: list[tuple[float, float, str]] = []
    for pattern_index, pattern in enumerate(patterns):
        for match in re.finditer(pattern, text):
            if pattern_index == 0:
                low_raw, unit, high_raw = match.group(1), match.group(2), match.group(3)
            else:
                low_raw, high_raw, unit = match.group(1), match.group(2), match.group(3)
            low = float(low_raw.replace(",", ""))
            high = float((high_raw or low_raw).replace(",", ""))
            if 0 < low <= high:
                candidates.append((low, high, unit))
        if candidates:
            break
    distinct = set(candidates)
    return candidates[0] if len(distinct) == 1 else None


def _normalize_monetary_target(value: object) -> str:
    parts = _monetary_target_parts(value)
    if not parts:
        return str(value or "").strip()
    low, high, unit = parts
    low_text = _display_number(str(low))
    high_text = _display_number(str(high))
    return f"{low_text}{unit}" if low == high else f"{low_text}〜{high_text}{unit}"


def _verified_management_amount_target(
        targets: object, value: object, category_keywords: tuple[str, ...],
        claim_id: object = "") -> dict | None:
    """原典情報を備え、同じKPIカテゴリ・金額に一致する中計目標を返す。"""
    amount = _normalize_monetary_target(value)
    requested_claim = str(claim_id or "").strip()
    matches: list[dict] = []
    for target in targets if isinstance(targets, list) else []:
        if not isinstance(target, dict) or target.get("claim_status") != "verified_external":
            continue
        if not all(str(target.get(field) or "").strip() for field in (
            "claim_id", "source_id", "source_url", "source_locator", "source_metric",
            "evidence_excerpt", "target",
        )) or not str(target.get("source_url")).startswith("https://"):
            continue
        if requested_claim and str(target.get("claim_id")) != requested_claim:
            continue
        evidence_kpi = f"{target.get('label', '')} {target.get('source_metric', '')}"
        if not any(keyword.lower() in evidence_kpi.lower() for keyword in category_keywords):
            continue
        if _normalize_monetary_target(target.get("target")) == amount:
            matches.append(target)
    return matches[0] if len(matches) == 1 else None
