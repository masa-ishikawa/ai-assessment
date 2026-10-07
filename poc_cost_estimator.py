"""OCI List Pricing APIを利用したPoC月額概算。"""

from __future__ import annotations

import json
from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import poc_cost_config as config


@dataclass(frozen=True)
class CostLine:
    name: str
    monthly_jpy: int
    assumption: str
    source: str


def _fetch_list_unit_price(part_number: str, currency: str) -> Decimal:
    """OCI公式List Pricing APIからPAY_AS_YOU_GOの単価を取得する。"""
    query = urlencode({"partNumber": part_number, "currencyCode": currency})
    request = Request(f"{config.PRICE_API_URL}?{query}", headers={"Accept": "application/json"})
    with urlopen(request, timeout=15) as response:  # nosec B310 - Oracle公式の固定HTTPS URL
        payload: dict[str, Any] = json.load(response)
    items = payload.get("items", [])
    if not items:
        raise ValueError(f"SKU {part_number} の価格情報が見つかりません。")
    for price_group in items[0].get("prices", []):
        if price_group.get("currencyCode") != currency:
            continue
        for price in price_group.get("prices", []):
            if price.get("model") == "PAY_AS_YOU_GO":
                return Decimal(str(price["value"]))
    raise ValueError(f"SKU {part_number} に {currency} のPAY_AS_YOU_GO価格がありません。")


def build_poc_cost_estimate() -> dict[str, Any]:
    """設定済みSKUはAPI単価、未設定項目は明示的なフォールバック概算で集計する。"""
    lines: list[CostLine] = []
    api_priced_count = 0
    for item in config.COST_ITEMS:
        part_number = item.get("part_number")
        quantity = item.get("monthly_quantity")
        if part_number and quantity is not None:
            try:
                unit_price = _fetch_list_unit_price(str(part_number), config.CURRENCY)
                monthly_jpy = int((unit_price * Decimal(str(quantity))).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
                source = f"OCI List Pricing API（SKU: {part_number}）"
                api_priced_count += 1
            except (OSError, ValueError, KeyError, json.JSONDecodeError):
                monthly_jpy = int(item["fallback_monthly_jpy"])
                source = "前提値（List Pricing API取得失敗時のフォールバック）"
        else:
            monthly_jpy = int(item["fallback_monthly_jpy"])
            source = "前提値（SKU未設定）"
        lines.append(CostLine(str(item["name"]), monthly_jpy, str(item["assumption"]), source))

    return {
        "currency": config.CURRENCY,
        "lines": lines,
        "total_monthly_jpy": sum(line.monthly_jpy for line in lines),
        "notes": config.NOTES,
        "api_priced_count": api_priced_count,
    }
