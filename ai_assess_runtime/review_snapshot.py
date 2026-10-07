"""凍結JSON、成果物ハッシュ、費用スナップショットの決定的な変換処理。"""

import copy
import hashlib
import json
from pathlib import Path
from types import SimpleNamespace


def canonical_json_bytes(value: object) -> bytes:
    """順序や整形差に左右されないJSONハッシュ用のバイト列を返す。"""
    return json.dumps(
        value, ensure_ascii=False, sort_keys=True, separators=(",", ":"), allow_nan=False,
    ).encode("utf-8")


def canonical_sha256(value: object) -> str:
    return hashlib.sha256(canonical_json_bytes(value)).hexdigest()


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def without_provenance(value: object) -> object:
    """自己参照を避け、実質的なリサーチ内容だけをハッシュ化する。"""
    copied = copy.deepcopy(value)
    if isinstance(copied, dict):
        copied.pop("provenance", None)
    return copied


def cost_estimate_snapshot(cost_estimate: dict | None) -> dict:
    """外部価格取得結果を、再描画可能なJSONスナップショットへ固定する。"""
    if not isinstance(cost_estimate, dict):
        return {}
    lines = []
    for line in cost_estimate.get("lines", []):
        if isinstance(line, dict):
            name, monthly_jpy = line.get("name"), line.get("monthly_jpy")
            assumption, source = line.get("assumption"), line.get("source")
        else:
            name, monthly_jpy = getattr(line, "name", None), getattr(line, "monthly_jpy", None)
            assumption, source = getattr(line, "assumption", None), getattr(line, "source", None)
        if not str(name or "").strip() or not isinstance(monthly_jpy, int):
            return {}
        lines.append({
            "name": str(name), "monthly_jpy": monthly_jpy,
            "assumption": str(assumption or ""), "source": str(source or ""),
        })
    total = cost_estimate.get("total_monthly_jpy")
    if not lines or not isinstance(total, int) or total != sum(line["monthly_jpy"] for line in lines):
        return {}
    return {
        "schema_version": "1",
        "currency": str(cost_estimate.get("currency") or ""),
        "lines": lines,
        "total_monthly_jpy": total,
        "notes": [str(note) for note in cost_estimate.get("notes", []) if str(note).strip()],
        "api_priced_count": int(cost_estimate.get("api_priced_count") or 0),
    }


def cost_estimate_from_snapshot(snapshot: object) -> dict | None:
    """凍結JSONの価格を、Pricing APIを呼ばずに既存描画器へ渡す。"""
    if not isinstance(snapshot, dict) or snapshot.get("schema_version") != "1":
        return None
    lines_raw = snapshot.get("lines")
    if not isinstance(lines_raw, list) or not lines_raw:
        return None
    lines: list[SimpleNamespace] = []
    for row in lines_raw:
        if (not isinstance(row, dict) or not str(row.get("name", "")).strip()
                or not isinstance(row.get("monthly_jpy"), int)):
            return None
        lines.append(SimpleNamespace(
            name=str(row["name"]), monthly_jpy=row["monthly_jpy"],
            assumption=str(row.get("assumption") or ""), source=str(row.get("source") or ""),
        ))
    total = snapshot.get("total_monthly_jpy")
    if not isinstance(total, int) or total != sum(line.monthly_jpy for line in lines):
        return None
    return {
        "currency": str(snapshot.get("currency") or ""),
        "lines": lines,
        "total_monthly_jpy": total,
        "notes": [str(note) for note in snapshot.get("notes", []) if str(note).strip()],
        "api_priced_count": int(snapshot.get("api_priced_count") or 0),
    }
