"""顧客向け成果物に表示する外部資料名の表記を整える。"""

from __future__ import annotations

import re
from typing import Any


_IR_INFORMATION_LABEL_PATTERN = re.compile(
    r"(?<![0-9A-Za-z０-９Ａ-Ｚａ-ｚ])"
    r"[iIｉＩ][rRｒＲ](?=\s*情報)"
)


def normalize_public_source_title(value: object) -> str:
    """外部資料名の既知の表記ゆれを顧客向け表記へ整える。

    検索結果のタイトルは ``Ir情報`` のように頭字語の大文字・小文字が
    原文と異なる場合がある。URLや本文は変更せず、資料名中のIR表記だけを
    決定的に正規化する。
    """
    return _IR_INFORMATION_LABEL_PATTERN.sub("IR", str(value or ""))


def normalize_ir_information_labels(value: Any) -> Any:
    """JSON互換構造中の ``Ir情報`` 表記を非破壊で正規化する。"""
    if isinstance(value, str):
        return normalize_public_source_title(value)
    if isinstance(value, dict):
        return {
            key: normalize_ir_information_labels(item)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [normalize_ir_information_labels(item) for item in value]
    if isinstance(value, tuple):
        return tuple(normalize_ir_information_labels(item) for item in value)
    return value
