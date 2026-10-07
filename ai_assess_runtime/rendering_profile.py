"""PPTX成果物の再現性契約で共有する描画プロファイル定数。"""

# 本文・表・カード・図中ラベルなど、共通ヘッダー／フッター以外の
# 編集可能テキストに適用する下限。
MIN_EDITABLE_BODY_FONT_PT = 12.0

# 全ページ共通のヘッダー／フッターだけに適用する固定サイズ。
COMMON_CHROME_FONT_PT = 10.0


__all__ = ["MIN_EDITABLE_BODY_FONT_PT", "COMMON_CHROME_FONT_PT"]
