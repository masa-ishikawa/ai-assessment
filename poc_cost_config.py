"""PoC概算費用の前提値。

金額はリスト価格ベースの概算値です。SKU（part_number）と月間数量を設定すると、
List Pricing APIから取得した単価で monthly_quantity × unit_price を自動計算します。
未設定の項目は fallback_monthly_jpy を使います。
"""

CURRENCY = "JPY"
PRICE_API_URL = "https://apexapps.oracle.com/pls/apex/cetools/api/v1/products/"

# 1か月 = 744時間（31日 × 24時間）で計算。
MONTHLY_HOURS = 744

COST_ITEMS = [
    {
        "name": "Compute",
        "part_number": None,
        "monthly_quantity": None,
        "fallback_monthly_jpy": 7808,
        "assumption": "1台（VM.Standard.E5.Flex、1 OCPU、16GBメモリ、100GB SSD）、744時間稼働",
    },
    {
        "name": "Autonomous AI Database 26ai",
        "part_number": None,
        "monthly_quantity": None,
        "fallback_monthly_jpy": 80753,
        "assumption": "Autonomous AI Database、2 ECPU、100GB、744時間稼働、バックアップを含む前提",
    },
    {
        "name": "OCI Generative AI（生成モデル）",
        "part_number": None,
        "monthly_quantity": None,
        "fallback_monthly_jpy": 4262,
        "assumption": "Grok 4.3、月間10,000リクエスト（入力100 token、出力500 token/回）を想定",
    },
    {
        "name": "OCI Generative AI（埋込モデル）",
        "part_number": None,
        "monthly_quantity": None,
        "fallback_monthly_jpy": 31,
        "assumption": "問い合わせ文のEmbeddingを対象。既存文書の初回Embeddingは別途スポット費用",
    },
]

NOTES = [
    "本表はリスト価格ベースの概算であり、税抜・月額です。割引契約、Universal Credits、無料枠、段階課金およびデータ転送料金は含みません。",
    "ComputeおよびAutonomous AI Databaseは24時間・31日（744時間）稼働を前提とします。夜間・休日に停止する場合、該当する従量課金は削減できます。",
    "正確なサービスSKUが決まったら、各項目のpart_numberとmonthly_quantityを設定し、OCI List Pricing APIの単価で再計算します。",
]
