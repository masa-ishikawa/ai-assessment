"""ローカル設定の例。assessment_config.py にコピーして設定する。"""

# 通常の oci_responses 実行で記入する3項目。認証情報は別途 ~/.oci/config に設定する。
GENAI_MODEL_ID = ""
OCI_REGION = ""
OCI_GENAI_PROJECT_OCID = ""

# 以下は既定値を変更するときだけ編集する。
AI_PROVIDER = "oci_responses"
# OPENAI_MODEL は AI_PROVIDER="openai" の場合に使う。
OPENAI_MODEL = "gpt-5.6-terra"
# COMPARTMENT_ID と GENAI_ENDPOINT は AI_PROVIDER="oci" の場合に使う。
COMPARTMENT_ID = ""
GENAI_ENDPOINT = ""
# OCI_CONFIG_FILE と OCI_PROFILE は OCI 認証設定の場所・プロファイルを変える場合に使う。
OCI_CONFIG_FILE = "~/.oci/config"
OCI_PROFILE = "DEFAULT"
REASONING_EFFORT = "medium"
