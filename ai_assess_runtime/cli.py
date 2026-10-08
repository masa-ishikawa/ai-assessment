"""PPTX専用CLIの引数契約。"""

import argparse
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace


@dataclass(frozen=True, slots=True)
class CliDefaults:
    input_file: Path
    provider: str
    settings: SimpleNamespace | object


def _research_rounds(value: str) -> int:
    try:
        rounds = int(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("1〜8の整数を指定してください。") from error
    if not 1 <= rounds <= 8:
        raise argparse.ArgumentTypeError("1〜8の範囲で指定してください。")
    return rounds


class _DeprecatedApiKeyAction(argparse.Action):
    """旧CLIを壊さず、秘密値を履歴へ残す指定から環境変数へ移行させる。"""

    def __call__(self, parser: argparse.ArgumentParser, namespace: argparse.Namespace,
                 values: str | None, option_string: str | None = None) -> None:
        del parser, option_string
        print(
            "警告: --openai-api-key は非推奨です。OPENAI_API_KEY環境変数を使用してください。",
            file=sys.stderr,
        )
        setattr(namespace, self.dest, values)


def build_argument_parser(defaults: CliDefaults) -> argparse.ArgumentParser:
    """副作用なくparserを組み立て、CLI契約を単体テスト可能にする。"""
    settings = defaults.settings
    parser = argparse.ArgumentParser(
        description="サービス情報（固定Excelまたはテキスト）からAIユースケースアセスメントPPTXを作成します。",
    )
    parser.add_argument(
        "input_file", nargs="?", type=Path, default=defaults.input_file,
        help=("入力ファイル名（.xlsx/.txt）。ファイル名だけならassessment_inputsから読み込みます。"
              f"省略時: assessment_inputs/{defaults.input_file.name}"),
    )
    parser.add_argument(
        "--provider", choices=["oci_responses", "oci", "openai"], default=defaults.provider,
        help="AIプロバイダー（既定値はAI_ASSESS_PROVIDERまたはassessment_config.py）",
    )
    parser.add_argument("--model-id", default=os.getenv("AI_ASSESS_MODEL_ID"),
                        help="使用するモデルID（省略時はプロバイダーの既定モデル）")
    parser.add_argument(
        "--openai-api-key", default=os.getenv("OPENAI_API_KEY"),
        action=_DeprecatedApiKeyAction, help=argparse.SUPPRESS,
    )
    parser.add_argument("--compartment-id", default=os.getenv("OCI_COMPARTMENT_ID", settings.COMPARTMENT_ID),
                        help="OCIコンパートメントOCID（provider=ociの場合のみ使用）")
    parser.add_argument("--endpoint", default=os.getenv("OCI_GENAI_ENDPOINT", settings.GENAI_ENDPOINT),
                        help="OCI Generative AI Inference endpoint")
    parser.add_argument("--profile", default=os.getenv("OCI_PROFILE", settings.OCI_PROFILE),
                        help="OCI設定ファイルのプロファイル名")
    parser.add_argument("--oci-config-file", default=os.getenv("OCI_CONFIG_FILE", settings.OCI_CONFIG_FILE),
                        help="OCI User Principal認証に使う秘密鍵設定ファイル（既定値はassessment_config.py）")
    parser.add_argument("--oci-project-ocid", default=os.getenv("OCI_GENAI_PROJECT_OCID", settings.OCI_GENAI_PROJECT_OCID),
                        help="OCI OpenAI互換Responses APIのGenerative AI Project OCID")
    parser.add_argument("--oci-region", default=os.getenv("OCI_REGION", settings.OCI_REGION),
                        help="OCI OpenAI互換Responses APIのリージョン（例: us-chicago-1）")
    parser.add_argument(
        "--oci-responses-auth-mode", choices=["auto", "user_principal", "resource_principal"],
        default=os.getenv("OCI_RESPONSES_AUTH_MODE", "auto"),
        help="OCI Responses API認証（autoはローカル=User Principal、OCI上=Resource Principal）",
    )
    parser.add_argument("--output", type=Path, help="PPTXの出力パス（省略時: output/pptx/<会社名>御中_<製品名>_AI活用ご提案.pptx）")
    parser.add_argument("--json-only", action="store_true",
                        help="PPTXを生成せず、レビュー用assessment JSONだけを出力する")
    parser.add_argument("--output-json", type=Path,
                        help="--json-onlyの出力JSONパス（省略時: output/json/）")
    parser.add_argument("--from-json", type=Path, help="承認済みassessment JSONからPPTXを生成する")
    parser.add_argument(
        "--revalidate-midterm-targets", action="store_true",
        help=("--from-json --json-onlyでのみ使用。現在の原典照合ルールで中期経営計画の"
              "数値目標を再検証し、安全なv2レビューJSONとして再凍結する"),
    )
    parser.add_argument(
        "--use-case-catalog-json", type=Path,
        help=(
            "過去の15件ユースケースカタログJSONを固定母集団として使う。"
            "過去assessment JSONまたは軽量カタログJSONを受け付け、PoC以降は今回の入力で再作成する。"
            "軽量JSONの明示poc_selectionだけはP1〜P3のID・順序を固定できる"
        ),
    )
    parser.add_argument("--validate-json", type=Path, help="assessment JSONを整合性チェックして終了する")
    parser.add_argument("--strict-json", action="store_true",
                        help="v2の凍結レビューJSONだけを受け入れる（旧v1の移行漏れを検知）")
    parser.add_argument("--include-source-text", action="store_true",
                        help="--json-onlyで入力原文もレビューJSONへ含める（既定ではハッシュのみを保存）")
    parser.add_argument("--verify-manifest", type=Path,
                        help="成果物の.manifest.jsonを検証して終了する")
    parser.add_argument("--verify-review-json", type=Path,
                        help="--verify-manifestと組み合わせ、対応する凍結レビューJSONも照合する")
    parser.add_argument("--extract-input-json", type=Path,
                        help="ISV入力フォームの必須回答だけを前処理JSONとして出力して終了する")
    parser.add_argument("--company-name", help="表紙に表示する会社名（省略時は入力中の「会社名：」を使用）")
    parser.add_argument("--architecture-image", type=Path,
                        help="構成図に使用するPNG/JPEGを指定（省略時は固定の標準構成図を使用）")
    parser.add_argument("--skip-industry-research", action="store_true",
                        help=("DuckDuckGoによる外部定量根拠の調査を省略する。PoC別の外部実績は"
                              "測定設計へ切り替え、会社KPIのAI導入効果目標は計画試算として"
                              "監査ログに区別して残す（中期経営計画は別フラグ）"))
    parser.add_argument(
        "--include-midterm-plan", action=argparse.BooleanOptionalAction, default=True,
        help=("中期経営計画を公開Webで確認し、確認できた資料だけを提案へ反映する（既定: 有効）。"
              "--no-include-midterm-planで無効化"),
    )
    parser.add_argument(
        "--research-max-rounds", type=_research_rounds, default=4, metavar="1..8",
        help="Responses APIが根拠不足を監査して追加検索する最大回数（1〜8、既定値: 4）",
    )
    return parser
