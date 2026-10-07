"""AIアセスメントの生成プロンプトと凍結JSON検証契約。"""

import copy
import hashlib
import json
import re

from ai_assess_runtime.rendering_profile import (
    COMMON_CHROME_FONT_PT,
    MIN_EDITABLE_BODY_FONT_PT,
)
from ai_assess_runtime.poc_contract import (
    _front_text,
    normalized_use_case_label,
    normalize_business_value_model,
    normalize_consulting_front_matter,
    normalize_customer_priority_binding,
    normalize_fixed_poc_selection_binding,
    normalize_multitenant_governance,
    normalize_poc_charters,
    normalize_poc_portfolio,
    normalize_poc_priority_decision,
    normalize_poc_selection_scorecard,
    normalize_poc_start_readiness,
    normalize_technical_proposal,
)
from ai_assess_runtime.quantitative_contract import (
    _ai_product_business_impact_is_consistent,
    displayed_quantitative_effect_issues,
    _llm_quantitative_estimate_is_consistent,
    _poc_measurement_design_is_verifiable,
    materialize_assessment_decision_contract,
    materialize_safe_reproducibility_v2_quantitative_contract,
    materialize_midterm_plan_targets,
    normalize_ai_product_business_impact_candidate,
    normalize_poc_measurement_design,
    sanitize_quantitative_evidence,
    validate_quantitative_evidence,
)
from ai_assess_runtime.review_snapshot import (
    canonical_sha256,
    cost_estimate_from_snapshot,
    without_provenance as _without_provenance,
)
from ai_assess_runtime.source_input import extract_target_services

CUSTOMER_FACING_TONE_GUIDANCE = (
    "顧客向け文章は、提案と共同検討の余地が伝わる穏やかな文体にする。"
    "『〜しません』『〜してはならない』『禁止』『必須です』のような否定・命令の強い言い切りを避け、"
    "『〜を想定しています』『〜が考えられます』『〜を基本とします』『〜を参考情報として扱います』"
    "『〜を確認しながら進めます』など、目的と推奨する進め方を肯定形で示す。"
    "事実・数値・安全上の条件は曖昧にせず、根拠区分や人による確認方法を丁寧に説明する。"
)

def build_prompt(source_text: str) -> str:
    target_services = extract_target_services(source_text)[:3]
    target_services_instruction = (
        "入力で指定された対象サービス（最大3件）: " + "、".join(target_services)
        if target_services else "対象サービスは入力本文から、異なる業務サービスを最大3件抽出してください。"
    )
    return f"""あなたはOracle Cloud Infrastructure（OCI）のAIアーキテクトです。
以下の顧客サービス情報を分析してください。対象サービスが複数ある場合は、異なる業務サービスごとに最大3件まで選び、各サービスに具体的に適したAIユースケースを15件、重複なく提案してください。

必ず次のJSONだけを返してください。Markdownのコードフェンスや説明文は不要です。
{{
  "company_name": "提案先の正式企業名。入力にない場合は空文字",
  "service_name": "サービス名（入力から特定）",
  "business_model_role": "provider / operator / mixed / unknown のいずれか。AI機能を外部顧客へ提供・商品化するISV/SaaSならprovider、導入企業の自社業務のみを扱うならoperator、両方を扱うならmixed。判断できなければunknown",
  "executive_summary": "サービスの現状・AIで目指す事業価値・実現アプローチを含む、社外向けのアセスメントサマリー。2〜3文、160文字以内。",
  "assessment_points": ["事業価値の評価観点", "実現性・運用の評価観点", "データ・統制の評価観点"],
  "business_value": {{
    "headline": "AI実装が事業にもたらす価値を表す短い見出し",
    "summary": "なぜ今AI実装が重要かを、サービスの業務フロー・市場変化・利用者・事業成長と結び付けた140文字以内の説明",
    "impact_areas": [
      {{"area": "価値領域名", "business_rationale": "現行業務・市場環境において、この領域へのAI実装が重要な理由", "ai_enabled": "AIで実現する具体的な判断・業務フロー・利用者行動の変化", "expected_impact": "価値創出の仕組みと、改善するKPIまたは期待する事業効果。根拠のない数値は書かない"}}
    ]
  }},
  "recommendation_message": "本アセスメントを踏まえた、PoCの進め方と意思決定の推奨方針。社外向けの丁寧な文体で100文字以内。",
  "poc_recommendations": [
    {{"use_case_id": "代表サービスの15候補のUC01〜UC15のいずれか", "theme": "use_case_idが参照する候補と一字一句同じAI技術テーマ", "reason": "代表テーマとして示す理由", "first_step": "PoCで最初に検証すること", "architecture_implementation": "業務連携・アプリケーション層、Autonomous AI Database、OCI Generative AIを使った具体的な実現方法。52文字以内。"}}
  ],
  "poc_logic_details": [
    {{"use_case_id": "poc_recommendationsの同じ順序のID", "theme": "poc_recommendationsの同じ順序・同一表記のテーマ", "business_challenge": "現状の業務課題と、PoCで実現・確認したい業務価値。100文字以内。", "implementation_summary": "使用データ、Oracle Database上の主処理、既存側へ返す結果を一文で示す実装方針。100文字以内。", "target_data": "PoCで扱う入力・対象データ、件数・帳票種別などの対象範囲。100文字以内。", "processing_steps": [{{"label": "処理ステップ名", "description": "入力・前処理を、対象データと出力を含めて具体的に説明。120文字以内。"}}, {{"label": "処理ステップ名", "description": "AI処理・判断ロジックを、使用するデータと生成・抽出結果を含めて具体的に説明。120文字以内。"}}, {{"label": "処理ステップ名", "description": "保存・業務返却・人手確認を、データの保存先と例外処理を含めて具体的に説明。120文字以内。"}}], "oci_roles": "業務連携・アプリケーション層、Autonomous AI Database、OCI Generative AIの役割分担。130文字以内。", "oracle_technologies": "使用候補となるOracle Database・OCIの機能名を2〜4件。実在機能だけを書く。", "database_objects": "作成する入力ビュー、特徴量・索引、結果・監査表などの論理オブジェクト。70文字以内。", "business_integration": "既存画面・API・人手確認へどう組み込むか。100文字以内。", "output_interface": "既存側へ返す結果、根拠、スコアと接続方式。80文字以内。", "implementation_boundary": "既存Oracle Database内実行またはAutonomous AI Databaseへの複製・Database Link・APIオフロードを選ぶ判断条件。100文字以内。", "validation_plan": "精度・応答時間・処理単価・人手確認率など、PoCで比較・検証する項目と成功基準。130文字以内。", "design_notes": "セキュリティ、権限、ログ、例外処理、本番展開に向けた留意点。110文字以内。", "control_design": "権限、例外時フォールバック、実行・監査ログの実装。80文字以内。"}}
  ],
  "service_use_case_groups": [
    {{"service_name": "対象となる具体的な業務サービス名", "service_type": "勤怠管理・ERP・店舗発注のような、サービス種別を表す簡潔な日本語（18文字以内）", "use_cases": [
      {{"no": 1, "use_case_id": "代表サービスではnoに対応するUC01〜UC15", "coverage_area": "他候補と重複しない業務価値領域", "use_case": "短いユースケース名", "ai_technology": "生成AI・RAG・文書理解AI のような必要技術（1〜3種類）", "description": "対象データとAIの処理を示す1文。業務効果を示す1文。"}}
    ]}}
  ]
}}

要件:
- {CUSTOMER_FACING_TONE_GUIDANCE}
- company_name は入力から特定できる提案先の正式企業名だけを返す。推測で補わず、特定できなければ空文字にする。
- business_model_role は、対象組織がAI機能を顧客へ提供して商品化する立場か、導入企業として自社業務で使う立場かを入力から判定して返す。MRR、ARPA、継続利用、サポート原価、横展開はprovider/mixedの場合も現状実績として創作せず、PoCで確認する事業KPI候補としてのみ扱う。
- 入力に「中期経営計画の公開資料」が含まれる場合は、その資料の記載を優先して、事業方針・重点施策とAI活用候補のつながりを executive_summary、business_value、PoC候補に反映する。資料に書かれていない数値・施策は創作しない。
- service_name は表紙に表示できる対象サービス名にする。対象サービスが複数ある場合は、入力で最優先・主力・代表と読み取れる製品を1つ選び、「製品名 等」とする。入力に複数あるサービスを単に並べず、「○○社の△△向けサービス群」のような会社名＋総称にも絶対にしない。代表製品を特定できない場合だけ、具体的な業務サービス名を返す。
- {target_services_instruction}
- service_use_case_groups は1〜3件。各グループは異なるサービスに対応し、service_name、service_type、use_cases を必ず含める。service_type は顧客・経営者が一目で業務領域を理解できる簡潔な日本語にする。use_cases は必ず15件、no は1から15まで連番。
- 代表サービス（service_use_case_groupsの先頭）の各候補には、no=1〜15に対応する use_case_id="UC01"〜"UC15" を必ず付ける。poc_recommendations は、その15候補から予測・異常検知、最適化、RAG、文書理解、生成支援など主となるAI技術方式ができるだけ重ならない3件を代表例として選び、use_case_id と theme を対応候補と完全一致させる。poc_logic_details も同じ use_case_id・theme・順序で返す。候補にない別名のPoCや、候補と異なるテーマの詳細は作らない。
- 対象サービスが1件だけの場合は、service_use_case_groups も1件だけにする。入力に明確な複数サービスがある場合のみ、最大3件に分ける。
- 入力に「AIで実現したいこと」「AIアセスメントにおける初期仮説」「要望」「課題」として明記された内容は、必ず各サービスの候補に反映する。ただし、同じ目的の言い換え・対象粒度違い・チャネル違い・前後工程違いを複数の候補に分割しない。最も業務価値が高い1件に統合する。
- 入力に「顧客指定の優先ユースケース」がある場合、その各テーマをservice_use_case_groupsの候補へ必ず含め、優先度「高」のテーマはpoc_recommendationsの3件にも必ず含める。名称を一般論へ置換せず、顧客が示した業務領域・利用シナリオ・必要データを具体的に反映する。
- 15件はフラットで多様なポートフォリオにする。各候補の coverage_area は必ず異なる短い名称とし、同じ業務価値領域を2件以上使わない。対象ユーザー、主データ、AI処理、意思決定または業務成果のうち3つ以上が共通する候補は重複と見なし、1件だけ残す。
- 相互に密接な候補は、個別の業務成果が明確に異なる場合だけ併記できる。併記する場合も、対象ユーザー・データ・意思決定を明確に分ける。
- 明示されたAI希望が少数の領域に偏る場合も、残りの候補はサービスの提供価値、利用者、周辺業務、収益性、リスク・統制、顧客体験、データ活用などから探索し、似た案で件数を満たしてはならない。
- business_value の impact_areas は必ず3件。入力されたサービスの収益性、顧客価値、業務生産性、リスク、意思決定などから、重要度の高い異なる3領域を案件ごとに選ぶ。根拠がない改善率や金額を創作しない。
- 「なぜ今、AI実装が重要か」は一般的なAIの利点を書かない。入力にあるサービス機能、主な利用者、扱うデータ、業務上のボトルネック、顧客への提供価値を必ず結び付ける。summary は事業・業務・競争力の関係を具体的に説明する。各 impact_area は、business_rationale（なぜ重要か）、ai_enabled（何がどのように変わるか）、expected_impact（どのような因果で事業効果・KPIにつながるか）の3層をすべて具体的に書く。
- business_rationale は70文字以内、ai_enabled と expected_impact は各85文字以内にする。文章量を確保しつつ、スライド上で読める長さに収める。
- assessment_points は必ず3件、poc_recommendations は必ず3件。
- assessment_points は、候補選定前の共通評価観点として、順に「事業価値」「実現性・運用」「データ・統制」を案件固有の表現で書く。ここでは優先PoCの固有テーマ名や結論を先出しせず、何を根拠に候補を比較するかが分かる文章にする。
- executive_summary は2〜3文・160文字以内とし、現状、目指す事業価値、OCI AIでの実現アプローチを具体的に書く。
- assessment_points は各48文字以内、recommendation_message は100文字以内、PoCの各項目は各60文字以内。
- recommendation_message は複数PoCの並行実施を前提にせず、顧客判断で対象を選定して段階的に検証する方針で書く。
- poc_recommendations の architecture_implementation は、固定の3サービス構成の中で当該PoCをどう実現するかを52文字以内で具体的に書く。
- 代表3テーマに選んだPoCが、需要予測・売上予測・時系列予測・異常検知・分類・スコアリング等の機械学習テーマである場合、Oracle Machine Learning（OML）の利用を必須とする。該当PoCの architecture_implementation、AI処理・判断ステップ、oci_roles には必ずOML（またはOracle Machine Learning）を明記する。生成AIのみで機械学習テーマを実現する案は出さない。
- poc_logic_details は必ず3件。poc_recommendations と同じ順序・同じテーマに対応させる。各テーマについて、入力データ、Oracleの具体的なAI機能、データベースオブジェクト、3段階の処理ロジック、既存側へ返す結果と接続方式、実装位置、権限・例外・監査を具体的に書く。詳細スライドではこの技術設計を主役とし、業務価値や選定理由を繰り返さない。
- processing_steps は必ず3件とし、入力・前処理、AI処理・判断、データ保存・業務への返却の順に書く。一般論を避け、処理対象、処理内容、出力、例外時の扱いを具体的に書く。
- OCR・帳票読取テーマでは、複数のマルチモーダルLLMまたはモデル候補で抽出精度・応答時間・処理単価を比較し、項目を共通の内部データ形式に正規化してAutonomous AI Databaseへ保存する流れを説明する。帳票からの項目抽出にRAGを主処理として使わない。
- 問い合わせ・検索テーマでは、Autonomous AI DatabaseのVector Search等を使ったRAG、根拠提示、回答の人手確認を具体的に書く。
- 異常検知・予測テーマでは、対象業務のイベント履歴、数量・状態・属性、確定結果を入力にし、OMLで候補・スコアを算出する。テーマにない金額・申請データを汎用的に持ち込まない。
- 優先順位・ランキング・配分・最適化テーマでは、期限、進捗、作業量、処理能力、業務制約を入力にし、OMLのスコアと制約ロジックで順位候補を算出する。生成AIによる抽出・要約を主処理にしない。
- ai_technology は1〜3種類のAI技術を「・」で区切る。生成AI、RAG、文書理解AI、画像認識、音声認識、機械学習、時系列予測、異常検知などから、実現に必要な組み合わせを選ぶ。
- description は50文字以内の2文にする。1文目は32文字以内で対象データとAIの処理を完結させ、2文目は得られる業務効果を具体的に書く。単なる一般論を避け、省略記号は使わない。
- OCIサービス名は表に不要。必要技術のみを ai_technology に書く。
- データベースの提案表記は必ず「Autonomous AI Database」とし、DBCSとは書かない。
- 顧客向け文章では「VM」「仮想マシン」を使用しない。業務画面・既存システムとの接続・APIは「業務連携・アプリケーション層」または「アプリケーション/API」と表現する。

顧客サービス情報:
---
{source_text}
---"""


def extract_json(response_text: str) -> dict:
    cleaned = response_text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.IGNORECASE)
    try:
        result = json.loads(cleaned)
    except json.JSONDecodeError as error:
        raise ValueError("GenAIの応答をJSONとして読み取れませんでした。再実行してください。") from error
    required_keys = {"no", "coverage_area", "use_case", "ai_technology", "description"}
    groups = result.get("service_use_case_groups")
    if not isinstance(groups, list) or not 1 <= len(groups) <= 3:
        raise ValueError("GenAIの応答に1〜3件のサービス別ユースケース一覧がありません。再実行してください。")
    for group in groups:
        cases = group.get("use_cases") if isinstance(group, dict) else None
        if not isinstance(group, dict) or not isinstance(group.get("service_name"), str) or not isinstance(group.get("service_type"), str) or not group["service_type"].strip() or not isinstance(cases, list) or len(cases) != 15:
            raise ValueError("各サービス別ユースケース一覧にはサービス名・サービス種別と15件の候補が必要です。再実行してください。")
        if any(not isinstance(case, dict) or not required_keys <= case.keys() for case in cases):
            raise ValueError("GenAIの応答に必要なユースケース項目がありません。再実行してください。")
        cases.sort(key=lambda case: int(case["no"]))
        if [int(case["no"]) for case in cases] != list(range(1, 16)):
            raise ValueError("ユースケース番号が1から15の連番ではありません。再実行してください。")
        coverage_areas = [re.sub(r"\s+", "", str(case["coverage_area"])).lower() for case in cases]
        if any(not area for area in coverage_areas) or len(set(coverage_areas)) != 15:
            raise ValueError("15件のユースケースの業務価値領域が重複しています。再実行してください。")
    # 既存のPoC生成ロジックとの互換性のため、代表サービスの候補を保持する。
    # LLMの表記ゆれを許さず、母集団番号からIDを決定する。2つ目以降のサービスも
    # 一覧としては識別できるが、優先PoCは代表サービスのUCxxだけを参照する。
    for group_index, group in enumerate(groups, 1):
        for case in group["use_cases"]:
            number = int(case["no"])
            case["use_case_id"] = (f"UC{number:02d}" if group_index == 1
                                   else f"S{group_index:02d}-UC{number:02d}")
    result["use_cases"] = groups[0]["use_cases"]
    if not isinstance(result.get("executive_summary"), str):
        raise ValueError("GenAIの応答にアセスメント要約がありません。再実行してください。")
    if not isinstance(result.get("company_name"), str):
        raise ValueError("GenAIの応答に提案先企業名がありません。再実行してください。")
    role = str(result.get("business_model_role") or "").strip().lower()
    role_aliases = {
        "provider": "provider", "isv": "provider", "saas": "provider", "software_provider": "provider",
        "operator": "operator", "enterprise": "operator", "customer": "operator",
        "mixed": "mixed", "unknown": "unknown",
        "提供者": "provider", "サービス提供者": "provider", "事業者": "provider",
        "導入企業": "operator", "事業会社": "operator", "運用者": "operator",
        "両方": "mixed", "不明": "unknown",
    }
    # 立場が読み取れない場合は提供者と推測せず、unknownを明示する。
    result["business_model_role"] = role_aliases.get(role, "unknown")
    if not isinstance(result.get("recommendation_message"), str):
        raise ValueError("GenAIの応答に推奨方針がありません。再実行してください。")
    if not isinstance(result.get("assessment_points"), list) or len(result["assessment_points"]) != 3:
        raise ValueError("GenAIの応答に3件のアセスメント着眼点がありません。再実行してください。")
    business_value = result.get("business_value")
    if not isinstance(business_value, dict) or not isinstance(business_value.get("headline"), str) or not isinstance(business_value.get("summary"), str):
        raise ValueError("GenAIの応答にAI実装のビジネス価値がありません。再実行してください。")
    impact_areas = business_value.get("impact_areas")
    if not isinstance(impact_areas, list) or len(impact_areas) != 3 or any(
        not isinstance(item, dict) or not {"area", "business_rationale", "ai_enabled", "expected_impact"} <= item.keys() for item in impact_areas
    ):
        raise ValueError("GenAIの応答に3件のビジネス効果領域がありません。再実行してください。")
    recommendations = result.get("poc_recommendations")
    required_recommendation_keys = {"theme", "reason", "first_step"}
    if not isinstance(recommendations, list) or len(recommendations) != 3 or any(
        not isinstance(item, dict) or not required_recommendation_keys <= item.keys() for item in recommendations
    ):
        raise ValueError("GenAIの応答に3件のPoC推奨テーマがありません。再実行してください。")
    return result


ASSESSMENT_JSON_FORMAT = "ai-assess/assessment-v1"


ASSESSMENT_JSON_FROZEN_FORMAT = "ai-assess/assessment-v2"


REPRODUCIBILITY_CONTRACT_V1_FORMAT = "ai-assess/reproducibility-v1"


REPRODUCIBILITY_CONTRACT_V2_FORMAT = "ai-assess/reproducibility-v2"


# 新規生成は入力レコード全体と描画プロファイルまで固定するv2を使う。
# v1は既存の承認済みJSONを再生するため、検証側だけで読み取り互換を保つ。
REPRODUCIBILITY_CONTRACT_FORMAT = REPRODUCIBILITY_CONTRACT_V2_FORMAT


_REPRODUCIBILITY_CONTRACT_FORMATS = {
    REPRODUCIBILITY_CONTRACT_V1_FORMAT,
    REPRODUCIBILITY_CONTRACT_V2_FORMAT,
}


_STANDARD_ASSESSMENT_CONTRACTS: dict[str, type] = {
    "executive_summary": str,
    "assessment_points": list,
    "business_value": dict,
    "consulting_front_matter": dict,
    "business_value_model": dict,
    "poc_portfolio": dict,
    "poc_logic_details": list,
    "poc_charters": dict,
    "poc_selection_scorecard": dict,
    "poc_priority_decision": dict,
    "poc_start_readiness": dict,
    "multitenant_governance": dict,
    "technical_proposal": dict,
    "poc_measurement_design": dict,
}


def _reused_catalog_provenance_issues(input_record: dict, assessment: dict) -> list[str]:
    """任意の固定カタログ来歴がある場合だけ、最終15件との同一性を確認する。"""
    preprocessing = input_record.get("preprocessing")
    if not isinstance(preprocessing, dict) or "use_case_catalog_reuse" not in preprocessing:
        return []
    reuse = preprocessing.get("use_case_catalog_reuse")
    if not isinstance(reuse, dict):
        return ["input.preprocessing.use_case_catalog_reuseが不正です。"]
    errors: list[str] = []
    if reuse.get("format") != "ai-assess/use-case-catalog-reuse/v1":
        errors.append("input.preprocessing.use_case_catalog_reuse.formatが不正です。")
    if reuse.get("schema_version") != "1":
        errors.append("input.preprocessing.use_case_catalog_reuse.schema_versionが不正です。")
    if not str(reuse.get("source_file") or "").strip():
        errors.append("input.preprocessing.use_case_catalog_reuse.source_fileが必要です。")
    if not re.fullmatch(r"[0-9a-f]{64}", str(reuse.get("source_sha256") or "")):
        errors.append("input.preprocessing.use_case_catalog_reuse.source_sha256が不正です。")
    if not str(reuse.get("source_format") or "").strip():
        errors.append("input.preprocessing.use_case_catalog_reuse.source_formatが必要です。")
    groups = assessment.get("service_use_case_groups")
    if not isinstance(groups, list) or not groups:
        errors.append("固定カタログ再利用時はassessment.service_use_case_groupsが必要です。")
        return errors
    if reuse.get("catalog_sha256") != canonical_sha256(groups):
        errors.append("固定カタログのcatalog_sha256とassessment.service_use_case_groupsが一致しません。")
    service_count = reuse.get("service_count")
    if (not isinstance(service_count, int) or isinstance(service_count, bool)
            or service_count != len(groups) or not 1 <= service_count <= 3):
        errors.append("input.preprocessing.use_case_catalog_reuse.service_countが不正です。")
    service_names = reuse.get("service_names")
    actual_names = [
        str(group.get("service_name") or "").strip()
        for group in groups if isinstance(group, dict)
    ]
    if (not isinstance(service_names, list) or service_names != actual_names
            or len(actual_names) != len(groups) or any(not name for name in actual_names)):
        errors.append("input.preprocessing.use_case_catalog_reuse.service_namesが不正です。")
    primary_cases = groups[0].get("use_cases") if isinstance(groups[0], dict) else None
    if primary_cases != assessment.get("use_cases"):
        errors.append("固定カタログの代表15件とassessment.use_casesが一致しません。")
    selection_keys = {"poc_selection", "poc_selection_sha256", "poc_selection_count"}
    if selection_keys & set(reuse):
        source_selection = reuse.get("poc_selection")
        selection_hash = reuse.get("poc_selection_sha256")
        selection_count = reuse.get("poc_selection_count")
        if source_selection is None:
            if selection_hash != "" or selection_count != 0:
                errors.append("固定カタログのpoc_selection来歴が不正です。")
        elif not isinstance(source_selection, dict):
            errors.append("固定カタログのpoc_selection来歴が不正です。")
        else:
            if selection_hash != canonical_sha256(source_selection):
                errors.append("固定カタログのpoc_selection_sha256が一致しません。")
            selected_items = source_selection.get("items")
            if (not isinstance(selected_items, list) or len(selected_items) != 3
                    or selection_count != 3):
                errors.append("固定カタログのpoc_selectionはP1〜P3の3件である必要があります。")
            else:
                primary_by_id = {
                    str(case.get("use_case_id") or ""): case
                    for case in primary_cases if isinstance(case, dict)
                } if isinstance(primary_cases, list) else {}
                for index, item in enumerate(selected_items, 1):
                    if not isinstance(item, dict):
                        errors.append("固定カタログのpoc_selection項目が不正です。")
                        break
                    case = primary_by_id.get(_front_text(item.get("use_case_id"), 8))
                    if (
                        _front_text(item.get("priority"), 10) != f"P{index}"
                        or case is None
                        or _front_text(item.get("use_case_no"), 3) != str(case.get("no"))
                        or normalized_use_case_label(item.get("theme"))
                        != normalized_use_case_label(case.get("use_case"))
                        or isinstance(item.get("detail_slide_order"), bool)
                        or item.get("detail_slide_order") != index
                    ):
                        errors.append("固定カタログのpoc_selection項目が15件一覧と一致しません。")
                        break
                fixed_binding = normalize_fixed_poc_selection_binding(
                    assessment.get("fixed_poc_selection_binding"), assessment,
                )
                if not fixed_binding:
                    errors.append("固定カタログのpoc_selectionがassessmentへ反映されていません。")
                elif fixed_binding.get("items") != selected_items:
                    errors.append("固定カタログのpoc_selectionとassessment固定選定が一致しません。")
                portfolio = assessment.get("poc_portfolio")
                portfolio_items = portfolio.get("items") if isinstance(portfolio, dict) else None
                if (
                    not isinstance(portfolio_items, list) or len(portfolio_items) != 3
                    or any(
                        not isinstance(actual, dict)
                        or _front_text(actual.get("priority"), 10) != _front_text(expected.get("priority"), 10)
                        or _front_text(actual.get("use_case_id"), 8) != _front_text(expected.get("use_case_id"), 8)
                        or _front_text(actual.get("use_case_no"), 3) != _front_text(expected.get("use_case_no"), 3)
                        or normalized_use_case_label(actual.get("theme"))
                        != normalized_use_case_label(expected.get("theme"))
                        for actual, expected in zip(portfolio_items or [], selected_items)
                    )
                ):
                    errors.append("固定カタログのpoc_selectionとpoc_portfolioが一致しません。")
    return errors


def validate_reproducibility_contract(payload: dict) -> list[str]:
    """v2の入力・分析・リサーチ・費用スナップショットの同一性を検証する。"""
    errors: list[str] = []
    input_record = payload.get("input")
    assessment = payload.get("assessment")
    research = payload.get("research")
    rendering = payload.get("rendering")
    provenance = payload.get("provenance")
    if not all(isinstance(value, dict) for value in (input_record, assessment, research, rendering, provenance)):
        return ["assessment-v2にはinput・assessment・research・rendering・provenanceが必要です。"]
    errors.extend(_reused_catalog_provenance_issues(input_record, assessment))
    contract_format = provenance.get("format")
    if contract_format not in _REPRODUCIBILITY_CONTRACT_FORMATS:
        errors.append("provenance.formatが不正です。")
        contract_format = REPRODUCIBILITY_CONTRACT_V1_FORMAT
    source_hash = input_record.get("source_text_sha256")
    if not isinstance(source_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", source_hash):
        errors.append("input.source_text_sha256が必要です。")
    elif input_record.get("source_text_included") is True and hashlib.sha256(
            str(input_record.get("source_text") or "").encode("utf-8")).hexdigest() != source_hash:
        errors.append("input.source_text_sha256がsource_textと一致しません。")
    generator_hash = provenance.get("generator_sha256")
    if not isinstance(generator_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", generator_hash):
        errors.append("provenance.generator_sha256は64桁のSHA-256である必要があります。")
    input_hash = canonical_sha256(input_record)
    expected = {
        "source_text_sha256": source_hash,
        "assessment_sha256": canonical_sha256(assessment),
        "research_content_sha256": canonical_sha256(_without_provenance(research)),
        "generator_sha256": generator_hash,
        "cost_estimate_sha256": canonical_sha256(rendering.get("cost_estimate") or {}) if rendering.get("cost_estimate") else "",
    }
    if contract_format == REPRODUCIBILITY_CONTRACT_V2_FORMAT:
        expected["input_record_sha256"] = input_hash
        expected["rendering_sha256"] = canonical_sha256(rendering)
    if not isinstance(provenance.get("assessment_run_id"), str) or not provenance.get("assessment_run_id"):
        errors.append("provenance.assessment_run_idが必要です。")
    for field, value in expected.items():
        if provenance.get(field) != value:
            errors.append(f"provenance.{field}が現在の内容と一致しません。")
    run_input_hash = input_hash if contract_format == REPRODUCIBILITY_CONTRACT_V2_FORMAT else source_hash
    run_components = [
        run_input_hash,
        expected["assessment_sha256"],
        expected["research_content_sha256"],
        expected["generator_sha256"],
    ]
    if contract_format == REPRODUCIBILITY_CONTRACT_V2_FORMAT:
        run_components.append(expected["rendering_sha256"])
    expected_run_id = hashlib.sha256(":".join(str(item) for item in run_components).encode("utf-8")).hexdigest()[:24]
    if provenance.get("assessment_run_id") != expected_run_id:
        errors.append("provenance.assessment_run_idが現在の内容と一致しません。")
    research_provenance = research.get("provenance")
    if not isinstance(research_provenance, dict):
        errors.append("research.provenanceが必要です。")
    else:
        if (contract_format == REPRODUCIBILITY_CONTRACT_V2_FORMAT
                and research_provenance.get("format") != contract_format):
            errors.append("research.provenance.formatがレビューJSONと一致しません。")
        research_fields = [
            "assessment_run_id", "source_text_sha256", "assessment_sha256",
            "research_content_sha256",
        ]
        if contract_format == REPRODUCIBILITY_CONTRACT_V2_FORMAT:
            research_fields.extend(("input_record_sha256", "rendering_sha256"))
        for field in research_fields:
            if research_provenance.get(field) != provenance.get(field):
                errors.append(f"research.provenance.{field}がレビューJSONと一致しません。")
    if contract_format == REPRODUCIBILITY_CONTRACT_V2_FORMAT:
        architecture_asset = rendering.get("architecture_asset")
        if not isinstance(architecture_asset, dict):
            errors.append("rendering.architecture_assetが必要です。")
        else:
            if architecture_asset.get("schema_version") != "1":
                errors.append("rendering.architecture_asset.schema_versionが不正です。")
            if architecture_asset.get("source") not in {"default", "cli"}:
                errors.append("rendering.architecture_asset.sourceが不正です。")
            if not str(architecture_asset.get("path") or "").strip():
                errors.append("rendering.architecture_asset.pathが必要です。")
            if not re.fullmatch(r"[0-9a-f]{64}", str(architecture_asset.get("sha256") or "")):
                errors.append("rendering.architecture_asset.sha256が不正です。")
            if (not isinstance(architecture_asset.get("bytes"), int)
                    or isinstance(architecture_asset.get("bytes"), bool)
                    or architecture_asset.get("bytes", 0) <= 0):
                errors.append("rendering.architecture_asset.bytesが不正です。")
        render_profile = rendering.get("render_profile")
        if not isinstance(render_profile, dict):
            errors.append("rendering.render_profileが必要です。")
        else:
            required_profile = {
                "schema_version": "1",
                "artifact_format": "pptx",
                "page_layout": "widescreen-16:9",
                "design": "default",
                "min_editable_font_pt": MIN_EDITABLE_BODY_FONT_PT,
                "common_header_footer_font_pt": COMMON_CHROME_FONT_PT,
            }
            for field, expected_value in required_profile.items():
                if render_profile.get(field) != expected_value:
                    errors.append(f"rendering.render_profile.{field}が不正です。")
    if rendering.get("cost_estimate") and cost_estimate_from_snapshot(rendering.get("cost_estimate")) is None:
        errors.append("rendering.cost_estimateの形式または合計値が不正です。")
    return errors


def migrate_legacy_payload_for_review(payload: dict) -> dict:
    """v1 JSONを破壊せず、v2レビュー用の安全な派生データへ移行する。"""
    migrated = copy.deepcopy(payload)
    assessment = migrated.get("assessment")
    research = migrated.get("research")
    if not isinstance(assessment, dict):
        return migrated
    if not isinstance(research, dict):
        research = {}
        migrated["research"] = research
    assessment.setdefault("business_model_role", "unknown")
    materialize_midterm_plan_targets(assessment, research)
    materialize_assessment_decision_contract(assessment)
    materialize_safe_reproducibility_v2_quantitative_contract(
        assessment, research, migrate_legacy=True,
    )
    return migrated


def validate_assessment_payload(payload: object, *, strict: bool = False,
                                allow_legacy_hypotheses: bool = False) -> list[str]:
    """Validate the reviewable JSON contract before approval or PPTX rendering."""
    if not isinstance(payload, dict):
        return ["JSONの最上位がオブジェクトではありません。"]
    errors: list[str] = []
    payload_format = payload.get("format")
    if payload_format not in {ASSESSMENT_JSON_FORMAT, ASSESSMENT_JSON_FROZEN_FORMAT}:
        errors.append(f"formatは{ASSESSMENT_JSON_FORMAT}または{ASSESSMENT_JSON_FROZEN_FORMAT}である必要があります。")
    if strict and payload_format != ASSESSMENT_JSON_FROZEN_FORMAT:
        errors.append(f"--strict-jsonでは{ASSESSMENT_JSON_FROZEN_FORMAT}が必要です。")
    provenance = payload.get("provenance")
    is_reproducibility_v2 = bool(
        payload_format == ASSESSMENT_JSON_FROZEN_FORMAT
        and isinstance(provenance, dict)
        and provenance.get("format") == REPRODUCIBILITY_CONTRACT_V2_FORMAT
    )
    if payload_format == ASSESSMENT_JSON_FROZEN_FORMAT:
        errors.extend(validate_reproducibility_contract(payload))
    assessment = payload.get("assessment")
    if not isinstance(assessment, dict):
        return errors + ["assessmentオブジェクトがありません。"]
    for field in ("company_name", "service_name"):
        if not str(assessment.get(field, "")).strip():
            errors.append(f"assessment.{field}がありません。")
    for field in ("use_cases", "poc_recommendations"):
        if not isinstance(assessment.get(field), list) or not assessment[field]:
            errors.append(f"assessment.{field}は空でない配列である必要があります。")
    if strict and payload_format == ASSESSMENT_JSON_FROZEN_FORMAT:
        for field, expected_type in _STANDARD_ASSESSMENT_CONTRACTS.items():
            value = assessment.get(field)
            if not isinstance(value, expected_type) or not value:
                errors.append(
                    f"strict検証ではassessment.{field}に現行標準デッキの凍結契約が必要です。"
                )
        use_cases = assessment.get("use_cases")
        recommendations = assessment.get("poc_recommendations")
        if isinstance(use_cases, list) and len(use_cases) != 15:
            errors.append("strict検証ではassessment.use_casesを15件固定する必要があります。")
        if isinstance(recommendations, list) and len(recommendations) != 3:
            errors.append("strict検証ではassessment.poc_recommendationsを3件固定する必要があります。")
    front_matter = assessment.get("consulting_front_matter")
    source_ids = {"I1"}
    research = payload.get("research")
    if isinstance(research, dict):
        midterm = research.get("midterm_plan")
        if isinstance(midterm, dict) and midterm.get("status") == "found":
            source_ids.add("M1")
        for index, source in enumerate(research.get("industry_sources", []) if isinstance(research.get("industry_sources"), list) else [], 1):
            if isinstance(source, dict) and str(source.get("title", "")).strip():
                source_ids.add(str(source.get("id") or f"R{index}"))
    normalized_front_matter = {}
    customer_priority_binding = assessment.get("customer_priority_binding")
    if customer_priority_binding is not None and not normalize_customer_priority_binding(
        customer_priority_binding, assessment,
    ):
        errors.append(
            "assessment.customer_priority_bindingの出所・P1〜P3・候補ID対応に問題があります。"
        )
    fixed_poc_selection_binding = assessment.get("fixed_poc_selection_binding")
    if fixed_poc_selection_binding is not None and not normalize_fixed_poc_selection_binding(
        fixed_poc_selection_binding, assessment,
    ):
        errors.append(
            "assessment.fixed_poc_selection_bindingの出所・P1〜P3・候補ID対応に問題があります。"
        )
    if customer_priority_binding is not None and fixed_poc_selection_binding is not None:
        errors.append("customer_priority_bindingとfixed_poc_selection_bindingは同時に指定できません。")
    if front_matter is not None:
        normalized_front_matter = normalize_consulting_front_matter(front_matter, assessment, source_ids)
        if not normalized_front_matter:
            errors.append("assessment.consulting_front_matterの件数・根拠区分・優先順位の整合性に問題があります。")
    poc_portfolio = assessment.get("poc_portfolio")
    if poc_portfolio is not None:
        normalized_portfolio = normalize_poc_portfolio(
            poc_portfolio, assessment, normalized_front_matter or None,
        )
        if not normalized_portfolio:
            errors.append("assessment.poc_portfolioのuse_case_id・優先順位・候補表の対応に問題があります。")
        else:
            recommendations = assessment.get("poc_recommendations")
            expected_items = normalized_portfolio["items"]
            if (not isinstance(recommendations, list) or len(recommendations) != 3
                    or any(not isinstance(item, dict) for item in recommendations)
                    or any(
                        _front_text(actual.get("priority"), 10) != expected["priority"]
                        or _front_text(actual.get("use_case_id"), 8) != expected["use_case_id"]
                        or _front_text(actual.get("use_case_no"), 3) != expected["use_case_no"]
                        or normalized_use_case_label(actual.get("theme")) != normalized_use_case_label(expected["theme"])
                        for actual, expected in zip(recommendations, expected_items)
                    )):
                errors.append("assessment.poc_recommendationsはpoc_portfolioと同じuse_case_id・優先順位・テーマである必要があります。")
            if normalized_front_matter:
                selection_logic = normalized_front_matter["use_case_prioritization"]["selection_logic"]
                if any(
                    normalized_use_case_label(item.get("theme")) != normalized_use_case_label(expected["theme"])
                    for item, expected in zip(selection_logic, expected_items)
                ):
                    errors.append("consulting_front_matterのPoC選定テーマはpoc_portfolioと一致する必要があります。")
            details = assessment.get("poc_logic_details")
            if isinstance(details, list) and details and (
                len(details) != 3
                or any(
                    not isinstance(detail, dict)
                    or _front_text(detail.get("use_case_id"), 8) != expected["use_case_id"]
                    or normalized_use_case_label(detail.get("theme")) != normalized_use_case_label(expected["theme"])
                    for detail, expected in zip(details, expected_items)
                )
            ):
                errors.append("poc_logic_detailsはpoc_portfolioと同じuse_case_id・テーマで対応する必要があります。")
    poc_charters = assessment.get("poc_charters")
    if poc_charters is not None and not normalize_poc_charters(
        poc_charters, assessment, normalized_front_matter or None,
    ):
        errors.append(
            "assessment.poc_chartersのP1〜P3対応・判断時点・データ条件・評価・責任・商品化判断の形式に問題があります。"
        )
    poc_selection_scorecard = assessment.get("poc_selection_scorecard")
    if poc_selection_scorecard is not None and not normalize_poc_selection_scorecard(
        poc_selection_scorecard, assessment, normalized_front_matter or None,
    ):
        errors.append(
            "assessment.poc_selection_scorecardのP1〜P3対応・評価仮説・開始ゲートの形式に問題があります。"
        )
    poc_priority_decision = assessment.get("poc_priority_decision")
    if poc_priority_decision is not None and not normalize_poc_priority_decision(
        poc_priority_decision, assessment, normalized_front_matter or None,
    ):
        errors.append(
            "assessment.poc_priority_decisionの重み・評価充足率・P1〜P3順位が候補評価と一致しません。"
        )
    poc_start_readiness = assessment.get("poc_start_readiness")
    if poc_start_readiness is not None and not normalize_poc_start_readiness(
        poc_start_readiness, assessment, normalized_front_matter or None,
    ):
        errors.append(
            "assessment.poc_start_readinessの開始状態・未充足ゲート・責任者がPoCチャーターと一致しません。"
        )
    multitenant_governance = assessment.get("multitenant_governance")
    if multitenant_governance is not None and not normalize_multitenant_governance(multitenant_governance):
        errors.append(
            "assessment.multitenant_governanceの境界・RAG文書・データライフサイクル・回答統制・モデル運用の形式に問題があります。"
        )
    technical_proposal = assessment.get("technical_proposal")
    if technical_proposal is not None and not normalize_technical_proposal(technical_proposal, assessment):
        errors.append("assessment.technical_proposalの状態区分・P1参照・技術提案データに問題があります。")
    business_value_model = assessment.get("business_value_model")
    if business_value_model is not None and not normalize_business_value_model(business_value_model):
        errors.append("assessment.business_value_modelの二層価値・商品化仮説・PoC接続の形式に問題があります。")
    ai_product_business_impact = assessment.get("ai_product_business_impact")
    if ai_product_business_impact is not None:
        normalized_business_impact = normalize_ai_product_business_impact_candidate(
            ai_product_business_impact, assessment,
        )
        if not normalized_business_impact:
            errors.append(
                "assessment.ai_product_business_impactは、評価対象企業に閉じた売上・利益・"
                "継続収益/事業継続性の3事業KPIである必要があります。"
            )
        elif strict and not _ai_product_business_impact_is_consistent(
            normalized_business_impact, research, assessment,
        ):
            errors.append(
                "AI製品事業インパクトの計画試算は、表示する会社KPI・目標レンジ・"
                "算定式・モデル・算定前提・生成理由・外部根拠件数を、"
                "evidence_mode=llm_estimateの独立監査ログと一致させる必要があります。"
            )
    poc_measurement_design = assessment.get("poc_measurement_design")
    if poc_measurement_design is not None:
        normalized_measurement_design = normalize_poc_measurement_design(poc_measurement_design, assessment)
        quantitative_analysis = (
            research.get("quantitative_analysis") if isinstance(research, dict) else None
        )
        declared_display_mode = (
            str(quantitative_analysis.get("final_display_mode") or "")
            if isinstance(quantitative_analysis, dict) else ""
        )
        if not normalized_measurement_design:
            errors.append(
                "assessment.poc_measurement_designは、未検証KPI設計、旧PoC目標、または原典照合済みの3件の定量効果である必要があります。"
            )
        elif declared_display_mode == "llm_estimate" \
                and normalized_measurement_design.get("status") != "decision_thresholds":
            errors.append(
                "research.quantitative_analysisはLLM定量効果の表示を宣言していますが、"
                "assessment.poc_measurement_designがdecision_thresholdsではありません。"
            )
        elif declared_display_mode == "external_verified" \
                and normalized_measurement_design.get("status") != "external_verified":
            errors.append(
                "research.quantitative_analysisは外部定量根拠の表示を宣言していますが、"
                "assessment.poc_measurement_designがexternal_verifiedではありません。"
            )
        elif normalized_measurement_design.get("status") == "external_verified":
            if not _poc_measurement_design_is_verifiable(normalized_measurement_design, research, assessment):
                errors.append(
                    "external_verifiedの定量効果3件は、承認済み原典・URL・数値抜粋・監査台帳と一対一で一致する必要があります。"
                )
        elif normalized_measurement_design.get("status") == "decision_thresholds":
            measurement_audit = (
                research.get("poc_measurement_design_audit")
                if isinstance(research, dict) else None
            )
            if strict and is_reproducibility_v2:
                errors.append(
                    "reproducibility-v2ではLLM独自試算のdecision_thresholdsを表示できません。"
                    "原文取得済みの公開定量根拠またはpre_poc測定設計を使用してください。"
                )
            elif isinstance(measurement_audit, dict) and measurement_audit.get("generation_source") == "llm_estimate":
                if not _llm_quantitative_estimate_is_consistent(
                    normalized_measurement_design, research, assessment,
                ):
                    errors.append(
                        "LLM定量効果は、表示する3件の数値・estimate_id・モデル・最終採用モードを監査ログと一致させる必要があります。"
                    )
            elif strict:
                errors.append(
                    "strict検証のdecision_thresholdsはgeneration_source=llm_estimateと、"
                    "表示値・P1〜P3対応・モデルを固定した監査ログが必要です。"
                )
            elif not (
                isinstance(measurement_audit, dict)
                and measurement_audit.get("status") == "decision_thresholds"
                and measurement_audit.get("basis_type") == "decision_threshold"
                and str(measurement_audit.get("reason") or "").strip()
            ):
                errors.append(
                    "decision_thresholdsにはresearch.poc_measurement_design_auditで、顧客実績・公開効果ではないPoC判定基準である根拠を記録する必要があります。"
                )
    errors.extend(validate_quantitative_evidence(
        payload, allow_legacy_hypotheses=allow_legacy_hypotheses,
    ))
    if strict and is_reproducibility_v2:
        display_contract = research.get("quantitative_display_contract") \
            if isinstance(research, dict) else None
        if not isinstance(display_contract, dict):
            errors.append(
                "reproducibility-v2にはresearch.quantitative_display_contractが必要です。"
            )
        quantitative_analysis = research.get("quantitative_analysis") \
            if isinstance(research, dict) else None
        if isinstance(quantitative_analysis, dict):
            if quantitative_analysis.get("final_display_mode") not in {
                    "external_verified", "measurement_design_only"}:
                errors.append(
                    "reproducibility-v2の定量表示はexternal_verifiedまたは"
                    "measurement_design_onlyに限定されます。"
                )
            if (quantitative_analysis.get("final_display_mode") == "measurement_design_only"
                    and quantitative_analysis.get("unsupported_claims_omitted") is not True):
                errors.append(
                    "measurement_design_onlyではunsupported_claims_omitted=trueが必要です。"
                )
        else:
            errors.append("reproducibility-v2にはresearch.quantitative_analysisが必要です。")
        llm_estimate = research.get("llm_quantitative_estimate") \
            if isinstance(research, dict) else None
        if isinstance(llm_estimate, dict) and llm_estimate.get("status") == "generated":
            errors.append(
                "reproducibility-v2には表示用のLLM独自定量試算を保持できません。"
            )
        errors.extend(displayed_quantitative_effect_issues(assessment, research))
    evidence = assessment.get("executive_evidence")
    if isinstance(evidence, dict) and evidence.get("benchmarks"):
        targets, premises = evidence.get("management_targets"), evidence.get("strategic_premises")
        if (not isinstance(targets, list) or not targets) and (not isinstance(premises, list) or len(premises) != 3):
            errors.append("中計目標がない場合はstrategic_premisesを3件含める必要があります。")
        if not isinstance(evidence.get("strategic_logic"), list) or len(evidence["strategic_logic"]) != 3:
            errors.append("executive_evidence.strategic_logicを3件含める必要があります。")
    return errors
