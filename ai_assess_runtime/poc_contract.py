"""PoC選定・事業価値・技術提案を構成する決定的なドメイン契約。"""

import copy
from difflib import SequenceMatcher
import re
import unicodedata

from ai_assess_runtime.pptx_canvas import _plain_text


def is_ml_theme(theme: str) -> bool:
    """優先PoCでOMLを必須とする機械学習テーマを判定する。"""
    normalized = theme.lower()
    keywords = (
        "需要予測", "売上予測", "予測", "forecast", "時系列", "機械学習", "ml",
        "異常検知", "異常", "anomaly", "分類", "classif", "スコアリング", "score",
        "回帰", "regression", "最適化", "optimization", "優先順位", "順位付け",
        "ランキング", "rank", "配分", "配置", "割当", "スケジュール", "経路",
    )
    return any(keyword in normalized for keyword in keywords)


def enforce_oml_for_ml_poc_details(details: list[dict]) -> list[dict]:
    """ML系のPriority PoCにOML記述を適用した新しい一覧を返す。

    描画系からも呼ばれる参照ヘルパーのため、入力JSONは変更しない。
    """
    normalized = copy.deepcopy(details)
    for detail in normalized:
        if not is_ml_theme(str(detail.get("theme", ""))):
            continue
        theme = str(detail.get("theme", ""))
        if any(word in theme.lower() for word in (
            "優先順位", "順位付け", "ランキング", "rank", "最適化", "optimization",
            "配分", "配置", "割当", "スケジュール", "経路",
        )):
            detail["processing_steps"][1] = {
                "label": "OML・制約ロジックによる順位算出",
                "description": "Autonomous AI Database上のOracle Machine Learning（OML）で優先度をスコア化し、期限・処理能力・業務ルールなどの制約を反映して順位候補を算出します。OCI Generative AIは順位理由と確認すべき制約を説明します。",
            }
        else:
            detail["processing_steps"][1] = {
                "label": "OMLによる予測・判定",
                "description": "Autonomous AI Database上のOracle Machine Learning（OML）で、過去実績と関連要因から予測値またはスコアを算出します。OCI Generative AIは根拠と確認観点を担当者向けに説明します。",
            }
        detail["oci_roles"] = "業務連携・アプリケーション層は既存業務データの連携と結果表示を担います。Autonomous AI Databaseは履歴・特徴量・Oracle Machine Learning（OML）の予測／スコアを管理し、OCI Generative AIは根拠と対応案を説明します。"
    return normalized


def _poc_detail_matches_modality(detail: dict, poc: dict) -> bool:
    """既存JSONの詳細が、選定テーマの主処理を具体的に説明しているか確認する。"""
    modality = technical_modality_for(poc, detail)
    text = " ".join(str(value or "") for value in (
        detail.get("target_data"), detail.get("oci_roles"), detail.get("validation_plan"),
        *(
            f"{step.get('label', '')} {step.get('description', '')}"
            for step in detail.get("processing_steps", []) if isinstance(step, dict)
        ),
    )).lower()
    if modality["kind"] == "optimization":
        return (
            any(token in text for token in ("優先", "順位", "制約", "割当", "配分", "最適"))
            and any(token in text for token in ("締切", "期限", "進捗", "処理能力", "作業量", "リソース"))
            and any(token in text for token in ("oml", "oracle machine learning", "スコア"))
        )
    if modality["kind"] == "anomaly_ml":
        # 金額・申請を前提にした旧フォールバックを、物流・設備・品質など別業務へ
        # 誤適用しない。テーマ自体が不正／取引審査なら金額を許容する。
        financial_theme = any(token in str(poc.get("theme", "")).lower() for token in ("不正", "取引", "決済", "審査"))
        return any(token in text for token in ("異常", "逸脱", "予測", "検知")) and (financial_theme or "金額" not in text)
    if modality["kind"] == "rag":
        return "vector search" in text and any(token in text for token in ("根拠", "参照元", "引用"))
    return True


def normalize_generated_poc_logic_details(raw: object, assessment: dict) -> list[dict]:
    """LLM生成詳細を最終P1〜P3へIDで結合し、コピー文を拒否する。"""
    generated = raw.get("poc_logic_details") if isinstance(raw, dict) else raw
    priority_pocs = priority_pocs_for(assessment)
    required_keys = {
        "theme", "business_challenge", "target_data", "processing_steps", "oci_roles",
        "business_integration", "validation_plan", "design_notes",
    }
    if isinstance(generated, list) and len(generated) == 3 and all(
        isinstance(item, dict)
        and required_keys <= item.keys()
        and isinstance(item["processing_steps"], list)
        and len(item["processing_steps"]) == 3
        and all(isinstance(step, dict) and {"label", "description"} <= step.keys() for step in item["processing_steps"])
        for item in generated
    ):
        # 旧実装は表示順だけで ``poc_recommendations`` と結び付けていたため、
        # 詳細ページが別テーマのPoCを説明することがあった。ID（旧JSONは厳密な
        # テーマ一致）で対応できる場合だけ、LLM生成の詳細を採用する。
        aligned: list[dict] = []
        for poc in priority_pocs:
            detail = next((
                item for item in generated if _matching_source_poc(item, poc) is not None
            ), None)
            if detail is None:
                aligned = []
                break
            copied = copy.deepcopy(detail)
            copied["theme"] = str(poc["theme"])
            if poc.get("use_case_id"):
                copied["use_case_id"] = str(poc["use_case_id"])
            aligned.append(copied)
        if len(aligned) == 3:
            aligned = enforce_oml_for_ml_poc_details(aligned)
            if all(_poc_detail_matches_modality(detail, poc) for detail, poc in zip(aligned, priority_pocs)):
                processing_texts = [
                    " ".join(str(step.get("description") or "") for step in detail["processing_steps"])
                    for detail in aligned
                ]
                if all(
                    SequenceMatcher(None, processing_texts[left], processing_texts[right]).ratio() < 0.90
                    for left in range(3) for right in range(left + 1, 3)
                ):
                    return aligned
    return []


def poc_logic_details_for(assessment: dict) -> list[dict]:
    """3件の優先PoCを、1テーマ1ページで説明できる詳細な実現ロジックへ展開する。"""
    normalized_generated = normalize_generated_poc_logic_details(
        assessment.get("poc_logic_details"), assessment,
    )
    if normalized_generated:
        return normalized_generated

    priority_pocs = priority_pocs_for(assessment)

    details: list[dict] = []
    service_name = str(assessment.get("service_name") or "対象サービス")
    for item in priority_pocs:
        theme = str(item["theme"])
        lowered = theme.lower()
        if any(word in lowered for word in ("ocr", "帳票", "証明書", "書類", "文書", "画像")):
            business_challenge = f"{service_name}で扱う書式が異なる添付書類の入力・確認を自動化し、処理遅延と確認工数を抑えることを目指します。"
            target_data = "帳票・画像データ、既存の業務レコードとマスタ情報、抽出対象項目の定義、担当者の確認・修正履歴を対象とします。"
            steps = [
                {"label": "帳票受付・前処理", "description": "既存の業務画面またはAPIで帳票・画像を受け付け、ファイル形式、解像度、ページ欠損、重複を確認します。文書種別を判定し、読み取り対象の画像と関連する業務レコードを処理キューへ渡します。"},
                {"label": "複数モデルによる項目抽出・照合", "description": "文書種別ごとの抽出スキーマを複数のマルチモーダルLLM候補へ渡し、必要な業務項目を抽出します。マスタ・既存レコードと照合し、項目別の信頼度と不一致理由を返します。"},
                {"label": "正規化・確認画面への返却", "description": "抽出値、原画像位置、信頼度、使用モデルを共通JSONへ正規化してAutonomous AI Databaseへ保存します。低信頼・不一致項目のみを既存画面で強調し、担当者の修正結果を次回評価用に記録します。"},
            ]
            oci_roles = "業務連携・アプリケーション層は画像受付・既存画面連携・処理制御を担います。OCI Generative AIは帳票画像からの項目抽出と信頼度評価を実行し、Autonomous AI Databaseは抽出結果・監査情報・修正履歴を一元管理します。"
            integration = "既存の業務画面には、抽出済み項目と根拠画像位置を返します。担当者は低信頼項目だけを確認し、確定・差戻し・再読取を既存の業務フローで実施します。"
            validation = "文書種別・項目別の完全一致率、低信頼項目の検知率、1文書あたりの応答時間と推論単価、担当者の確認時間・修正率をモデル候補ごとに比較します。"
            notes = "画像・個人情報のアクセス権限、保持期間、推論・修正ログを設計します。文書種別追加時は抽出スキーマと正解データを管理し、精度劣化を定期的に監視します。"
        elif any(word in lowered for word in ("問い合わせ", "faq", "検索", "回答", "rag")):
            business_challenge = f"{service_name}に関する利用者・運用担当者からの問い合わせに対し、根拠を確認できる回答案を迅速に提示し、一次対応負荷を下げます。"
            target_data = "サービス仕様、FAQ、操作マニュアル、過去回答、利用者からの質問文、回答の採否・修正履歴を対象とします。"
            steps = [
                {"label": "ナレッジ整備・版管理", "description": "サービス仕様、FAQ、操作マニュアル、過去回答を文書単位で取り込み、対象機能・公開状態・改訂日などのメタデータを付与します。改訂版を識別して、検索対象を常に有効な文書に限定します。"},
                {"label": "根拠検索・回答案の生成", "description": "質問文を解析し、Autonomous AI DatabaseのVector Searchで関連箇所を検索します。OCI Generative AIは検索結果だけを根拠として回答案、参照元、回答できない場合の確認先を生成します。"},
                {"label": "業務画面での確認・改善", "description": "回答案と根拠リンクを問い合わせ画面へ返却し、担当者が承認・修正・未回答を判定します。質問、提示根拠、最終回答を蓄積し、未解決領域はFAQ・マニュアル改訂の対象にします。"},
            ]
            oci_roles = "業務連携・アプリケーション層は問い合わせ画面・認証・既存FAQ連携を担います。Autonomous AI Databaseは文書、メタデータ、ベクトル、回答履歴を管理し、OCI Generative AIは根拠付き回答案の生成と回答不能時の案内を担います。"
            integration = "利用者向けの自動回答と、オペレーター向けの回答支援を段階的に適用します。根拠が不足する場合は自動回答を抑止し、担当チームへのエスカレーションとして既存フローに引き継ぎます。"
            validation = "検索根拠の適合率、回答正確性、回答不能時の抑止精度、応答時間、担当者修正率、問い合わせ1件あたりの対応時間を評価し、公開範囲を判断します。"
            notes = "サービス仕様・運用ルールの改訂時には文書更新責任者と公開承認を明確にします。回答・参照ログを保管し、個人情報を含む質問は検索対象・表示内容を権限別に制御します。"
        elif any(word in lowered for word in (
            "優先順位", "順位付け", "ランキング", "rank", "最適化", "optimization",
            "配分", "配置", "割当", "スケジュール", "経路",
        )):
            business_challenge = f"{service_name}で同時に発生する作業・案件の優先順位を、締切・期限と処理能力を踏まえて一貫して提示し、担当者の判断時間と遅延リスクを抑えます。"
            target_data = "対象業務の案件・作業指示、締切・期限、進捗、作業量、利用可能な人員・設備・処理能力、業務ルール、担当者による順位変更履歴を対象とします。"
            steps = [
                {"label": "案件・締切・制約の整備", "description": "業務画面またはAPIから案件、締切・期限、進捗、作業量、利用可能なリソースをAutonomous AI Databaseへ連携します。欠損、重複、優先ルールの適用範囲を確認し、順位算出に使う条件を揃えます。"},
                {"label": "OML・制約ロジックによる順位算出", "description": "Oracle Machine Learning（OML）で遅延リスクや処理難易度をスコア化し、期限・処理能力・業務ルールを反映して順位候補を算出します。OCI Generative AIは順位理由と確認すべき制約を説明します。"},
                {"label": "現場確認・順位変更の学習", "description": "順位、期限リスク、理由を既存画面へ返し、担当者が採用・変更・保留を判断します。変更理由と完了実績を蓄積し、スコア、制約、画面表示を継続的に改善します。"},
            ]
            oci_roles = "業務連携・アプリケーション層は案件・期限・進捗・制約の連携と順位表示を担います。Autonomous AI Databaseは履歴・特徴量・OMLスコア・制約条件を管理し、OCI Generative AIは順位理由と例外時の確認観点を説明します。"
            integration = "既存の作業一覧へ順位候補と理由を追加し、担当者が採用・変更できる支援機能として組み込みます。自動指示は行わず、変更履歴と例外処理を現行運用へ戻します。"
            validation = "担当者順位との一致率、順位変更率、判断時間、期限内完了率、遅延案件の上位捕捉率、制約違反件数を比較し、業務上妥当なスコアと表示方法を決定します。"
            notes = "優先ルールの責任者、同点時の扱い、緊急割込み、担当者による上書き、監査ログを定義します。業務量や人員構成が変わった場合の再評価条件も合意します。"
        elif any(word in lowered for word in ("レポート", "日報", "週報", "月報", "報告書")):
            business_challenge = f"{service_name}の実績把握と報告作成を、元データへ遡れる形で迅速化し、担当者が要因確認と次の対応判断へ時間を使える状態を目指します。"
            target_data = "取引・業務実績、対象期間、比較期間、組織・商品等のマスタ、例外イベント、担当者による修正・配信履歴を対象とします。"
            steps = [
                {"label": "実績集計・比較条件の確定", "description": "業務システムまたはAPIから対象期間の実績とマスタを連携し、締め時刻、比較期間、集計粒度、欠損・取消・再処理の扱いを確認します。"},
                {"label": "差分分析・レポート案生成", "description": "Autonomous AI Databaseで指標と前年差・前日差を算出し、OCI Generative AIが変動要因、注視点、参照レコードを含むレポート案を生成します。"},
                {"label": "担当者確認・配信記録", "description": "レポート案と根拠指標を既存画面へ返し、担当者が追記・修正・承認します。確定版、修正箇所、配信先を記録し、次回評価へ利用します。"},
            ]
            oci_roles = "業務連携層は実績データの受付とレポート画面連携を担います。Autonomous AI Databaseは集計・比較指標と根拠を管理し、OCI Generative AIは変動要因と報告文案を生成します。"
            integration = "既存の日報・レポート承認画面へ、指標、変動要因、文章案、根拠リンクを返し、担当者の確認後に既存の配信経路へ渡します。"
            validation = "集計値の一致率、重要変動の捕捉率、根拠参照の正確性、作成時間、修正率、承認までの所要時間を現行作業と比較します。"
            notes = "締め処理前後の再生成、閲覧権限、元データ訂正時の版管理、誤要約時の差戻し、生成文と確定文の監査ログを設計します。"
        elif any(word in lowered for word in ("販促", "接客", "レコメンド", "推薦", "推奨", "購買", "顧客")):
            business_challenge = f"{service_name}の顧客・取引データを、説明可能な販促・接客候補へ変換し、担当者が対象と提案内容を安全に選べる状態を目指します。"
            target_data = "顧客属性、購買・利用履歴、商品・サービスマスタ、在庫・提供条件、施策履歴、提案の採否・反応結果を対象とします。"
            steps = [
                {"label": "顧客・購買データの許諾付き整備", "description": "顧客ID、購買履歴、商品マスタ、施策履歴を連携し、同意・利用目的、欠損、重複、対象外顧客、提案可能な商品条件を確認します。"},
                {"label": "候補抽出・提案理由の生成", "description": "Autonomous AI Databaseで顧客群と関連商品候補を抽出し、OCI Generative AIが購買傾向と適用条件に基づく提案文、理由、除外条件を生成します。"},
                {"label": "販促担当者の承認・反応記録", "description": "対象顧客群、提案内容、理由を販促・接客画面へ返し、担当者が承認・修正・除外します。配信・提示結果と反応を施策単位で記録します。"},
            ]
            oci_roles = "業務連携層は顧客・商品・施策データと販促画面を接続します。Autonomous AI Databaseは対象抽出と反応履歴を管理し、OCI Generative AIは条件に沿う提案文と説明を生成します。"
            integration = "販促管理または接客支援画面へ対象候補と提案理由を返し、担当者の承認後だけ既存チャネルへ連携します。自動配信はPoC範囲外とします。"
            validation = "対象抽出の妥当性、提案採用率、担当者修正率、対象外提示率、反応率、作成時間を現行施策と比較し、顧客群別に評価します。"
            notes = "利用目的・同意、要配慮情報の除外、テナント境界、過剰接触の抑止、提案理由と承認履歴、停止・オプトアウト手順を設計します。"
        elif any(word in lowered for word in ("開発", "保守", "コード", "テスト", "障害", "チケット")):
            business_challenge = f"{service_name}の開発・保守で、仕様・コード・障害情報の調査と変更案作成を支援し、レビュー品質を保ちながら初動を短縮します。"
            target_data = "リポジトリ、仕様書、API定義、課題・障害チケット、変更履歴、テストコード、レビューでの採否・修正履歴を対象とします。"
            steps = [
                {"label": "仕様・コード・履歴の権限別索引化", "description": "リポジトリ、仕様書、API定義、障害チケットを版・製品・権限とともに取り込み、秘密情報を除外して変更対象を検索できる状態にします。"},
                {"label": "関連箇所検索・変更／テスト案生成", "description": "Autonomous AI DatabaseのVector Searchで関連仕様・コード・障害を検索し、OCI Generative AIが影響範囲、変更案、テスト観点を根拠付きで生成します。"},
                {"label": "開発者レビュー・CI結果の記録", "description": "IDEまたはチケットへ候補と根拠を返し、開発者が採用・修正・却下します。コードレビューとCI結果を記録し、既存の承認フローを経て反映する構成を想定します。"},
            ]
            oci_roles = "業務連携層はリポジトリ・チケット・CIを接続します。Autonomous AI Databaseは仕様・コード・履歴のベクトルと監査記録を管理し、OCI Generative AIは根拠付き変更・テスト案を生成します。"
            integration = "IDE、コードレビュー、課題管理のいずれかへ候補と参照元を提示し、開発者レビューと既存CIを必須ゲートとして変更フローへ組み込みます。"
            validation = "関連箇所の再現率、提案採用率、レビュー修正率、テスト成功率、調査時間、誤った変更候補の抑止率を代表チケットで比較します。"
            notes = "リポジトリ権限、秘密情報・顧客データの取扱い、生成コードのライセンス確認、プロンプト・採否ログ、自動マージの適用範囲と承認方法を設計します。"
        elif any(word in lowered for word in ("不正", "異常", "予測", "分析", "検知")):
            business_challenge = f"{service_name}の大量処理の中で確認すべき業務レコードを優先度付けし、見落としを抑えながら担当者の確認時間を重点案件へ振り向けます。"
            target_data = "対象業務のイベント履歴、時刻、数量・状態・対象属性、正常／要確認の確定結果、既存ルールを対象とします。"
            steps = [
                {"label": "履歴データの収集・特徴量化", "description": "既存の業務システムまたはAPIからイベント履歴、数量、状態、更新時刻、対象属性、過去の確認結果をAutonomous AI Databaseへ連携します。業務種別や処理状態を分析用の特徴量として整備します。"},
                {"label": "異常候補の算出・理由付け", "description": "Autonomous AI Database上のOracle Machine Learningで、過去傾向や業務ルールから逸脱する候補をスコアリングします。OCI Generative AIは判定に使われた項目と確認観点を担当者向けの説明文へ変換します。"},
                {"label": "確認キューへの振り分け・学習", "description": "スコア、判定理由、関連データを確認画面へ返し、優先度順の確認キューを作成します。担当者の確定・誤検知・保留の結果をAutonomous AI Databaseへ蓄積し、閾値や特徴量の見直しに利用します。"},
            ]
            oci_roles = "業務連携・アプリケーション層は既存業務データの連携と確認キューの表示を担います。Autonomous AI Databaseは履歴データ・特徴量・Oracle Machine Learningのスコアを管理し、OCI Generative AIは担当者が理解できる検知理由と確認手順を生成します。"
            integration = "初期PoCでは自動確定を行わず、確認対象の絞り込みとして既存の審査画面へ組み込みます。担当者の判定結果をルール・モデル改善へ戻し、段階的に対象範囲を広げます。"
            validation = "既知事象に対する検知率、誤検知率、見逃し率、上位候補に含まれる割合、担当者の確認時間、理由説明の理解度を評価し、業務上許容できる閾値を決定します。"
            notes = "業務ルールやデータ分布の変化を監視し、再学習・閾値調整の責任者を定めます。スコアと根拠を担当者が確認し、最終判定の監査ログを保持する運用を想定します。"
        else:
            business_challenge = f"{theme}に関する判断・入力・確認を支援し、既存サービスの利用者体験と業務生産性を改善することを目指します。"
            target_data = f"{theme}に必要な業務入力、関連マスタ・履歴データ、担当者の確認・修正結果を対象とします。"
            steps = [
                {"label": "業務入力・関連データの収集", "description": "業務画面またはAPIで入力を受け付け、申請・マスタ・履歴データをAutonomous AI Databaseから参照します。対象データの欠損、権限、処理対象期間を確認してAI処理を開始します。"},
                {"label": "AI処理・業務データとの照合", "description": "OCI Generative AIで抽出、要約、分類、提案を実行し、Autonomous AI Databaseの業務データと照合します。結果だけでなく、参照情報、信頼度、担当者が確認すべき観点を生成します。"},
                {"label": "人手確認・既存業務への反映", "description": "生成結果を既存画面またはAPIへ返却し、担当者が確認・修正・承認します。確定結果と操作履歴をAutonomous AI Databaseへ記録し、評価・改善に必要なデータとして蓄積します。"},
            ]
            oci_roles = "業務連携・アプリケーション層は業務画面・API・認証を担います。Autonomous AI Databaseは業務データ、検索・分析結果、監査ログを管理し、OCI Generative AIは入力内容に応じた抽出・回答・提案を生成します。"
            integration = "既存業務を置き換えず、結果の確認・修正・承認を残した補助機能として組み込みます。既存画面・APIの変更範囲と利用者への表示方法をPoCで確認します。"
            validation = "出力品質、応答時間、処理単価、担当者の修正率、例外時の操作性、既存業務フローへの組み込みやすさを評価し、本番展開の対象範囲を決定します。"
            notes = "アクセス権限、入力データのマスキング、プロンプト・出力ログ、例外時の代替手順を設計します。本番化前に負荷、監視、障害時の復旧手順を確認します。"
        details.append({
            "theme": theme,
            "use_case_id": str(item.get("use_case_id", "")),
            "business_challenge": business_challenge,
            "target_data": target_data,
            "processing_steps": steps,
            "oci_roles": oci_roles,
            "business_integration": integration,
            "validation_plan": validation,
            "design_notes": notes,
        })
    return enforce_oml_for_ml_poc_details(details)


def normalize_adb_terminology(value: object) -> object:
    """レポート全体のDB提案表記をAutonomous AI Database／ADBへ統一する。"""
    if isinstance(value, str):
        value = re.sub(r"(?:OCI\s*)?DBCS", "Autonomous AI Database", value, flags=re.IGNORECASE)
        # 日本語文中では \b が境界として働かないため、略称は直接置換する。
        return re.sub(r"AADB", "ADB", value, flags=re.IGNORECASE)
    if isinstance(value, list):
        return [normalize_adb_terminology(item) for item in value]
    if isinstance(value, dict):
        return {key: normalize_adb_terminology(item) for key, item in value.items()}
    return value

CONSULTING_BASIS_VALUES = frozenset({"顧客入力", "公開資料", "分析仮説", "要確認"})
CONSULTING_LEVEL_VALUES = frozenset({"high", "medium", "low", "confirm"})
CONSULTING_PRIORITY_VALUES = frozenset({"P1", "P2", "P3", "Watch"})
POC_PORTFOLIO_SCHEMA_VERSION = "1"
TECHNICAL_PROPOSAL_SCHEMA_VERSION = "1"
# 構成図・概念ワイヤーフレームでは、現在確認できている前提と提案設計、PoCで
# 確かめるべき事項を同じ見た目で混在させない。JSONにも状態を残し、出力だけの
# 後付け注記にならないようにする。
TECHNICAL_PROPOSAL_STATE_VALUES = frozenset({"confirmed", "proposed", "verify"})
TECHNICAL_PROPOSAL_NODE_IDS = (
    "source_system", "event_api", "integration", "data_layer", "ai_processing",
    "human_action", "record_and_audit",
)
TECHNICAL_PROPOSAL_CONTROL_IDS = (
    "identity_tenant", "data_quality", "fallback", "operations",
)
# 優先PoCを「実現ロジック」だけでなく、判断・データ・評価・責任を持つ
# 再利用可能なチャーターとしてレビュー可能にする。未知の値は推測せず、JSONでは
# 明示的に ``confirm`` と保持し、表示時に「要確認」と変換する。
POC_CHARTER_SCHEMA_VERSION = "1"
MULTITENANT_GOVERNANCE_SCHEMA_VERSION = "1"
POC_CONFIRM_VALUE = "confirm"
# 優先候補の評価は「実績スコア」ではなく、現時点の比較仮説として扱う。値が
# 未確認なら 0 点や仮の数値を入れず ``confirm`` を保持することで、開始条件と
# 事実を混同しない。
POC_SELECTION_SCORECARD_SCHEMA_VERSION = "2"
POC_SELECTION_SCORECARD_LEGACY_SCHEMA_VERSION = "1"
POC_SELECTION_SCORE_VALUES = frozenset({1, 3, 5})
POC_SELECTION_SCORE_FIELDS = (
    "business_value", "feasibility", "data_readiness", "scale_readiness",
)
POC_SELECTION_START_GATE_FIELDS = (
    "scope_and_population", "data_contract_and_boundary", "baseline_and_comparator",
    "human_approval_and_fallback", "decision_owner",
)
# P1〜P3はLLMのラベルをそのまま採用せず、15件の候補に対する同じ評価式から
# 再現可能に選ぶ。未確認（confirm）は0点として扱わず、評価充足率を下げる。
POC_PRIORITY_DECISION_SCHEMA_VERSION = "3"
POC_PRIORITY_DECISION_LEGACY_SCHEMA_VERSION = "1"
POC_PRIORITY_DECISION_PREVIOUS_SCHEMA_VERSION = "2"
POC_PRIORITY_SCORE_WEIGHTS = {
    "business_value": 0.40,
    "feasibility": 0.30,
    "data_readiness": 0.20,
    "scale_readiness": 0.10,
}
POC_PRIORITY_MINIMUM_COVERAGE = 0.90
POC_PRIORITY_REQUIRED_DIMENSIONS = (
    "business_value", "feasibility", "data_readiness",
)
POC_PRIORITY_REASON_FIELDS = {
    "business_value": "value_reason",
    "feasibility": "feasibility_reason",
    "data_readiness": "data_readiness_reason",
    "scale_readiness": "scale_reason",
}
POC_PRIORITY_SOURCE_FIELDS = {
    "business_value": "value",
    "feasibility": "feasibility",
    "data_readiness": "data_readiness",
    "scale_readiness": "scale",
}
POC_LEVEL_SCORE_VALUES = {"high": 5, "medium": 3, "low": 1}
# P1は「開始済み」を意味しない。開始可否は、個別のチャーターと共通統制から
# 別データとして導出し、未充足の確認事項を隠さない。
POC_START_READINESS_SCHEMA_VERSION = "1"
POC_START_STATUS_VALUES = frozenset({"blocked", "ready_for_decision", "approved"})
# 導入先で測る業務成果と、提供者が判断する商品化・事業成果を混同しないための
# 任意JSON契約。既存JSONにこの項目が無い場合は、既存の優先PoC・価値領域から
# 数値を創作せずに安全な仮説モデルを導出する。
BUSINESS_VALUE_MODEL_SCHEMA_VERSION = "1"
BUSINESS_VALUE_MODEL_ROLES = frozenset({"provider", "operator", "mixed", "unknown"})
BUSINESS_VALUE_MODEL_BASELINE_VALUES = frozenset({"要確認", "入力済み", "公開資料", "PoCで取得"})
BUSINESS_VALUE_MODEL_DIMENSIONS = ("品質", "時間", "リスク", "定着")


def _front_text(value: object, limit: int = 120) -> str:
    """表示・契約用テキストを正規化するが、内容は削らずに保持する。

    ``limit`` は過去の呼び出しとの互換性のために残す。版面上の都合でデータを
    途中で失うと、PPTX・凍結JSONのどちらからも原文を復元できないため、折返し・
    領域拡張・追加ページを描画側で選ぶ方針にする。
    """
    _ = limit
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text


def consultative_page_lead(value: object, *, title: str = "", limit: int = 220) -> str:
    """ページ冒頭を、判断の背景と読後のアクションが伝わる説明文へ整える。

    入力・調査から得た固有の主張は先頭に残し、そのページで何を読み取り、
    どの判断につなげるのかを補う。未取得値を強調するだけの注意書きにはせず、
    顧客との合意形成に役立つ前向きな表現へ置き換える。
    """
    base = re.sub(r"\s+", " ", _plain_text(str(value or ""))).strip()
    if any(token in base for token in ("未取得", "仮置き", "事実化しません")):
        return (
            "本ページでは、優先テーマごとに必要となるデータ、評価方法、役割分担を具体化し、"
            "PoC開始前に関係者間で合意すべき前提を明確にします。対象範囲と成功基準を揃えることで、"
            "検証後の結果を実務に沿った次の投資判断へ確実につなげます。"
        )
    if not base:
        base = "対象サービスの業務・データ・利用者の状況を踏まえ、AI活用の論点を整理します。"
    # 呼び出し元が用意したページ固有の説明を尊重する。同じ汎用文を全ページへ
    # 自動追記すると、資料の論点が均質化し、ページ固有の意思決定が見えにくくなる。
    # 未確認値の前向きな言い換えだけを共通責務として残す。
    return _front_text(base, limit)


def _front_basis(value: object) -> str:
    basis = _front_text(value, 12)
    return basis if basis in CONSULTING_BASIS_VALUES else ""


def _front_confirm_or_text(value: object, limit: int) -> str:
    """未取得のPoC条件を、見かけ上のもっともらしい文章へ置き換えない。

    ``confirm`` はJSON上の機械可読な未確認値である。入力・モデル応答にある
    「要確認」「未確認」等もここで同じ値に揃えるため、後段の描画と検証が
    判定できる。空文字も未確認として扱うが、存在する情報を勝手に補わない。
    """
    text = _front_text(value, limit)
    if text.lower() in {
        "", "confirm", "tbd", "unknown", "n/a", "na", "-", "要確認", "未確認", "不明", "確認中",
    }:
        return POC_CONFIRM_VALUE
    return text


def poc_confirm_display(value: object) -> str:
    """顧客向けページでは内部値 ``confirm`` を明示的な「要確認」として表示する。"""
    return "要確認" if _front_text(value, 120) == POC_CONFIRM_VALUE else _front_text(value, 120)

def _front_items(raw: object, *, count: int, fields: dict[str, int],
                 enum_fields: dict[str, frozenset[str]] | None = None,
                 basis_fields: tuple[str, ...] = ()) -> list[dict[str, str]]:
    """LLM生成した表形式コンテンツの件数・列挙値・文字量を一括で検証する。"""
    if not isinstance(raw, list) or len(raw) != count:
        return []
    normalized: list[dict[str, str]] = []
    for item in raw:
        if not isinstance(item, dict):
            return []
        row = {field: _front_text(item.get(field), limit) for field, limit in fields.items()}
        if not all(row.values()):
            return []
        for field, allowed in (enum_fields or {}).items():
            if row.get(field) not in allowed:
                return []
        for field in basis_fields:
            if _front_basis(row.get(field)) == "":
                return []
        normalized.append(row)
    return normalized


def fallback_candidate_score_reason(item: dict[str, object], score_field: str) -> str:
    """旧JSONでも5/3/1の意味を隠さず、未確認事項を事実化しない理由文を返す。"""
    explicit = _front_text(item.get(POC_PRIORITY_REASON_FIELDS.get(score_field, "")), 96)
    if explicit:
        return explicit
    if score_field == "business_value":
        return _front_text(
            item.get("rationale") or "対象KPI・利用者・発生頻度への寄与を比較する分析仮説。", 96,
        )
    if score_field == "feasibility":
        return "必要なAI方式、既存業務との連携、人手確認の実装条件をPoCで確認する。"
    if score_field == "data_readiness":
        return "関連データは候補として特定し、期間・件数・品質・権限・正解条件を開始前に確認する。"
    return "共通機能化、顧客別設定、テナント分離、運用再利用の条件をPoC後に確認する。"


def fallback_candidate_data_next_action(item: dict[str, object]) -> str:
    """データ準備度を上げるための、顧客・業種に依存しない次アクション。"""
    explicit = _front_text(item.get("data_next_action"), 112)
    if explicit:
        return explicit
    return "代表期間の対象データを抽出し、件数・欠損・粒度・権限・正解／比較条件を確認する。"


def candidate_priority_evaluation(item: dict[str, object]) -> dict[str, object]:
    """候補1件の優先度を、未確認値を数値化せずに算定する。

    ``high/medium/low`` のみを 5/3/1 に写像し、``confirm`` は未評価として
    分母からも除く。従って、値が高そうに見えても必要な評価が欠ける候補は
    評価充足率が下がり、P1が「開始可能」を意味しないことを保てる。
    """
    weighted_total = 0.0
    evaluated_weight = 0.0
    dimensions: dict[str, dict[str, object]] = {}
    for score_field, source_field in POC_PRIORITY_SOURCE_FIELDS.items():
        level = str(item.get(source_field, "")).strip().lower()
        weight = POC_PRIORITY_SCORE_WEIGHTS[score_field]
        score = POC_LEVEL_SCORE_VALUES.get(level)
        dimensions[score_field] = {
            "level": level if level in CONSULTING_LEVEL_VALUES else POC_CONFIRM_VALUE,
            "score": score if score is not None else POC_CONFIRM_VALUE,
            "weight": weight,
            "reason": fallback_candidate_score_reason(item, score_field),
            "basis": _front_basis(item.get("basis")) or "要確認",
        }
        if score is not None:
            weighted_total += weight * score
            evaluated_weight += weight
    score = round(weighted_total / evaluated_weight, 3) if evaluated_weight else None
    coverage = round(evaluated_weight, 3)
    eligibility_reasons: list[str] = []
    if coverage < POC_PRIORITY_MINIMUM_COVERAGE:
        eligibility_reasons.append(
            f"評価充足率が{POC_PRIORITY_MINIMUM_COVERAGE * 100:.0f}%未満",
        )
    for field in POC_PRIORITY_REQUIRED_DIMENSIONS:
        if dimensions[field]["score"] == POC_CONFIRM_VALUE:
            eligibility_reasons.append(f"{field}が未確認")
    data_score = dimensions["data_readiness"]["score"]
    if isinstance(data_score, int) and data_score < 3:
        eligibility_reasons.append("data_readinessが3未満")
    return {
        "score": score,
        # 順位指数は、未確認項目を分母から除いて平均点を膨らませない。未確認を
        # 0点の事実とみなすのではなく、根拠がある加重点だけを順位へ寄与させる。
        "ranking_index": round(weighted_total, 3),
        "evaluation_coverage": coverage,
        "eligible_for_priority": not eligibility_reasons,
        "eligibility_reasons": eligibility_reasons,
        "dimensions": dimensions,
    }


def normalize_front_candidate_priorities(
    candidates: list[dict[str, str]], *, priority_override: dict[str, str] | None = None,
    catalog_by_no: dict[str, dict[str, str]] | None = None,
) -> list[dict[str, str]]:
    """15候補を評価し、実装方式の異なる代表3テーマへまとめる。

    基礎点は事業価値40%、実現性30%、データ準備度20%、横展開可能性10%で
    算定する。その後、適格候補を得点順に走査し、予測・異常検知、最適化、RAG、
    文書理解など未選択の技術方式を先に採る。顧客の実施テーマが未確定の段階で、
    同じ方式だけを三つ並べず、技術的な選択肢を比較できるようにする。
    """
    ranked: list[tuple[int, float, float, float, int, int, dict[str, object]]] = []
    for index, item in enumerate(candidates):
        evaluation = candidate_priority_evaluation(item)
        ranking_index = float(evaluation["ranking_index"])
        coverage = float(evaluation["evaluation_coverage"])
        data_score = evaluation["dimensions"]["data_readiness"]["score"]
        data_rank = float(data_score) if isinstance(data_score, (int, float)) else -1.0
        try:
            use_case_no = int(str(item.get("use_case_no", index + 1)))
        except (TypeError, ValueError):
            use_case_no = index + 1
        ranked.append((
            0 if evaluation["eligible_for_priority"] else 1,
            -ranking_index, -coverage, -data_rank, use_case_no, index, evaluation,
        ))
    ranked.sort()
    modality_by_index: dict[int, str] = {}
    for entry in ranked:
        index = entry[5]
        case = (catalog_by_no or {}).get(
            _front_text(candidates[index].get("use_case_no"), 3), {},
        )
        modality_seed = {
            "theme": " ".join(filter(None, (
                _front_text(case.get("theme"), 72),
                _front_text(case.get("ai_technology"), 72),
            ))),
            "reason": _front_text(candidates[index].get("rationale"), 120),
        }
        modality_by_index[index] = technical_modality_for(modality_seed)["kind"]

    selected_entries = ranked[:3]
    selection_method = "weighted_model_v2"
    if catalog_by_no:
        selected_entries = []
        selected_indices: set[int] = set()
        selected_modalities: set[str] = set()
        # 適格な方式を一つずつ先に採り、三方式に満たない場合だけ次点候補で補う。
        for entry in ranked:
            index, evaluation = entry[5], entry[6]
            modality = modality_by_index[index]
            if evaluation["eligible_for_priority"] and modality not in selected_modalities:
                selected_entries.append(entry)
                selected_indices.add(index)
                selected_modalities.add(modality)
                if len(selected_entries) == 3:
                    break
        for entry in ranked:
            if len(selected_entries) == 3:
                break
            if entry[5] not in selected_indices:
                selected_entries.append(entry)
                selected_indices.add(entry[5])
        selection_method = "technology_portfolio_v1"
    priority_by_index = {
        entry[5]: f"P{rank + 1}" for rank, entry in enumerate(selected_entries)
    }
    if priority_override:
        overridden = {
            index: priority_override.get(_front_text(item.get("use_case_no"), 3), "Watch")
            for index, item in enumerate(candidates)
        }
        if sorted(value for value in overridden.values() if value in {"P1", "P2", "P3"}) == ["P1", "P2", "P3"]:
            priority_by_index = {
                index: value for index, value in overridden.items() if value in {"P1", "P2", "P3"}
            }
            selection_method = "explicit_selection_override"
    evaluation_by_index = {entry[5]: entry[6] for entry in ranked}
    for index, item in enumerate(candidates):
        evaluation = evaluation_by_index[index]
        item["priority"] = priority_by_index.get(index, "Watch")
        item["priority_score"] = evaluation["score"]
        item["priority_ranking_index"] = evaluation["ranking_index"]
        item["priority_score_coverage"] = evaluation["evaluation_coverage"]
        item["priority_eligible"] = evaluation["eligible_for_priority"]
        item["priority_selection_status"] = (
            "eligible" if evaluation["eligible_for_priority"] else "provisional"
        )
        item["priority_selection_method"] = selection_method
        item["technical_modality"] = modality_by_index.get(index, "")
    return candidates


def legacy_priority_override_for(assessment: dict | None) -> dict[str, str]:
    """v1レビューJSONから移行した凍結順位を、再描画時も後方互換として引き継ぐ。"""
    if not isinstance(assessment, dict):
        return {}
    decision = assessment.get("poc_priority_decision")
    if not isinstance(decision, dict):
        return {}
    schema_version = _front_text(decision.get("schema_version"), 8)
    policy = decision.get("policy") if isinstance(decision.get("policy"), dict) else {}
    if (schema_version != POC_PRIORITY_DECISION_LEGACY_SCHEMA_VERSION
            and not (
                schema_version == POC_PRIORITY_DECISION_SCHEMA_VERSION
                and policy.get("legacy_priority_preserved") is True
            )):
        return {}
    items = decision.get("items")
    if not isinstance(items, list) or len(items) != 15:
        return {}
    result = {
        _front_text(item.get("use_case_no"), 3): _front_text(item.get("priority"), 10)
        for item in items if isinstance(item, dict)
    }
    priorities = sorted(value for value in result.values() if value in {"P1", "P2", "P3"})
    return result if priorities == ["P1", "P2", "P3"] else {}


CUSTOMER_PRIORITY_BINDING_SCHEMA_VERSION = "1"
CUSTOMER_PRIORITY_BINDING_SOURCE = "structured_customer_input"
FIXED_POC_SELECTION_BINDING_SCHEMA_VERSION = "1"
FIXED_POC_SELECTION_BINDING_SOURCE = "fixed_use_case_catalog"


def normalize_customer_priority_binding(
    raw: object, assessment: dict | None,
) -> dict[str, object]:
    """入力フォームで明示された優先テーマの拘束を検証する。

    通常の候補順位は重み付き評価で決める。一方、入力フォームで顧客が高優先と
    明示したテーマは、評価点が低い場合も勝手に別テーマへ置換せず、PoC開始
    ゲートの未充足として扱う必要がある。この小さな契約をassessment側に保持し、
    front matterを何度正規化しても同じP1〜P3へ収束させる。
    """
    if not isinstance(raw, dict) or not isinstance(assessment, dict):
        return {}
    if (
        _front_text(raw.get("schema_version"), 8)
        != CUSTOMER_PRIORITY_BINDING_SCHEMA_VERSION
        or _front_text(raw.get("source"), 40)
        != CUSTOMER_PRIORITY_BINDING_SOURCE
    ):
        return {}
    catalog = primary_use_case_catalog_for(assessment)
    if len(catalog) != 15:
        return {}
    catalog_by_id = {item["use_case_id"]: item for item in catalog}
    required_ids_raw = raw.get("required_use_case_ids")
    items_raw = raw.get("items")
    if (
        not isinstance(required_ids_raw, list)
        or not 1 <= len(required_ids_raw) <= 3
        or not isinstance(items_raw, list)
        or len(items_raw) != 3
    ):
        return {}
    required_ids = [_front_text(value, 8) for value in required_ids_raw]
    if not all(required_ids) or len(set(required_ids)) != len(required_ids):
        return {}

    items: list[dict[str, str]] = []
    seen_ids: set[str] = set()
    for index, item in enumerate(items_raw, 1):
        if not isinstance(item, dict):
            return {}
        priority = _front_text(item.get("priority"), 10)
        use_case_id = _front_text(item.get("use_case_id"), 8)
        case = catalog_by_id.get(use_case_id)
        row = {
            "priority": priority,
            "use_case_id": use_case_id,
            "use_case_no": _front_text(item.get("use_case_no"), 3),
            "theme": _front_text(item.get("theme"), 72),
        }
        if (
            priority != f"P{index}"
            or case is None
            or use_case_id in seen_ids
            or row["use_case_no"] != case["use_case_no"]
            or normalized_use_case_label(row["theme"])
            != normalized_use_case_label(case["theme"])
        ):
            return {}
        items.append(row)
        seen_ids.add(use_case_id)
    if any(use_case_id not in seen_ids for use_case_id in required_ids):
        return {}
    return {
        "schema_version": CUSTOMER_PRIORITY_BINDING_SCHEMA_VERSION,
        "source": CUSTOMER_PRIORITY_BINDING_SOURCE,
        "required_use_case_ids": required_ids,
        "items": items,
    }


def customer_priority_override_for(assessment: dict | None) -> dict[str, str]:
    """検証済みの顧客指定拘束を、候補番号から優先度への写像で返す。"""
    if not isinstance(assessment, dict):
        return {}
    binding = normalize_customer_priority_binding(
        assessment.get("customer_priority_binding"), assessment,
    )
    return {
        str(item["use_case_no"]): str(item["priority"])
        for item in binding.get("items", [])
        if isinstance(item, dict)
    }


def normalize_fixed_poc_selection_binding(
    raw: object, assessment: dict | None,
) -> dict[str, object]:
    """軽量カタログで明示したP1〜P3のID・順序拘束を検証する。

    これは顧客入力の高優先指定とは別の、合意済みカタログの再利用契約である。
    保持するのは一覧の優先マーカーと後続詳細をつなぐ不変ID・順序だけであり、
    PoCの理由、データ条件、実装内容は今回の分析で生成した値を用いる。
    """
    if not isinstance(raw, dict) or not isinstance(assessment, dict):
        return {}
    if (
        _front_text(raw.get("schema_version"), 8)
        != FIXED_POC_SELECTION_BINDING_SCHEMA_VERSION
        or _front_text(raw.get("source"), 40)
        != FIXED_POC_SELECTION_BINDING_SOURCE
    ):
        return {}
    catalog = primary_use_case_catalog_for(assessment)
    if len(catalog) != 15:
        return {}
    catalog_by_id = {item["use_case_id"]: item for item in catalog}
    items_raw = raw.get("items")
    if not isinstance(items_raw, list) or len(items_raw) != 3:
        return {}
    items: list[dict[str, str]] = []
    seen_ids: set[str] = set()
    for index, item in enumerate(items_raw, 1):
        if not isinstance(item, dict):
            return {}
        priority = _front_text(item.get("priority"), 10)
        use_case_id = _front_text(item.get("use_case_id"), 8)
        case = catalog_by_id.get(use_case_id)
        row = {
            "priority": priority,
            "use_case_id": use_case_id,
            "use_case_no": _front_text(item.get("use_case_no"), 3),
            "theme": _front_text(item.get("theme"), 72),
            "detail_slide_order": item.get("detail_slide_order"),
        }
        if (
            priority != f"P{index}"
            or case is None
            or use_case_id in seen_ids
            or row["use_case_no"] != case["use_case_no"]
            or normalized_use_case_label(row["theme"])
            != normalized_use_case_label(case["theme"])
            or isinstance(row["detail_slide_order"], bool)
            or row["detail_slide_order"] != index
        ):
            return {}
        items.append({
            "priority": priority,
            "use_case_id": use_case_id,
            "use_case_no": row["use_case_no"],
            "theme": row["theme"],
            "detail_slide_order": index,
        })
        seen_ids.add(use_case_id)
    return {
        "schema_version": FIXED_POC_SELECTION_BINDING_SCHEMA_VERSION,
        "source": FIXED_POC_SELECTION_BINDING_SOURCE,
        "selection_reason": _front_text(raw.get("selection_reason"), 160),
        "items": items,
    }


def fixed_poc_selection_override_for(assessment: dict | None) -> dict[str, str]:
    """検証済みの固定カタログ選定を、候補番号から優先度への写像で返す。"""
    if not isinstance(assessment, dict):
        return {}
    binding = normalize_fixed_poc_selection_binding(
        assessment.get("fixed_poc_selection_binding"), assessment,
    )
    return {
        str(item["use_case_no"]): str(item["priority"])
        for item in binding.get("items", [])
        if isinstance(item, dict)
    }


def normalize_consulting_front_matter(raw: object, assessment: dict | None = None,
                                      allowed_source_ids: set[str] | None = None) -> dict:
    """標準PPTXの構造化分析を検証し、旧JSONでも壊れない安全な形に正規化する。

    各ページで事実と仮説を区別できるよう、顧客固有の記述には根拠区分を必須にする。
    取得済みの公開資料以外のsource_idは受け入れないため、もっともらしい引用を混入させない。
    """
    if not isinstance(raw, dict):
        return {}
    allowed_source_ids = allowed_source_ids or {"I1", "M1", *(f"R{index}" for index in range(1, 7))}
    context_raw = raw.get("decision_context")
    service_raw = raw.get("service_model")
    diagnosis_raw = raw.get("operating_diagnosis")
    target_raw = raw.get("target_operating_model")
    prioritization_raw = raw.get("use_case_prioritization")
    delivery_raw = raw.get("delivery_plan")
    if not all(isinstance(value, dict) for value in (context_raw, service_raw, diagnosis_raw, target_raw, prioritization_raw, delivery_raw)):
        return {}

    decision_context = {
        "decision_question": _front_text(context_raw.get("decision_question"), 96),
        "objective": _front_text(context_raw.get("objective"), 140),
        "success_definition": _front_text(context_raw.get("success_definition"), 150),
        "in_scope": [_front_text(item, 48) for item in context_raw.get("in_scope", [])] if isinstance(context_raw.get("in_scope"), list) else [],
        "out_of_scope": [_front_text(item, 48) for item in context_raw.get("out_of_scope", [])] if isinstance(context_raw.get("out_of_scope"), list) else [],
    }
    if (not all(decision_context[key] for key in ("decision_question", "objective", "success_definition"))
            or len(decision_context["in_scope"]) != 3 or not all(decision_context["in_scope"])
            or len(decision_context["out_of_scope"]) != 2 or not all(decision_context["out_of_scope"])):
        return {}

    value_chain = _front_items(
        service_raw.get("value_chain"), count=5,
        fields={"stage": 20, "actor": 20, "activity": 68, "decision": 68, "data": 68, "output": 58, "basis": 12},
        basis_fields=("basis",),
    )
    stakeholders = _front_items(
        service_raw.get("stakeholders"), count=3,
        fields={"role": 24, "job": 60, "pain": 68, "value": 68, "basis": 12}, basis_fields=("basis",),
    )
    workflow = _front_items(
        diagnosis_raw.get("workflow"), count=5,
        fields={"phase": 20, "activity": 62, "bottleneck": 68, "impact": 60, "basis": 12}, basis_fields=("basis",),
    )
    data_assets = _front_items(
        diagnosis_raw.get("data_assets"), count=6,
        fields={"domain": 25, "records": 55, "decision_use": 60, "quality_check": 55, "readiness": 10, "basis": 12},
        enum_fields={"readiness": CONSULTING_LEVEL_VALUES}, basis_fields=("basis",),
    )
    issue_tree = _front_items(
        diagnosis_raw.get("issue_tree"), count=4,
        fields={"issue": 46, "cause": 68, "business_effect": 62, "validation_question": 68, "basis": 12}, basis_fields=("basis",),
    )
    principles = _front_items(
        target_raw.get("principles"), count=4,
        fields={"title": 28, "detail": 84, "basis": 12}, basis_fields=("basis",),
    )
    future_workflow = _front_items(
        target_raw.get("future_workflow"), count=5,
        fields={"step": 20, "human_role": 55, "ai_role": 68, "control": 55, "basis": 12}, basis_fields=("basis",),
    )
    criteria = _front_items(
        prioritization_raw.get("criteria"), count=4,
        fields={"name": 22, "definition": 58},
    )
    candidates = _front_items(
        prioritization_raw.get("candidates"), count=15,
        fields={"use_case_no": 3, "workflow_stage": 20, "value": 10, "feasibility": 10, "data_readiness": 10,
                "scale": 10, "priority": 10, "rationale": 80, "basis": 12},
        enum_fields={"value": CONSULTING_LEVEL_VALUES, "feasibility": CONSULTING_LEVEL_VALUES,
                     "data_readiness": CONSULTING_LEVEL_VALUES, "scale": CONSULTING_LEVEL_VALUES,
                     "priority": CONSULTING_PRIORITY_VALUES}, basis_fields=("basis",),
    )
    if candidates:
        raw_candidates = prioritization_raw.get("candidates")
        for candidate, raw_candidate in zip(
            candidates, raw_candidates if isinstance(raw_candidates, list) else [],
        ):
            raw_candidate = raw_candidate if isinstance(raw_candidate, dict) else {}
            for score_field, reason_field in POC_PRIORITY_REASON_FIELDS.items():
                candidate[reason_field] = (
                    _front_text(raw_candidate.get(reason_field), 96)
                    or fallback_candidate_score_reason(candidate, score_field)
                )
            candidate["data_next_action"] = (
                _front_text(raw_candidate.get("data_next_action"), 112)
                or fallback_candidate_data_next_action(candidate)
            )
    if [item["use_case_no"] for item in candidates] != [str(index) for index in range(1, 16)]:
        return {}
    value_chain_stages = [item["stage"] for item in value_chain]
    if (len(set(value_chain_stages)) != 5
            or set(item["workflow_stage"] for item in candidates) != set(value_chain_stages)
            or any(sum(item["workflow_stage"] == stage for item in candidates) != 3 for stage in value_chain_stages)):
        return {}
    fixed_override = fixed_poc_selection_override_for(assessment)
    customer_override = customer_priority_override_for(assessment)
    # 一つの候補表に異なる2つの拘束を同居させると、再描画時にどちらを優先するか
    # 曖昧になる。通常ワークフローはこの状態を作らないが、凍結JSONの改変も防ぐ。
    if fixed_override and customer_override:
        return {}
    priority_override = fixed_override or customer_override or legacy_priority_override_for(assessment)
    candidates = normalize_front_candidate_priorities(
        candidates, priority_override=priority_override,
        catalog_by_no={
            item["use_case_no"]: item for item in primary_use_case_catalog_for(assessment)
        },
    )
    if fixed_override:
        for candidate in candidates:
            selected = candidate.get("priority") in {"P1", "P2", "P3"}
            candidate["priority_selection_method"] = "fixed_use_case_catalog_v1"
            candidate["priority_selection_status"] = (
                "fixed_catalog_selected" if selected else "fixed_catalog_watch"
            )
            # カタログで合意済みの対象は、未確認データを「別候補へ置換する理由」
            # にせず、PoC開始ゲートで明示して解消する。
            if selected:
                candidate["priority_eligible"] = True
    elif customer_override:
        binding = normalize_customer_priority_binding(
            assessment.get("customer_priority_binding"), assessment,
        )
        required_ids = set(binding.get("required_use_case_ids", []))
        for candidate in candidates:
            use_case_id = canonical_use_case_id(candidate.get("use_case_no"))
            selected = candidate.get("priority") in {"P1", "P2", "P3"}
            candidate["priority_selection_method"] = "customer_priority_binding_v1"
            candidate["priority_selection_status"] = (
                "customer_required" if use_case_id in required_ids
                else "customer_selected_companion" if selected
                else "customer_binding_watch"
            )
            # 顧客指定テーマはデータ準備不足を自動除外理由にせず、開始ゲートで
            # 明示して解消する。評価点そのものはpriority decisionに保持する。
            if use_case_id in required_ids:
                candidate["priority_eligible"] = True
    # P1〜P3の表示名やページ上の並び順ではなく、15件の母集団に対する不変IDで
    # 後続のPoC、選定表、実装詳細を参照できるようにする。入力JSONの旧契約には
    # この列がなかったため、番号から常に再現可能なIDを派生する。
    for candidate in candidates:
        candidate["use_case_id"] = f"UC{int(candidate['use_case_no']):02d}"
    selection_logic = _front_items(
        prioritization_raw.get("selection_logic"), count=3,
        fields={"theme": 40, "why_now": 96, "proof_needed": 96, "depends_on": 68, "basis": 12}, basis_fields=("basis",),
    )
    phases = _front_items(
        delivery_raw.get("phases"), count=4,
        fields={"period": 22, "goal": 48, "deliverables": 104, "decision_gate": 68, "basis": 12}, basis_fields=("basis",),
    )
    governance = _front_items(
        delivery_raw.get("governance"), count=4,
        fields={"workstream": 26, "owner": 28, "decision": 62, "evidence": 68, "basis": 12}, basis_fields=("basis",),
    )
    evidence_notes = _front_items(
        raw.get("evidence_notes"), count=1,
        fields={"id": 12, "claim": 120, "basis": 12, "source_id": 16, "source_title": 72}, basis_fields=("basis",),
    )
    if not all((item["source_id"] in allowed_source_ids) for item in evidence_notes):
        return {}
    if not all((value_chain, stakeholders, workflow, data_assets, issue_tree, principles, future_workflow,
                criteria, candidates, selection_logic, phases, governance, evidence_notes)):
        return {}
    return {
        "decision_context": decision_context,
        "service_model": {"takeaway": _front_text(service_raw.get("takeaway"), 130),
                          "value_chain": value_chain, "stakeholders": stakeholders},
        "operating_diagnosis": {"workflow": workflow, "data_assets": data_assets, "issue_tree": issue_tree},
        "target_operating_model": {"principles": principles, "future_workflow": future_workflow},
        "use_case_prioritization": {"criteria": criteria, "candidates": candidates,
                                      "selection_logic": selection_logic},
        "delivery_plan": {"phases": phases, "governance": governance},
        "evidence_notes": evidence_notes,
    }

def fallback_consulting_front_matter(assessment: dict) -> dict:
    """既存の承認済みJSONを詳細レイアウトで開けるようにする、事実を増やさない後方互換用の骨格。"""
    business_value = assessment.get("business_value") if isinstance(assessment.get("business_value"), dict) else {}
    areas = [item for item in business_value.get("impact_areas", []) if isinstance(item, dict)]
    points = [_front_text(item, 64) for item in assessment.get("assessment_points", []) if _front_text(item, 64)]
    pocs = [item for item in assessment.get("poc_recommendations", []) if isinstance(item, dict)]
    use_cases = [item for item in assessment.get("use_cases", []) if isinstance(item, dict)]
    service_name = _front_text(assessment.get("service_name") or "対象サービス", 48)
    company_name = _front_text(assessment.get("company_name") or "ご提案先企業様", 48)
    summary = _front_text(assessment.get("executive_summary") or "対象サービスの業務・データ・利用者を基に、AI実装の優先テーマとPoCの判断条件を整理します。", 140)
    area_names = [_front_text(item.get("area") or f"価値領域{index}", 28) for index, item in enumerate(areas[:3], 1)]
    while len(area_names) < 3:
        area_names.append(f"価値領域{len(area_names) + 1}")
    while len(points) < 3:
        points.append("対象業務・データ・利用者を確認し、PoCで検証する判断条件を定義します。")
    cases_by_no = {str(item.get("no")): item for item in use_cases}

    def case_name(index: int) -> str:
        return _front_text(cases_by_no.get(str(index), {}).get("use_case") or f"AI活用候補 {index}", 48)

    def poc(index: int) -> dict:
        item = pocs[index] if index < len(pocs) else {}
        return {
            "theme": _front_text(item.get("theme") or case_name(index + 1), 40),
            "reason": _front_text(item.get("reason") or "事業価値・実現性・データ利用可否をPoCで確認するため。", 90),
            "first_step": _front_text(item.get("first_step") or "代表データ・対象業務・評価指標を合意する。", 90),
        }

    pocs = [poc(index) for index in range(3)]
    basis = "分析仮説"
    return {
        "decision_context": {
            "decision_question": f"{service_name}で、どの業務判断からAI実装を開始すべきか。",
            "objective": summary,
            "in_scope": [f"{service_name}の業務・利用者", "利用可能な業務データと判断", "優先PoCの範囲・評価条件"],
            "out_of_scope": ["顧客実績・外部効果の事実認定", "人による最終判断を代替する自動承認"],
            "success_definition": "代表業務・代表データで、事業価値、業務受容性、実装条件を同じ評価基準で判断できる状態。",
        },
        "service_model": {
            "takeaway": f"{service_name}の価値連鎖を、利用者の行動・業務判断・データ・提供価値の順に分解してAIの介在点を整理する。",
            "value_chain": [
                {"stage": "利用・入力", "actor": "利用者", "activity": "業務の依頼・入力・照会を行う。", "decision": "次に必要な情報・処理を選ぶ。", "data": "業務入力・依頼内容", "output": "受付・処理開始", "basis": "顧客入力"},
                {"stage": "業務処理", "actor": "担当者", "activity": "関連情報を確認し、処理を進める。", "decision": "優先順位・対応方法を決める。", "data": "業務レコード・マスタ", "output": "処理案・対応結果", "basis": basis},
                {"stage": "例外対応", "actor": "担当者・管理者", "activity": "例外・不足情報を確認する。", "decision": "エスカレーションと対応順を決める。", "data": "履歴・通知・操作ログ", "output": "確認キュー・対応指示", "basis": basis},
                {"stage": "品質・統制", "actor": "管理者", "activity": "処理状況と品質をレビューする。", "decision": "是正・ルール変更の要否を決める。", "data": "実績・監査ログ", "output": "改善アクション", "basis": basis},
                {"stage": "価値還元", "actor": "利用者・顧客", "activity": "結果を利用し、次の業務へつなげる。", "decision": "継続利用・追加対応を選ぶ。", "data": "利用結果・フィードバック", "output": "利用価値・継続改善", "basis": basis},
            ],
            "stakeholders": [
                {"role": "サービス利用者", "job": "必要な情報・結果を業務のタイミングで得る。", "pain": points[0], "value": "根拠と次アクションを理解しやすく受け取る。", "basis": "顧客入力"},
                {"role": "業務担当者", "job": "例外を判断し、正確に処理を完了する。", "pain": points[1], "value": "優先順位と確認観点に集中できる。", "basis": "顧客入力"},
                {"role": "サービス責任者", "job": "品質・収益性・利用定着を管理する。", "pain": points[2], "value": "標準化・横展開の判断材料を得る。", "basis": "顧客入力"},
            ],
        },
        "operating_diagnosis": {
            "workflow": [
                {"phase": "受付", "activity": "依頼・入力・照会を受け付ける。", "bottleneck": "情報の不足や表記揺れを人が確認する。", "impact": "初動と利用者への応答が遅れる。", "basis": basis},
                {"phase": "情報収集", "activity": "関連する記録・文書・履歴を参照する。", "bottleneck": "情報探索と照合が担当者依存になる。", "impact": "判断の再現性が下がる。", "basis": basis},
                {"phase": "優先判断", "activity": "対応順・担当・処理方法を選ぶ。", "bottleneck": "例外の優先度を一貫して付けにくい。", "impact": "重要案件への集中が難しい。", "basis": basis},
                {"phase": "実行・確認", "activity": "処理結果を確認し、必要に応じて修正する。", "bottleneck": "確認観点と根拠が画面上で分散する。", "impact": "工数と品質の両立が難しい。", "basis": basis},
                {"phase": "改善", "activity": "結果・例外・利用状況を振り返る。", "bottleneck": "判断結果を改善へ戻す仕組みが未整備。", "impact": "標準化・横展開の判断が遅れる。", "basis": basis},
            ],
            "data_assets": [
                {"domain": "業務トランザクション", "records": "受付・処理・完了の業務記録", "decision_use": "対象・優先度・状態を把握する。", "quality_check": "欠損・更新時刻・粒度を確認する。", "readiness": "medium", "basis": basis},
                {"domain": "マスタ情報", "records": "商品・顧客・設定などの基礎情報", "decision_use": "業務文脈と制約を補う。", "quality_check": "最新版・コード体系を確認する。", "readiness": "medium", "basis": basis},
                {"domain": "履歴・操作ログ", "records": "対応履歴・操作・承認結果", "decision_use": "類似事例と判断根拠を参照する。", "quality_check": "時系列・権限・保存期間を確認する。", "readiness": "confirm", "basis": basis},
                {"domain": "文書・ナレッジ", "records": "手順・FAQ・仕様・通知", "decision_use": "回答案・確認観点を提示する。", "quality_check": "更新責任・版管理を確認する。", "readiness": "medium", "basis": basis},
                {"domain": "品質・例外記録", "records": "差戻し・障害・エスカレーション", "decision_use": "リスク兆候と改善対象を特定する。", "quality_check": "分類基準・原因コードを確認する。", "readiness": "confirm", "basis": basis},
                {"domain": "利用者フィードバック", "records": "修正・評価・問い合わせ結果", "decision_use": "AI支援の受容性を測定する。", "quality_check": "同意・匿名化・評価方法を確認する。", "readiness": "low", "basis": basis},
            ],
            "issue_tree": [
                {"issue": area_names[0], "cause": "情報と判断根拠が複数の記録・担当者に分散する。", "business_effect": "対応の速さと顧客体験にばらつきが生じる。", "validation_question": "AIが必要情報と次アクションを根拠付きで提示できるか。", "basis": basis},
                {"issue": area_names[1], "cause": "例外確認と優先順位付けが人の経験に依存する。", "business_effect": "重要業務への工数配分が最適化しにくい。", "validation_question": "AI支援で確認対象と確認時間を絞り込めるか。", "basis": basis},
                {"issue": area_names[2], "cause": "結果・修正・例外の知見が運用改善へ戻りにくい。", "business_effect": "品質改善と横展開の意思決定が遅れる。", "validation_question": "結果を評価・改善サイクルへ接続できるか。", "basis": basis},
                {"issue": "実装・統制条件", "cause": "権限、例外、人手確認の設計を業務ごとに確認する必要がある。", "business_effect": "有効でも業務へ定着しないリスクがある。", "validation_question": "既存の業務画面・権限・監査へ無理なく組み込めるか。", "basis": basis},
            ],
        },
        "target_operating_model": {
            "principles": [
                {"title": "根拠を示す", "detail": "AIの提案と参照情報を利用者が確認できる形で提示する。", "basis": basis},
                {"title": "人が最終判断する", "detail": "例外・承認・顧客影響のある判断は担当者が確定する。", "basis": basis},
                {"title": "業務に埋め込む", "detail": "既存の画面・API・運用ルールに沿ってAI支援を提供する。", "basis": basis},
                {"title": "実測で拡大する", "detail": "利用・品質・事業KPIを測定し、対象範囲を段階的に広げる。", "basis": basis},
            ],
            "future_workflow": [
                {"step": "受付", "human_role": "対象範囲・権限・例外を確認する。", "ai_role": "入力を整理し、必要情報を補助する。", "control": "入力範囲・権限を制御する。", "basis": basis},
                {"step": "探索", "human_role": "業務文脈と重要度を確認する。", "ai_role": "関連記録・根拠・類似事例を提示する。", "control": "参照元と回答根拠を記録する。", "basis": basis},
                {"step": "判断支援", "human_role": "提案を採用・修正・却下する。", "ai_role": "優先順位・回答案・確認観点を提示する。", "control": "自動確定を行わず判断履歴を残す。", "basis": basis},
                {"step": "実行", "human_role": "既存フローで業務を完了する。", "ai_role": "次工程への連携情報を整形する。", "control": "例外時は既存の手順へ戻す。", "basis": basis},
                {"step": "学習・改善", "human_role": "効果と運用課題をレビューする。", "ai_role": "修正・評価データを分析用に蓄積する。", "control": "評価指標と変更承認を管理する。", "basis": basis},
            ],
        },
        "use_case_prioritization": {
            "criteria": [
                {"name": "事業価値", "definition": "顧客価値・収益性・品質への寄与"},
                {"name": "実現性", "definition": "既存システム・業務への組込みやすさ"},
                {"name": "データ準備度", "definition": "対象データの利用可否・品質・権限"},
                {"name": "横展開性", "definition": "他の顧客・業務へ標準化できる可能性"},
            ],
            "candidates": [
                {"use_case_no": str(index), "workflow_stage": ["利用・入力", "業務処理", "例外対応", "品質・統制", "価値還元"][(index - 1) // 3], "value": "high" if index <= 3 else "medium", "feasibility": "high" if index <= 5 else "medium",
                 "data_readiness": "medium", "scale": "high" if index <= 3 else "medium",
                 "priority": "P1" if index == 1 else "P2" if index == 2 else "P3" if index == 3 else "Watch",
                 "value_reason": f"{case_name(index)}が対象業務の品質・時間・リスクへ与える寄与を比較する。",
                 "feasibility_reason": "必要なAI方式、既存業務との連携、人手確認の実装条件をPoCで確認する。",
                 "data_readiness_reason": "関連データは候補として特定し、期間・件数・品質・権限・正解条件を開始前に確認する。",
                 "data_next_action": "代表期間の対象データを抽出し、件数・欠損・粒度・権限・正解／比較条件を確認する。",
                 "scale_reason": "共通機能化、顧客別設定、テナント分離、運用再利用の条件を確認する。",
                 "rationale": f"{case_name(index)}を代表業務・代表データで評価する。", "basis": basis}
                for index in range(1, 16)
            ],
            "selection_logic": [
                {"theme": pocs[index]["theme"], "why_now": pocs[index]["reason"], "proof_needed": pocs[index]["first_step"], "depends_on": "対象データ、利用者、既存業務との連携条件", "basis": basis}
                for index in range(3)
            ],
        },
        "delivery_plan": {
            "phases": [
                {"period": "0〜4週", "goal": "対象・評価設計を合意する", "deliverables": "対象業務、代表データ、KPI、権限・例外条件", "decision_gate": "PoC対象と比較条件を承認する。", "basis": basis},
                {"period": "5〜12週", "goal": "代表業務でPoCを検証する", "deliverables": "AI支援機能、評価結果、利用者フィードバック", "decision_gate": "品質・工数・受容性の合格条件を評価する。", "basis": basis},
                {"period": "PoC後", "goal": "事業化の範囲を決める", "deliverables": "本番要件、連携設計、運用・費用計画", "decision_gate": "標準機能化または追加検証を判断する。", "basis": basis},
                {"period": "横展開", "goal": "運用を標準化し拡大する", "deliverables": "テンプレート、監視・改善手順、展開計画", "decision_gate": "対象顧客・業務への展開順を決める。", "basis": basis},
            ],
            "governance": [
                {"workstream": "事業価値", "owner": "事業責任者", "decision": "優先業務と成功基準を決める。", "evidence": "KPI基準値・利用者の評価・顧客影響", "basis": basis},
                {"workstream": "業務・データ", "owner": "業務・データ責任者", "decision": "対象データと人手確認範囲を決める。", "evidence": "データ品質・権限・例外ケース", "basis": basis},
                {"workstream": "技術・統制", "owner": "IT・セキュリティ責任者", "decision": "連携・監査・運用方式を決める。", "evidence": "接続方式・ログ・障害時手順", "basis": basis},
                {"workstream": "展開", "owner": "サービス企画責任者", "decision": "標準機能化と横展開の順序を決める。", "evidence": "PoC結果・費用・保守体制", "basis": basis},
            ],
        },
        "evidence_notes": [{"id": "F1", "claim": f"{company_name}・{service_name}に関する顧客固有の事実は、入力内容または公開資料で確認してから確定する。", "basis": "顧客入力", "source_id": "I1", "source_title": "顧客入力・アセスメント前提"}],
    }


def repair_consulting_front_matter(raw: object, assessment: dict,
                                   allowed_source_ids: set[str]) -> dict:
    """部分的に形式を外したLLM出力を、安全な既定値とマージして回復する。

    生成内容は basis が有効な項目だけを採用する。足りない項目・列挙値・工程配分は、
    顧客固有の事実を追加しないフォールバックで補うため、形式不備だけで全ページが
    汎用文になることを避けつつ、根拠のない主張は通さない。
    """
    base = fallback_consulting_front_matter(assessment)
    if not isinstance(raw, dict):
        return base

    def merged_text(value: object, fallback_value: str, limit: int = 120) -> str:
        return _front_text(value, limit) or fallback_value

    def merged_rows(raw_rows: object, base_rows: list[dict[str, str]], fields: tuple[str, ...], *,
                    require_basis: bool = True) -> list[dict[str, str]]:
        if not isinstance(raw_rows, list):
            return base_rows
        rows: list[dict[str, str]] = []
        for index, fallback_row in enumerate(base_rows):
            candidate = raw_rows[index] if index < len(raw_rows) else None
            if not isinstance(candidate, dict) or (require_basis and _front_basis(candidate.get("basis")) == ""):
                rows.append(fallback_row.copy())
                continue
            row = {field: merged_text(candidate.get(field), fallback_row[field]) for field in fields}
            if require_basis:
                row["basis"] = _front_basis(candidate.get("basis"))
            rows.append(row)
        return rows

    def normalized_level(value: object, fallback_value: str) -> str:
        mapping = {"高": "high", "中": "medium", "低": "low", "要確認": "confirm"}
        value = mapping.get(_front_text(value, 12), _front_text(value, 12))
        return value if value in CONSULTING_LEVEL_VALUES else fallback_value

    def normalized_priority(value: object, fallback_value: str) -> str:
        mapping = {"優先1": "P1", "優先2": "P2", "優先3": "P3", "保留": "Watch", "watch": "Watch"}
        value = mapping.get(_front_text(value, 12), _front_text(value, 12))
        return value if value in CONSULTING_PRIORITY_VALUES else fallback_value

    raw_context = raw.get("decision_context") if isinstance(raw.get("decision_context"), dict) else {}
    base_context = base["decision_context"]
    context = {
        "decision_question": merged_text(raw_context.get("decision_question"), base_context["decision_question"], 96),
        "objective": merged_text(raw_context.get("objective"), base_context["objective"], 140),
        "success_definition": merged_text(raw_context.get("success_definition"), base_context["success_definition"], 150),
        "in_scope": [_front_text(item, 48) for item in raw_context.get("in_scope", [])[:3]] if isinstance(raw_context.get("in_scope"), list) and len(raw_context.get("in_scope", [])) >= 3 else base_context["in_scope"],
        "out_of_scope": [_front_text(item, 48) for item in raw_context.get("out_of_scope", [])[:2]] if isinstance(raw_context.get("out_of_scope"), list) and len(raw_context.get("out_of_scope", [])) >= 2 else base_context["out_of_scope"],
    }
    if not all(context["in_scope"]) or not all(context["out_of_scope"]):
        context["in_scope"], context["out_of_scope"] = base_context["in_scope"], base_context["out_of_scope"]

    raw_service = raw.get("service_model") if isinstance(raw.get("service_model"), dict) else {}
    base_service = base["service_model"]
    value_chain = merged_rows(raw_service.get("value_chain"), base_service["value_chain"],
                              ("stage", "actor", "activity", "decision", "data", "output"))
    # 候補の工程配置は5工程を三つずつ使うため、工程名の重複は許可しない。
    if len({item["stage"] for item in value_chain}) != 5:
        value_chain = base_service["value_chain"]
    service = {
        "takeaway": merged_text(raw_service.get("takeaway"), base_service["takeaway"], 130),
        "value_chain": value_chain,
        "stakeholders": merged_rows(raw_service.get("stakeholders"), base_service["stakeholders"], ("role", "job", "pain", "value")),
    }

    raw_diagnosis = raw.get("operating_diagnosis") if isinstance(raw.get("operating_diagnosis"), dict) else {}
    base_diagnosis = base["operating_diagnosis"]
    data_assets = merged_rows(raw_diagnosis.get("data_assets"), base_diagnosis["data_assets"],
                              ("domain", "records", "decision_use", "quality_check", "readiness"))
    for index, item in enumerate(data_assets):
        item["readiness"] = normalized_level(item.get("readiness"), base_diagnosis["data_assets"][index]["readiness"])
    diagnosis = {
        "workflow": merged_rows(raw_diagnosis.get("workflow"), base_diagnosis["workflow"], ("phase", "activity", "bottleneck", "impact")),
        "data_assets": data_assets,
        "issue_tree": merged_rows(raw_diagnosis.get("issue_tree"), base_diagnosis["issue_tree"], ("issue", "cause", "business_effect", "validation_question")),
    }

    raw_target = raw.get("target_operating_model") if isinstance(raw.get("target_operating_model"), dict) else {}
    base_target = base["target_operating_model"]
    target = {
        "principles": merged_rows(raw_target.get("principles"), base_target["principles"], ("title", "detail")),
        "future_workflow": merged_rows(raw_target.get("future_workflow"), base_target["future_workflow"], ("step", "human_role", "ai_role", "control")),
    }

    raw_prioritization = raw.get("use_case_prioritization") if isinstance(raw.get("use_case_prioritization"), dict) else {}
    base_prioritization = base["use_case_prioritization"]
    criteria = merged_rows(raw_prioritization.get("criteria"), base_prioritization["criteria"], ("name", "definition"), require_basis=False)
    raw_candidates = raw_prioritization.get("candidates") if isinstance(raw_prioritization.get("candidates"), list) else []
    candidate_by_no = {str(item.get("use_case_no")): item for item in raw_candidates if isinstance(item, dict)}
    candidates: list[dict[str, str]] = []
    for index, fallback_candidate in enumerate(base_prioritization["candidates"], 1):
        candidate = candidate_by_no.get(str(index))
        if not isinstance(candidate, dict) or _front_basis(candidate.get("basis")) == "":
            candidates.append(fallback_candidate.copy())
            continue
        candidates.append({
            "use_case_no": str(index), "workflow_stage": value_chain[(index - 1) // 3]["stage"],
            "value": normalized_level(candidate.get("value"), fallback_candidate["value"]),
            "feasibility": normalized_level(candidate.get("feasibility"), fallback_candidate["feasibility"]),
            "data_readiness": normalized_level(candidate.get("data_readiness"), fallback_candidate["data_readiness"]),
            "scale": normalized_level(candidate.get("scale"), fallback_candidate["scale"]),
            "priority": normalized_priority(candidate.get("priority"), fallback_candidate["priority"]),
            "value_reason": merged_text(candidate.get("value_reason"), fallback_candidate["value_reason"], 96),
            "feasibility_reason": merged_text(candidate.get("feasibility_reason"), fallback_candidate["feasibility_reason"], 96),
            "data_readiness_reason": merged_text(candidate.get("data_readiness_reason"), fallback_candidate["data_readiness_reason"], 96),
            "data_next_action": merged_text(candidate.get("data_next_action"), fallback_candidate["data_next_action"], 112),
            "scale_reason": merged_text(candidate.get("scale_reason"), fallback_candidate["scale_reason"], 96),
            "rationale": merged_text(candidate.get("rationale"), fallback_candidate["rationale"], 80),
            "basis": _front_basis(candidate.get("basis")),
        })
    prioritization = {
        "criteria": criteria, "candidates": candidates,
        "selection_logic": merged_rows(raw_prioritization.get("selection_logic"), base_prioritization["selection_logic"], ("theme", "why_now", "proof_needed", "depends_on")),
    }

    raw_delivery = raw.get("delivery_plan") if isinstance(raw.get("delivery_plan"), dict) else {}
    base_delivery = base["delivery_plan"]
    delivery = {
        "phases": merged_rows(raw_delivery.get("phases"), base_delivery["phases"], ("period", "goal", "deliverables", "decision_gate")),
        "governance": merged_rows(raw_delivery.get("governance"), base_delivery["governance"], ("workstream", "owner", "decision", "evidence")),
    }

    raw_notes = raw.get("evidence_notes") if isinstance(raw.get("evidence_notes"), list) else []
    note = raw_notes[0] if raw_notes and isinstance(raw_notes[0], dict) else None
    fallback_note = base["evidence_notes"][0]
    if not isinstance(note, dict) or _front_basis(note.get("basis")) == "" or _front_text(note.get("source_id"), 16) not in allowed_source_ids:
        evidence_notes = [fallback_note]
    else:
        evidence_notes = [{
            "id": merged_text(note.get("id"), fallback_note["id"], 12),
            "claim": merged_text(note.get("claim"), fallback_note["claim"], 120),
            "basis": _front_basis(note.get("basis")),
            "source_id": _front_text(note.get("source_id"), 16),
            "source_title": merged_text(note.get("source_title"), fallback_note["source_title"], 72),
        }]
    return {
        "decision_context": context, "service_model": service, "operating_diagnosis": diagnosis,
        "target_operating_model": target, "use_case_prioritization": prioritization,
        "delivery_plan": delivery, "evidence_notes": evidence_notes,
    }


def consulting_front_matter_for(assessment: dict) -> dict:
    """新規の詳細JSONを優先し、旧JSONでは事実を追加しないフォールバックを返す。"""
    normalized = normalize_consulting_front_matter(assessment.get("consulting_front_matter"), assessment)
    return normalized or fallback_consulting_front_matter(assessment)


def canonical_use_case_id(value: object) -> str:
    """Primary serviceの候補番号を、文書全体で使う不変IDへ正規化する。"""
    try:
        number = int(str(value).strip())
    except (TypeError, ValueError):
        return ""
    return f"UC{number:02d}" if 1 <= number <= 15 else ""


def normalized_use_case_label(value: object) -> str:
    """表示上の空白・記号の差だけを吸収した、候補名の厳密照合キーを返す。"""
    text = unicodedata.normalize("NFKC", str(value or "")).casefold()
    return re.sub(r"[\s\u3000・/／,:：;；()（）\[\]【】「」『』\-－_]+", "", text)


def primary_use_case_catalog_for(assessment: dict) -> list[dict[str, str]]:
    """優先化の母集団となる代表サービスの15候補を、番号・ID・名称で固定する。"""
    raw_cases = assessment.get("use_cases")
    if not isinstance(raw_cases, list):
        return []
    by_no: dict[int, dict] = {}
    for item in raw_cases:
        if not isinstance(item, dict):
            continue
        try:
            number = int(str(item.get("no", "")).strip())
        except (TypeError, ValueError):
            continue
        title = _front_text(item.get("use_case"), 72)
        if not 1 <= number <= 15 or not title or number in by_no:
            continue
        by_no[number] = item
    if set(by_no) != set(range(1, 16)):
        return []
    return [
        {
            "use_case_id": canonical_use_case_id(number),
            "use_case_no": str(number),
            "theme": _front_text(by_no[number].get("use_case"), 72),
            "ai_technology": _front_text(by_no[number].get("ai_technology"), 72),
        }
        for number in range(1, 16)
    ]


def _matching_source_poc(item: object, case: dict[str, str]) -> dict | None:
    """元のPoC提案を、IDまたは候補名が厳密に対応するときだけ採用する。"""
    if not isinstance(item, dict):
        return None
    source_id = _front_text(item.get("use_case_id"), 8)
    source_theme = normalized_use_case_label(item.get("theme"))
    target_theme = normalized_use_case_label(case["theme"])
    if source_id:
        return item if source_id == case["use_case_id"] and source_theme == target_theme else None
    return item if source_theme and source_theme == target_theme else None


def _matching_selection_logic(item: object, case: dict[str, str]) -> dict | None:
    """旧selection_logicにも安全に対応し、別テーマの検証条件を流用しない。"""
    if not isinstance(item, dict):
        return None
    source_id = _front_text(item.get("use_case_id"), 8)
    source_theme = normalized_use_case_label(item.get("theme"))
    target_theme = normalized_use_case_label(case["theme"])
    if source_id:
        return item if source_id == case["use_case_id"] and source_theme == target_theme else None
    return item if source_theme and source_theme == target_theme else None


def normalize_poc_portfolio(raw: object, assessment: dict,
                            front_matter: dict | None = None) -> dict:
    """優先PoCの単一ソースを検証する。

    すべての表示ページは ``poc_portfolio.items`` の ``use_case_id`` を参照する。
    この契約により、候補表のP1〜P3、PoC選定、詳細設計、構成図が別々のテーマを
    示すことを防ぐ。旧JSONは ``fallback_poc_portfolio`` で安全に派生できるため、
    既存のレビュー済みJSONを破壊しない。
    """
    if not isinstance(raw, dict) or _front_text(raw.get("schema_version"), 8) != POC_PORTFOLIO_SCHEMA_VERSION:
        return {}
    catalog = primary_use_case_catalog_for(assessment)
    if len(catalog) != 15:
        return {}
    case_by_id = {item["use_case_id"]: item for item in catalog}
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    prioritization = front.get("use_case_prioritization") if isinstance(front, dict) else None
    candidates = prioritization.get("candidates") if isinstance(prioritization, dict) else None
    if not isinstance(candidates, list):
        return {}
    candidate_by_priority: dict[str, dict] = {}
    for candidate in candidates:
        if not isinstance(candidate, dict):
            continue
        priority = _front_text(candidate.get("priority"), 10)
        candidate_id = canonical_use_case_id(candidate.get("use_case_no"))
        if priority in {"P1", "P2", "P3"} and candidate_id and priority not in candidate_by_priority:
            candidate_by_priority[priority] = candidate
    if set(candidate_by_priority) != {"P1", "P2", "P3"}:
        return {}

    items_raw = raw.get("items")
    if not isinstance(items_raw, list) or len(items_raw) != 3:
        return {}
    normalized: list[dict[str, str]] = []
    seen_ids: set[str] = set()
    seen_priorities: set[str] = set()
    for item in items_raw:
        if not isinstance(item, dict):
            return {}
        priority = _front_text(item.get("priority"), 10)
        use_case_id = _front_text(item.get("use_case_id"), 8)
        case = case_by_id.get(use_case_id)
        candidate = candidate_by_priority.get(priority)
        row = {
            "priority": priority,
            "use_case_id": use_case_id,
            "use_case_no": _front_text(item.get("use_case_no"), 3),
            "theme": _front_text(item.get("theme"), 72),
            "reason": _front_text(item.get("reason"), 108),
            "first_step": _front_text(item.get("first_step"), 108),
            "depends_on": _front_text(item.get("depends_on"), 96),
            "basis": _front_basis(item.get("basis")),
            "architecture_implementation": _front_text(item.get("architecture_implementation"), 108),
        }
        if (priority not in {"P1", "P2", "P3"} or priority in seen_priorities
                or use_case_id in seen_ids or case is None or candidate is None
                or row["use_case_no"] != case["use_case_no"]
                or canonical_use_case_id(candidate.get("use_case_no")) != use_case_id
                or normalized_use_case_label(row["theme"]) != normalized_use_case_label(case["theme"])
                or not all(row[key] for key in ("reason", "first_step", "depends_on", "basis"))):
            return {}
        normalized.append(row)
        seen_ids.add(use_case_id)
        seen_priorities.add(priority)
    if [item["priority"] for item in normalized] != ["P1", "P2", "P3"]:
        return {}
    return {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": normalized}


def fallback_poc_portfolio(assessment: dict, front_matter: dict | None = None) -> dict:
    """旧形式の候補・選定データから、整合した優先PoCポートフォリオを派生する。"""
    catalog = primary_use_case_catalog_for(assessment)
    if len(catalog) != 15:
        return {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": []}
    case_by_id = {item["use_case_id"]: item for item in catalog}
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    prioritization = front.get("use_case_prioritization") if isinstance(front, dict) else {}
    candidates = prioritization.get("candidates") if isinstance(prioritization, dict) else []
    selection_logic = prioritization.get("selection_logic") if isinstance(prioritization, dict) else []
    if not isinstance(candidates, list) or not isinstance(selection_logic, list):
        return {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": []}
    candidate_by_priority = {
        str(candidate.get("priority")): candidate
        for candidate in candidates if isinstance(candidate, dict)
        and str(candidate.get("priority")) in {"P1", "P2", "P3"}
    }
    if set(candidate_by_priority) != {"P1", "P2", "P3"}:
        return {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": []}
    source_pocs = assessment.get("poc_recommendations")
    source_pocs = source_pocs if isinstance(source_pocs, list) else []
    items: list[dict[str, str]] = []
    for priority in ("P1", "P2", "P3"):
        candidate = candidate_by_priority[priority]
        use_case_id = canonical_use_case_id(candidate.get("use_case_no"))
        case = case_by_id.get(use_case_id)
        if case is None:
            return {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": []}
        source_poc = next((item for item in source_pocs if _matching_source_poc(item, case)), None)
        selection = next((item for item in selection_logic if _matching_selection_logic(item, case)), None)
        reason = _front_text(
            (source_poc or {}).get("reason") or (selection or {}).get("why_now")
            or candidate.get("rationale") or "対象業務の事業価値・実現性・データ利用可否を代表範囲で確認するため。",
            108,
        )
        first_step = _front_text(
            (source_poc or {}).get("first_step") or (selection or {}).get("proof_needed")
            or "代表業務、対象データ、比較条件、判断責任者を合意してPoCの範囲を定める。",
            108,
        )
        depends_on = _front_text(
            (selection or {}).get("depends_on")
            or "対象データの利用可否、業務責任者、人手確認、既存業務との連携条件を確認する。",
            96,
        )
        basis = _front_basis(candidate.get("basis")) or "分析仮説"
        items.append({
            "priority": priority,
            "use_case_id": use_case_id,
            "use_case_no": case["use_case_no"],
            "theme": case["theme"],
            "reason": reason,
            "first_step": first_step,
            "depends_on": depends_on,
            "basis": basis,
            "architecture_implementation": _front_text((source_poc or {}).get("architecture_implementation"), 108),
        })
    portfolio = normalize_poc_portfolio(
        {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": items}, assessment, front,
    )
    return portfolio or {"schema_version": POC_PORTFOLIO_SCHEMA_VERSION, "items": []}


def poc_portfolio_for(assessment: dict, front_matter: dict | None = None) -> dict:
    """明示されたポートフォリオを優先し、旧JSONでは候補表から安全に導出する。"""
    normalized = normalize_poc_portfolio(assessment.get("poc_portfolio"), assessment, front_matter)
    return normalized or fallback_poc_portfolio(assessment, front_matter)


def normalize_poc_charters(raw: object, assessment: dict,
                           front_matter: dict | None = None) -> dict:
    """優先PoCを判断可能なチャーターとして検証・正規化する。

    チャーターは候補名の言い換えではなく、いつ何を判断するか、どのデータを
    どの条件で評価するか、誰が責任を持つかを保持する。実績がない項目は空欄に
    せず ``confirm`` を必須値として保持するため、顧客へ未確認事項を隠さない。
    """
    if (not isinstance(raw, dict)
            or _front_text(raw.get("schema_version"), 8) != POC_CHARTER_SCHEMA_VERSION):
        return {}
    portfolio = poc_portfolio_for(assessment, front_matter)
    expected_items = portfolio.get("items") if isinstance(portfolio, dict) else []
    charters_raw = raw.get("charters")
    if not isinstance(expected_items, list) or len(expected_items) != 3:
        return {}
    if not isinstance(charters_raw, list) or len(charters_raw) != 3:
        return {}

    required_fields = {
        "priority", "use_case_id", "theme", "decision_moment", "scope", "data_period", "data_volume",
        "missingness", "ground_truth", "evaluator", "comparator", "baseline", "success_criteria",
        "stop_criteria", "owners", "productization_decision", "basis",
    }
    owner_fields = ("business_owner", "data_owner", "product_owner", "technical_owner", "approval_owner")
    normalized: list[dict[str, object]] = []
    for expected, item in zip(expected_items, charters_raw):
        if not isinstance(expected, dict) or not isinstance(item, dict) or not required_fields <= item.keys():
            return {}
        owners_raw = item.get("owners")
        if not isinstance(owners_raw, dict) or not all(field in owners_raw for field in owner_fields):
            return {}
        row: dict[str, object] = {
            "priority": _front_text(item.get("priority"), 10),
            "use_case_id": _front_text(item.get("use_case_id"), 8),
            "theme": _front_text(item.get("theme"), 72),
            "decision_moment": _front_confirm_or_text(item.get("decision_moment"), 108),
            "scope": _front_confirm_or_text(item.get("scope"), 128),
            "data_period": _front_confirm_or_text(item.get("data_period"), 52),
            "data_volume": _front_confirm_or_text(item.get("data_volume"), 52),
            "missingness": _front_confirm_or_text(item.get("missingness"), 68),
            "ground_truth": _front_confirm_or_text(item.get("ground_truth"), 82),
            "evaluator": _front_confirm_or_text(item.get("evaluator"), 52),
            "comparator": _front_confirm_or_text(item.get("comparator"), 82),
            "baseline": _front_confirm_or_text(item.get("baseline"), 82),
            "success_criteria": _front_confirm_or_text(item.get("success_criteria"), 96),
            "stop_criteria": _front_confirm_or_text(item.get("stop_criteria"), 96),
            "owners": {
                field: _front_confirm_or_text(owners_raw.get(field), 52)
                for field in owner_fields
            },
            "productization_decision": _front_confirm_or_text(item.get("productization_decision"), 96),
            "basis": _front_basis(item.get("basis")),
        }
        if (not all(row[key] for key in ("priority", "use_case_id", "theme", "decision_moment", "scope",
                                         "data_period", "data_volume", "missingness", "ground_truth", "evaluator",
                                         "comparator", "baseline", "success_criteria", "stop_criteria",
                                         "productization_decision", "basis"))
                or row["priority"] != _front_text(expected.get("priority"), 10)
                or row["use_case_id"] != _front_text(expected.get("use_case_id"), 8)
                or normalized_use_case_label(row["theme"]) != normalized_use_case_label(expected.get("theme"))):
            return {}
        normalized.append(row)
    return {"schema_version": POC_CHARTER_SCHEMA_VERSION, "charters": normalized}


def fallback_poc_charters(assessment: dict, front_matter: dict | None = None) -> dict:
    """既存入力だけでは分からないPoC条件を、明示的な確認項目として補う。

    会社・業界に依存した数値、対象件数、責任者を創作せず、既存の優先PoCと
    ``confirm`` を用いる。これにより、旧JSONでも安全なレビュー用チャーターを
    出力できる。
    """
    pocs = priority_pocs_for(assessment, front_matter)
    if len(pocs) != 3:
        return {"schema_version": POC_CHARTER_SCHEMA_VERSION, "charters": []}
    charters: list[dict[str, object]] = []
    for item in pocs:
        theme = _front_text(item.get("theme"), 72)
        priority = _front_text(item.get("priority"), 10)
        use_case_id = _front_text(item.get("use_case_id"), 8)
        if not theme or priority not in {"P1", "P2", "P3"} or not use_case_id:
            return {"schema_version": POC_CHARTER_SCHEMA_VERSION, "charters": []}
        charters.append({
            "priority": priority,
            "use_case_id": use_case_id,
            "theme": theme,
            "decision_moment": f"代表範囲で「{theme}」の有用性・統制条件を確認後、次段階への進行可否を判断する。",
            "scope": f"「{theme}」の対象業務・利用者・対象範囲をPoC開始時に確認する。",
            "data_period": POC_CONFIRM_VALUE,
            "data_volume": POC_CONFIRM_VALUE,
            "missingness": POC_CONFIRM_VALUE,
            "ground_truth": POC_CONFIRM_VALUE,
            "evaluator": POC_CONFIRM_VALUE,
            "comparator": POC_CONFIRM_VALUE,
            "baseline": POC_CONFIRM_VALUE,
            "success_criteria": POC_CONFIRM_VALUE,
            "stop_criteria": POC_CONFIRM_VALUE,
            "owners": {
                "business_owner": POC_CONFIRM_VALUE,
                "data_owner": POC_CONFIRM_VALUE,
                "product_owner": POC_CONFIRM_VALUE,
                "technical_owner": POC_CONFIRM_VALUE,
                "approval_owner": POC_CONFIRM_VALUE,
            },
            "productization_decision": "限定導入、追加検証、保留のいずれに進むかを確認する。",
            "basis": "要確認",
        })
    return {"schema_version": POC_CHARTER_SCHEMA_VERSION, "charters": charters}


def poc_charters_for(assessment: dict, front_matter: dict | None = None) -> dict:
    """明示されたPoCチャーターを優先し、旧JSONでは安全な確認項目を派生する。"""
    normalized = normalize_poc_charters(assessment.get("poc_charters"), assessment, front_matter)
    return normalized or fallback_poc_charters(assessment, front_matter)


def normalize_multitenant_governance(raw: object) -> dict:
    """マルチテナントAIの必須統制を、欠落なく確認可能な構造へ正規化する。"""
    if (not isinstance(raw, dict)
            or _front_text(raw.get("schema_version"), 8) != MULTITENANT_GOVERNANCE_SCHEMA_VERSION):
        return {}
    blocks = {
        "boundaries": ("tenant", "shipper", "warehouse", "user"),
        "rag_documents": ("permission", "version", "delete", "prompt_injection"),
        "data_lifecycle": ("retention", "audit", "residency"),
        "answer_controls": ("responsibility", "human_approval", "rollback"),
        "model_operations": ("update", "drift", "monitoring", "stop"),
    }
    limits = {
        "boundaries": 72, "rag_documents": 96, "data_lifecycle": 82,
        "answer_controls": 82, "model_operations": 82,
    }
    normalized: dict[str, object] = {"schema_version": MULTITENANT_GOVERNANCE_SCHEMA_VERSION}
    for block, fields in blocks.items():
        raw_block = raw.get(block)
        if not isinstance(raw_block, dict) or not all(field in raw_block for field in fields):
            return {}
        normalized[block] = {
            field: _front_confirm_or_text(raw_block.get(field), limits[block])
            for field in fields
        }
    basis = _front_basis(raw.get("basis"))
    if not basis:
        return {}
    normalized["basis"] = basis
    return normalized


def fallback_multitenant_governance() -> dict:
    """未確認のテナント統制を仮定せず、設計・合意が必要な項目を残す。"""
    return {
        "schema_version": MULTITENANT_GOVERNANCE_SCHEMA_VERSION,
        "boundaries": {field: POC_CONFIRM_VALUE for field in ("tenant", "shipper", "warehouse", "user")},
        "rag_documents": {field: POC_CONFIRM_VALUE for field in ("permission", "version", "delete", "prompt_injection")},
        "data_lifecycle": {field: POC_CONFIRM_VALUE for field in ("retention", "audit", "residency")},
        "answer_controls": {field: POC_CONFIRM_VALUE for field in ("responsibility", "human_approval", "rollback")},
        "model_operations": {field: POC_CONFIRM_VALUE for field in ("update", "drift", "monitoring", "stop")},
        "basis": "要確認",
    }


def multitenant_governance_for(assessment: dict) -> dict:
    """顧客入力済みの統制方針を優先し、なければ確認台帳を返す。"""
    normalized = normalize_multitenant_governance(assessment.get("multitenant_governance"))
    return normalized or fallback_multitenant_governance()


def materialize_poc_decision_data(assessment: dict, front_matter: dict | None = None) -> dict:
    """JSONレビューと描画が同じPoCチャーター・統制データを参照するようにする。"""
    charters = poc_charters_for(assessment, front_matter)
    governance = multitenant_governance_for(assessment)
    assessment["poc_charters"] = charters
    assessment["multitenant_governance"] = governance
    return {"poc_charters": charters, "multitenant_governance": governance}


def poc_score_for_level(level: object) -> int | str:
    """既存の高・中・低評価を、比較用の 5/3/1 仮説へだけ変換する。"""
    return {"high": 5, "medium": 3, "low": 1}.get(str(level), POC_CONFIRM_VALUE)




def normalize_poc_selection_scorecard(raw: object, assessment: dict,
                                      front_matter: dict | None = None) -> dict:
    """優先3テーマの評価仮説と開始ゲートを、P1と同じIDで検証する。"""
    if not isinstance(raw, dict):
        return {}
    schema_version = _front_text(raw.get("schema_version"), 8)
    if schema_version not in {
        POC_SELECTION_SCORECARD_LEGACY_SCHEMA_VERSION, POC_SELECTION_SCORECARD_SCHEMA_VERSION,
    }:
        return {}
    items = raw.get("items")
    start_gates = raw.get("start_gates")
    if not isinstance(items, list) or len(items) != 3 or not isinstance(start_gates, dict):
        return {}
    pocs = priority_pocs_for(assessment, front_matter)
    if len(pocs) != 3:
        return {}
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    prioritization = front.get("use_case_prioritization") if isinstance(front, dict) else {}
    candidates = prioritization.get("candidates") if isinstance(prioritization, dict) else []
    candidate_by_no = {
        _front_text(item.get("use_case_no"), 3): item
        for item in candidates if isinstance(item, dict)
    } if isinstance(candidates, list) else {}
    normalized_items: list[dict[str, object]] = []
    for raw_item, poc in zip(items, pocs):
        if not isinstance(raw_item, dict):
            return {}
        score_values = raw_item.get("scores")
        if not isinstance(score_values, dict):
            return {}
        scores: dict[str, int | str] = {}
        for field in POC_SELECTION_SCORE_FIELDS:
            value = score_values.get(field)
            if value not in POC_SELECTION_SCORE_VALUES and value != POC_CONFIRM_VALUE:
                return {}
            scores[field] = value
        candidate = candidate_by_no.get(_front_text(raw_item.get("use_case_no"), 3), {})
        raw_reasons = raw_item.get("score_reasons")
        if schema_version == POC_SELECTION_SCORECARD_SCHEMA_VERSION:
            if not isinstance(raw_reasons, dict):
                return {}
            score_reasons = {
                field: _front_text(raw_reasons.get(field), 96)
                for field in POC_SELECTION_SCORE_FIELDS
            }
            data_next_action = _front_text(raw_item.get("data_next_action"), 112)
            if not all(score_reasons.values()) or not data_next_action:
                return {}
        else:
            score_reasons = {
                field: fallback_candidate_score_reason(candidate, field)
                for field in POC_SELECTION_SCORE_FIELDS
            }
            data_next_action = fallback_candidate_data_next_action(candidate)
        row = {
            "priority": _front_text(raw_item.get("priority"), 8),
            "use_case_id": _front_text(raw_item.get("use_case_id"), 8),
            "use_case_no": _front_text(raw_item.get("use_case_no"), 3),
            "theme": _front_text(raw_item.get("theme"), 72),
            "scores": scores,
            "score_reasons": score_reasons,
            "data_next_action": data_next_action,
            "evidence_to_collect": _front_text(raw_item.get("evidence_to_collect"), 126),
            "basis": _front_basis(raw_item.get("basis")),
        }
        if (not row["theme"] or not row["evidence_to_collect"]
                or row["priority"] != _front_text(poc.get("priority"), 8)
                or row["use_case_id"] != _front_text(poc.get("use_case_id"), 8)
                or row["use_case_no"] != _front_text(poc.get("use_case_no"), 3)
                or normalized_use_case_label(row["theme"]) != normalized_use_case_label(poc.get("theme"))):
            return {}
        normalized_items.append(row)
    gates = {
        field: _front_confirm_or_text(start_gates.get(field), 104)
        for field in POC_SELECTION_START_GATE_FIELDS
    }
    decision = _front_text(raw.get("decision"), 140)
    basis = _front_basis(raw.get("basis"))
    if not decision or not basis:
        return {}
    return {
        "schema_version": POC_SELECTION_SCORECARD_SCHEMA_VERSION,
        "score_scale": "5/3/1/confirm",
        "items": normalized_items,
        "start_gates": gates,
        "decision": decision,
        "basis": basis,
    }


def fallback_poc_selection_scorecard(assessment: dict, front_matter: dict | None = None) -> dict:
    """既存候補から、実績を創作しない評価仮説・開始条件を導出する。"""
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    pocs = priority_pocs_for(assessment, front)
    candidates = (front.get("use_case_prioritization", {}).get("candidates", [])
                  if isinstance(front.get("use_case_prioritization"), dict) else [])
    candidate_by_no = {
        _front_text(item.get("use_case_no"), 3): item
        for item in candidates if isinstance(item, dict)
    }
    items: list[dict[str, object]] = []
    for poc in pocs[:3]:
        candidate = candidate_by_no.get(_front_text(poc.get("use_case_no"), 3), {})
        items.append({
            "priority": _front_text(poc.get("priority"), 8),
            "use_case_id": _front_text(poc.get("use_case_id"), 8),
            "use_case_no": _front_text(poc.get("use_case_no"), 3),
            "theme": _front_text(poc.get("theme"), 72),
            "scores": {
                "business_value": poc_score_for_level(candidate.get("value")),
                "feasibility": poc_score_for_level(candidate.get("feasibility")),
                "data_readiness": poc_score_for_level(candidate.get("data_readiness")),
                # 横展開・商用化の見込みは、テナント統制と提供形態の確認前には
                # 点数化しない。二層の事業価値モデル・統制ページで検証する。
                "scale_readiness": POC_CONFIRM_VALUE,
            },
            "score_reasons": {
                field: fallback_candidate_score_reason(candidate, field)
                for field in POC_SELECTION_SCORE_FIELDS
            },
            "data_next_action": fallback_candidate_data_next_action(candidate),
            "evidence_to_collect": _front_text(
                poc.get("first_step") or candidate.get("rationale")
                or "対象業務・代表データ・比較条件を確認する。", 126,
            ),
            "basis": _front_basis(poc.get("basis") or candidate.get("basis")),
        })
    # 完全な3候補を持たない旧JSONでは、表示側が安全に要確認へフォールバックする。
    if len(items) != 3:
        return {}
    return {
        "schema_version": POC_SELECTION_SCORECARD_SCHEMA_VERSION,
        "score_scale": "5/3/1/confirm",
        "items": items,
        "start_gates": {field: POC_CONFIRM_VALUE for field in POC_SELECTION_START_GATE_FIELDS},
        "decision": "開始ゲートを満たすテーマだけをPoCへ進め、実測結果で採用・追加検証・保留を判断する。",
        "basis": "分析仮説",
    }


def poc_selection_scorecard_for(assessment: dict, front_matter: dict | None = None) -> dict:
    """保存済みのスコアカードを優先し、旧JSONには安全な評価仮説を使う。"""
    normalized = normalize_poc_selection_scorecard(
        assessment.get("poc_selection_scorecard"), assessment, front_matter,
    )
    return normalized or fallback_poc_selection_scorecard(assessment, front_matter)


def materialize_poc_selection_scorecard(assessment: dict, front_matter: dict | None = None) -> dict:
    """JSONレビュー・付録Bの評価ページを同じ開始ゲートへ固定する。"""
    scorecard = poc_selection_scorecard_for(assessment, front_matter)
    if scorecard:
        assessment["poc_selection_scorecard"] = scorecard
    return scorecard


def priority_decision_for(assessment: dict, front_matter: dict | None = None) -> dict:
    """候補全件の優先化根拠を、JSON上でレビュー可能な形に固定する。"""
    if isinstance(front_matter, dict):
        # 呼び出し元が保存前のLLM応答を直接渡した場合も、表示時と同じ評価・
        # 技術分散を適用する。未正規化のP1〜P3を契約へ固定しない。
        front = normalize_consulting_front_matter(front_matter, assessment) or front_matter
    else:
        front = consulting_front_matter_for(assessment)
    prioritization = front.get("use_case_prioritization") if isinstance(front, dict) else {}
    candidates = prioritization.get("candidates") if isinstance(prioritization, dict) else []
    if not isinstance(candidates, list) or len(candidates) != 15:
        return {}
    items: list[dict[str, object]] = []
    for candidate in candidates:
        if not isinstance(candidate, dict):
            return {}
        evaluation = candidate_priority_evaluation(candidate)
        use_case_no = _front_text(candidate.get("use_case_no"), 3)
        try:
            use_case_id = f"UC{int(use_case_no):02d}"
        except ValueError:
            return {}
        items.append({
            "use_case_id": use_case_id,
            "use_case_no": use_case_no,
            "priority": _front_text(candidate.get("priority"), 10),
            "weighted_score": evaluation["score"] if evaluation["score"] is not None else POC_CONFIRM_VALUE,
            "ranking_index": evaluation["ranking_index"],
            "evaluation_coverage": evaluation["evaluation_coverage"],
            "eligible_for_priority": evaluation["eligible_for_priority"],
            "eligibility_reasons": evaluation["eligibility_reasons"],
            "technical_modality": _front_text(candidate.get("technical_modality"), 40),
            "dimensions": evaluation["dimensions"],
            "basis": _front_basis(candidate.get("basis")),
        })
    selected_priorities = sorted(
        (item["priority"] for item in items if item["priority"] in {"P1", "P2", "P3"}),
    )
    if selected_priorities != ["P1", "P2", "P3"]:
        return {}
    legacy_override = legacy_priority_override_for(assessment)
    selection_methods = {
        _front_text(candidate.get("priority_selection_method"), 40)
        for candidate in candidates
        if candidate.get("priority") in {"P1", "P2", "P3"}
    }
    selection_method = (
        selection_methods.pop() if len(selection_methods) == 1
        else "technology_portfolio_v1"
    )
    selected_modalities = {
        _front_text(item.get("technical_modality"), 40)
        for item in items if item.get("priority") in {"P1", "P2", "P3"}
        and _front_text(item.get("technical_modality"), 40)
    }
    return {
        "schema_version": POC_PRIORITY_DECISION_SCHEMA_VERSION,
        "selection_method": selection_method,
        "weights": POC_PRIORITY_SCORE_WEIGHTS.copy(),
        "policy": {
            "minimum_coverage": POC_PRIORITY_MINIMUM_COVERAGE,
            "required_dimensions": list(POC_PRIORITY_REQUIRED_DIMENSIONS),
            "minimum_data_readiness": 3,
            "provisional_fill": True,
            "legacy_priority_preserved": bool(legacy_override),
            "legacy_priority_override": legacy_override,
            "technology_diversity": {
                "distinct_modality_first": selection_method == "technology_portfolio_v1",
                "selected_modality_count": len(selected_modalities),
                "fallback_when_fewer_than_three": "評価順位で補完",
            },
        },
        "tie_break_rule": "適格候補を優先し、順位指数、評価充足率、データ準備度の順で比較する。完全同点時だけユースケース番号順で暫定配置する。",
        "items": items,
    }


def normalize_poc_priority_decision(raw: object, assessment: dict,
                                    front_matter: dict | None = None) -> dict:
    """保存済みの優先化根拠が、現在の候補評価式と一致するか検証する。"""
    expected = priority_decision_for(assessment, front_matter)
    if not expected or not isinstance(raw, dict):
        return {}
    schema_version = _front_text(raw.get("schema_version"), 8)
    if schema_version not in {
        POC_PRIORITY_DECISION_LEGACY_SCHEMA_VERSION,
        POC_PRIORITY_DECISION_PREVIOUS_SCHEMA_VERSION,
        POC_PRIORITY_DECISION_SCHEMA_VERSION,
    }:
        return {}
    old_method = (
        schema_version in {POC_PRIORITY_DECISION_LEGACY_SCHEMA_VERSION,
                           POC_PRIORITY_DECISION_PREVIOUS_SCHEMA_VERSION}
        and raw.get("selection_method") in {"weighted_model", "weighted_model_v2"}
    )
    if (not old_method and raw.get("selection_method") != expected["selection_method"]) or raw.get("weights") != expected["weights"]:
        return {}
    if (schema_version == POC_PRIORITY_DECISION_SCHEMA_VERSION
            and raw.get("policy") != expected["policy"]):
        return {}
    raw_items = raw.get("items")
    if not isinstance(raw_items, list) or len(raw_items) != len(expected["items"]):
        return {}
    for actual, anticipated in zip(raw_items, expected["items"]):
        if not isinstance(actual, dict):
            return {}
        common_fields = (
            "use_case_id", "use_case_no", "weighted_score", "evaluation_coverage", "basis",
        )
        if not old_method:
            common_fields += ("priority",)
        if any(actual.get(field) != anticipated[field] for field in common_fields):
            return {}
        if schema_version == POC_PRIORITY_DECISION_LEGACY_SCHEMA_VERSION:
            actual_dimensions = actual.get("dimensions")
            if not isinstance(actual_dimensions, dict):
                return {}
            legacy_expected = {
                field: {
                    key: anticipated["dimensions"][field][key]
                    for key in ("level", "score", "weight")
                }
                for field in POC_SELECTION_SCORE_FIELDS
            }
            if actual_dimensions != legacy_expected:
                return {}
        elif schema_version == POC_PRIORITY_DECISION_PREVIOUS_SCHEMA_VERSION:
            if any(actual.get(field) != anticipated[field] for field in (
                "ranking_index", "eligible_for_priority", "eligibility_reasons", "dimensions",
            )):
                return {}
        elif schema_version == POC_PRIORITY_DECISION_SCHEMA_VERSION:
            if any(actual.get(field) != anticipated[field] for field in (
                "ranking_index", "eligible_for_priority", "eligibility_reasons",
                "technical_modality", "dimensions",
            )):
                return {}
    return expected


def materialize_poc_priority_decision(assessment: dict, front_matter: dict | None = None) -> dict:
    """評価式・重み・充足率を、表示順位と同じレビューJSONへ保存する。"""
    decision = priority_decision_for(assessment, front_matter)
    if decision:
        assessment["poc_priority_decision"] = decision
    return decision


def _poc_start_gate_values(charter: dict, governance: dict) -> dict[str, list[object]]:
    """PoCの開始判断に必要な値を、チャーターと統制台帳から集約する。"""
    owners = charter.get("owners") if isinstance(charter.get("owners"), dict) else {}
    answer_controls = governance.get("answer_controls") if isinstance(governance, dict) else {}
    return {
        "scope_and_population": [charter.get("scope")],
        "data_contract_and_boundary": [
            charter.get("data_period"), charter.get("data_volume"), charter.get("missingness"),
        ],
        "baseline_and_comparator": [
            charter.get("ground_truth"), charter.get("evaluator"), charter.get("comparator"),
            charter.get("baseline"), charter.get("success_criteria"), charter.get("stop_criteria"),
        ],
        "human_approval_and_fallback": [
            answer_controls.get("human_approval"), answer_controls.get("rollback"),
        ],
        "decision_owner": [owners.get(field) for field in (
            "business_owner", "data_owner", "product_owner", "technical_owner", "approval_owner",
        )],
    }


def _poc_start_approval_for(assessment: dict, use_case_id: str) -> dict:
    """明示された開始承認だけを採用し、推測でapprovedへ進めない。"""
    approvals = assessment.get("poc_start_approvals")
    if not isinstance(approvals, dict):
        return {}
    record = approvals.get(use_case_id)
    if not isinstance(record, dict):
        return {}
    if (record.get("status") == "approved" and _front_text(record.get("approved_by"), 80)
            and _front_text(record.get("approved_at"), 40)):
        return record
    return {}


def poc_start_readiness_for(assessment: dict, front_matter: dict | None = None) -> dict:
    """P1〜P3ごとに開始可否・未充足ゲート・次アクションを明示する。"""
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    portfolio = priority_pocs_for(assessment, front)
    charters = poc_charters_for(assessment, front).get("charters", [])
    governance = multitenant_governance_for(assessment)
    if not isinstance(charters, list) or len(portfolio) != 3 or len(charters) != 3:
        return {}
    next_action_by_gate = {
        "scope_and_population": "対象業務・利用者・母集団を確定する。",
        "data_contract_and_boundary": "対象データ、品質、利用範囲・境界を確認する。",
        "baseline_and_comparator": "基準値・比較条件・評価・停止基準を合意する。",
        "human_approval_and_fallback": "人手承認と根拠不足時の復帰手順を確認する。",
        "decision_owner": "業務・データ・技術・承認の責任者を確定する。",
    }
    items: list[dict[str, object]] = []
    for poc, charter in zip(portfolio, charters):
        if not isinstance(poc, dict) or not isinstance(charter, dict):
            return {}
        gate_values = _poc_start_gate_values(charter, governance)
        open_gate_ids = [
            gate_id for gate_id, values in gate_values.items()
            if any(_front_confirm_or_text(value, 120) == POC_CONFIRM_VALUE for value in values)
        ]
        approval_owner = _front_confirm_or_text(
            (charter.get("owners") or {}).get("approval_owner") if isinstance(charter.get("owners"), dict) else "",
            80,
        )
        approval = _poc_start_approval_for(assessment, _front_text(poc.get("use_case_id"), 8))
        if open_gate_ids:
            status = "blocked"
        elif approval:
            status = "approved"
        else:
            status = "ready_for_decision"
        items.append({
            "priority": _front_text(poc.get("priority"), 8),
            "use_case_id": _front_text(poc.get("use_case_id"), 8),
            "theme": _front_text(poc.get("theme"), 72),
            "status": status,
            "open_gate_ids": open_gate_ids,
            "open_gate_count": len(open_gate_ids),
            "next_action": next_action_by_gate.get(open_gate_ids[0], "開始承認を行う。"),
            "decision_owner": approval_owner if approval_owner != POC_CONFIRM_VALUE else POC_CONFIRM_VALUE,
            "basis": _front_basis(charter.get("basis")) or "要確認",
        })
    return {
        "schema_version": POC_START_READINESS_SCHEMA_VERSION,
        "gate_ids": list(POC_SELECTION_START_GATE_FIELDS),
        "items": items,
    }


def normalize_poc_start_readiness(raw: object, assessment: dict,
                                  front_matter: dict | None = None) -> dict:
    """保存済みの開始可否が、実際の未充足条件と一致するか検証する。"""
    expected = poc_start_readiness_for(assessment, front_matter)
    if not expected or not isinstance(raw, dict):
        return {}
    if (_front_text(raw.get("schema_version"), 8) != POC_START_READINESS_SCHEMA_VERSION
            or raw.get("gate_ids") != expected["gate_ids"]):
        return {}
    raw_items = raw.get("items")
    if not isinstance(raw_items, list) or len(raw_items) != 3:
        return {}
    for actual, anticipated in zip(raw_items, expected["items"]):
        if not isinstance(actual, dict) or actual != anticipated:
            return {}
    return expected


def materialize_poc_start_readiness(assessment: dict, front_matter: dict | None = None) -> dict:
    """PoCの優先順位と開始可否を別概念としてレビューJSONへ保持する。"""
    readiness = poc_start_readiness_for(assessment, front_matter)
    if readiness:
        assessment["poc_start_readiness"] = readiness
    return readiness


def priority_pocs_for(assessment: dict, front_matter: dict | None = None) -> list[dict]:
    """全描画経路が同じ優先PoCを使うための唯一の参照点。"""
    portfolio = poc_portfolio_for(assessment, front_matter)
    if len(portfolio.get("items", [])) == 3:
        return [item.copy() for item in portfolio["items"]]
    # 15件の旧候補を持たない簡易テスト・後方互換JSONでは、従来の表示を維持する。
    raw = assessment.get("poc_recommendations")
    return [item.copy() for item in raw[:3] if isinstance(item, dict)] if isinstance(raw, list) else []




def technical_modality_for(poc: dict, detail: dict | None = None) -> dict[str, str]:
    """優先PoCの主処理を、RAG・ML・最適化などで混同せずに表現する。

    生成AIを万能な処理として描かないため、根拠検索、予測・異常検知、制約付き
    優先化は別々の入出力・評価観点を持つモダリティとして判定する。
    """
    detail = detail or {}
    # まずテーマ名と実現ロジックの主題で処理方式を定める。選定理由に他テーマの
    # 例示語が混じっても、優先順位支援をRAGのように誤表示しないためである。
    primary_text = " ".join(str(value or "") for value in (
        poc.get("theme"), detail.get("theme"), detail.get("business_challenge"),
    )).lower()
    supporting_text = " ".join(str(value or "") for value in (
        poc.get("reason"), poc.get("first_step"), detail.get("target_data"), detail.get("oci_roles"),
    )).lower()
    text = f"{primary_text} {supporting_text}"
    modality_text = primary_text or text
    if any(keyword in modality_text for keyword in ("問い合わせ", "faq", "rag", "ナレッジ", "検索", "回答", "チケット")):
        return {
            "kind": "rag",
            "label": "RAG（根拠検索＋回答案）",
            "candidate_label": "根拠付きの回答・切り分け候補",
            "evidence_label": "引用箇所・文書版・参照範囲",
            "processing_note": "Autonomous AI DatabaseのVector Searchで許可済み文書を検索し、OCI Generative AIは検索根拠に限定して回答案を整える。",
        }
    if any(keyword in modality_text for keyword in ("異常", "予測", "forecast", "anomaly", "分類", "classif", "スコア", "検知", "回帰")):
        return {
            "kind": "anomaly_ml",
            "label": "異常検知・予測／ML",
            "candidate_label": "確認すべき候補・スコア",
            "evidence_label": "判定条件・関連履歴・スコア",
            "processing_note": "Autonomous AI Database上のOracle Machine Learning（OML）で履歴と条件から候補・スコアを算出し、OCI Generative AIは確認観点を説明する。",
        }
    if any(keyword in modality_text for keyword in (
        "優先順位", "優先提案", "作業順", "順序", "最適化", "optimization",
        "配分", "経路", "ルート", "配置", "割当", "割付", "スケジュー",
    )):
        return {
            "kind": "optimization",
            "label": "最適化・優先順位付け",
            "candidate_label": "制約を考慮した順位・配分候補",
            "evidence_label": "期限・制約・入力条件・順位理由",
            "processing_note": "業務ルールと利用可能な条件を基に候補順位を算出し、OCI Generative AIは順位理由と利用者が確認すべき制約を説明する。",
        }
    if any(keyword in modality_text for keyword in ("帳票", "ocr", "文書", "画像", "書類", "読取")):
        return {
            "kind": "document_understanding",
            "label": "文書理解・抽出",
            "candidate_label": "抽出値・照合候補",
            "evidence_label": "原文位置・照合結果・信頼度",
            "processing_note": "文書・画像から対象項目を抽出して業務データと照合し、低信頼または不一致の項目を人が確認する。",
        }
    return {
        "kind": "generative_assist",
        "label": "生成AIによる要約・提案支援",
        "candidate_label": "業務上の確認・提案候補",
        "evidence_label": "入力条件・参照情報・生成理由",
        "processing_note": "対象データと業務ルールを参照して候補を生成し、根拠不足時は人の確認または既存フローへ戻す。",
    }


def _technical_field_text(value: object, fallback: object, limit: int) -> str:
    """LLMが文字列・配列のどちらで返しても、提案書向けの通常表記へ揃える。"""
    selected = value or fallback
    if isinstance(selected, (list, tuple, set)):
        parts = [_front_text(item, limit) for item in selected if _front_text(item, limit)]
        selected = "／".join(parts)
    return _front_text(selected, limit)


def _oracle_technology_text(value: object, fallback: object) -> str:
    """正式名の表記揺れを、スライドで読めるOracle機能名へ正規化する。"""
    text = _technical_field_text(value, fallback, 240)
    replacements = {
        "Oracle Machine Learning for SQL": "Oracle Machine Learning（OML4SQL）",
        "Oracle Database Analytic Functions": "SQL分析関数",
        "Oracle REST Data Services": "ORDS・REST API",
        "Oracle Scheduler": "DBMS_SCHEDULER",
        "Oracle Database 23ai": "Oracle AI Database 26ai",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = re.sub(r"12\s*[〜～~-]\s*24か月", "対象履歴期間", text)
    text = re.sub(r"4\s*[〜～~-]\s*12週間", "対象評価期間", text)
    text = re.sub(r"数百\s*[〜～~-]\s*数千文書", "対象文書", text)
    text = text.replace("過去対象履歴期間を目安に", "利用可能な履歴期間を確認し")
    text = text.replace("対象評価期間分扱います", "評価対象期間のデータを扱います")
    text = text.replace("対象文書規模で扱い", "利用可能な対象文書を扱い")
    return text


def display_theme_name(detail: dict, poc: dict) -> str:
    """実装方式より強い効果を題名で約束しない表示名へ整える。"""
    theme = _front_text(detail.get("theme") or poc.get("theme"), 72)
    modality = technical_modality_for(poc, detail).get("kind")
    implementation = _technical_field_text(
        detail.get("oracle_technologies") or detail.get("implementation_summary"), "", 320,
    )
    has_solver = bool(re.search(
        r"最適化ソルバー|数理最適化|線形計画|整数計画|constraint solver",
        implementation, re.I,
    ))
    if modality == "optimization" and "最適化" in theme and not has_solver:
        return theme.replace("最適化", "提案")
    return theme


def _polish_technical_display_text(value: object) -> str:
    """技術欄に残りやすい未完文と接続手段の混同を表示前に整える。"""
    text = _front_text(value, 360)
    replacements = {
        "Oracle Database 23ai": "Oracle AI Database 26ai",
        "Database Link・APIで返却": "Database Linkで参照、またはAPIで返却",
        "Database Link・APIでオフロード": "検証データ複製、Database Linkでの参照、またはAPI連携でオフロード",
        "信頼表示": "引用充足状況・人手確認状態",
        "DB内実行を選び。": "DB内実行を選択します。",
        "軽量なSQL処理は既存Oracle Database。": "軽量なSQL処理は既存Oracle Database内で実行します。",
        "Vector Search対応かを確認し。": "Vector Searchへの対応状況を確認します。",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    text = re.sub(r"12\s*[〜～~-]\s*24か月", "対象履歴期間", text)
    text = re.sub(r"4\s*[〜～~-]\s*12週間", "対象評価期間", text)
    text = re.sub(r"数百\s*[〜～~-]\s*数千文書", "対象文書", text)
    text = text.replace("過去対象履歴期間を目安に", "利用可能な履歴期間を確認し")
    text = text.replace("対象評価期間分扱います", "評価対象期間のデータを扱います")
    text = text.replace("対象文書規模で扱い", "利用可能な対象文書を扱い")
    return text


def poc_technical_design_for(detail: dict, poc: dict) -> dict[str, str]:
    """優先PoC詳細を、案件非依存のOracle実装仕様へ正規化する。

    新しいJSONが技術項目を返す場合はその内容を使う。旧JSONではテーマの
    modalityから安全な候補機能とDBオブジェクトを補い、会社・製品固有の
    テーブル名や導入済み機能は推測しない。
    """
    modality = technical_modality_for(poc, detail)
    defaults = {
        "optimization": {
            "oracle_technologies": "Oracle Machine Learning（OML4SQL）／SQL制約ロジック／DBMS_SCHEDULER",
            "database_objects": "入力・特徴量ビュー／スコア・制約表／順位候補・監査ログ",
        },
        "anomaly_ml": {
            "oracle_technologies": "Oracle Machine Learning（OML4SQL）／SQL分析／DBMS_SCHEDULER",
            "database_objects": "入力・特徴量ビュー／学習・スコア表／判定結果・監査ログ",
        },
        "rag": {
            "oracle_technologies": "Oracle AI Vector Search／Select AI（DBMS_CLOUD_AI）／OCI Generative AI",
            "database_objects": "文書・チャンク表／VECTOR列・ベクトル索引／回答・参照ログ",
        },
        "document_understanding": {
            "oracle_technologies": "OCI Generative AI／Autonomous AI Database JSON／ORDS・REST API",
            "database_objects": "原本・受付表／抽出JSON・照合表／修正・監査ログ",
        },
        "generative_assist": {
            "oracle_technologies": "Select AI（DBMS_CLOUD_AI）／OCI Generative AI／ORDS・REST API",
            "database_objects": "入力・参照ビュー／生成結果表／承認・修正・監査ログ",
        },
    }
    selected = defaults.get(modality["kind"], defaults["generative_assist"])
    boundary = (
        "既存Oracle Database内で実行、または検証データをAutonomous AI Databaseへ複製／"
        "Database Link・APIでオフロードします。DBバージョン、データ量、権限、遅延要件で選択します。"
    )
    proposed_objects = _technical_field_text(detail.get("database_objects"), "", 320)
    # LLMが未確認の物理表名を仮置きした場合は表示しない。ここでは実装時に
    # 確定する論理オブジェクトの役割だけを示す。
    if re.search(r"(?:^|[、／,\s])[A-Z][A-Z0-9]*_[A-Z0-9_]+", proposed_objects):
        proposed_objects = ""
    return {
        "implementation_summary": _polish_technical_display_text(_technical_field_text(
            detail.get("implementation_summary"), modality["processing_note"], 180,
        )),
        "input_data": _polish_technical_display_text(
            _technical_field_text(detail.get("target_data"), "", 160)
        ),
        "oracle_technologies": _polish_technical_display_text(_oracle_technology_text(
            detail.get("oracle_technologies"), selected["oracle_technologies"],
        )),
        "database_objects": _technical_field_text(
            proposed_objects, selected["database_objects"], 160,
        ),
        "output_interface": _polish_technical_display_text(_technical_field_text(
            detail.get("output_interface"), detail.get("business_integration"), 180,
        )),
        "implementation_boundary": _polish_technical_display_text(_technical_field_text(
            detail.get("implementation_boundary"), boundary, 220,
        )),
        "control_design": _technical_field_text(
            detail.get("control_design"), detail.get("design_notes"), 180,
        ),
        "validation_plan": _polish_technical_display_text(_technical_field_text(
            detail.get("validation_plan"), "現行業務との比較条件、品質、時間、利用者確認を評価します。", 180,
        )),
    }


def _technical_primary_detail(assessment: dict, primary_poc: dict) -> dict:
    """P1と同一ID・テーマの実現ロジックだけを技術提案へ利用する。"""
    for detail in poc_logic_details_for(assessment):
        if _matching_source_poc(detail, primary_poc):
            return detail
    return {}


def build_default_technical_proposal(assessment: dict) -> dict:
    """顧客固有の未確認仕様を創作しない、状態付きの技術提案データを作る。

    入力から確認できる対象サービスだけを ``confirmed`` とし、連携・認証・
    テナント境界・運用条件は ``verify``、OCI上の実装案は ``proposed`` に分ける。
    これによりPPTXの構成図とレビューJSONが同じ根拠状態を共有する。
    """
    priority_pocs = priority_pocs_for(assessment)
    primary = priority_pocs[0] if priority_pocs else {}
    primary_detail = _technical_primary_detail(assessment, primary) if primary else {}
    service_name = _front_text(assessment.get("service_name"), 72) or "対象サービス"
    theme = _front_text(primary.get("theme"), 72) or "優先PoCテーマ（要選定）"
    target_data = _front_text(primary_detail.get("target_data"), 120)
    modality = technical_modality_for(primary, primary_detail)
    canonical_poc = {
        "priority": _front_text(primary.get("priority"), 8) or "P1",
        "use_case_id": _front_text(primary.get("use_case_id"), 8),
        "theme": theme,
        "state": "proposed",
    }
    nodes = [
        {
            "id": "source_system", "state": "confirmed",
            "label": f"{service_name}の対象業務システム",
            "description": (
                f"{target_data}を{theme}のPoC対象として整理する。既存の連携仕様そのものは未確認。"
                if target_data else
                f"{service_name}で扱う業務とデータを、{theme}のPoC対象として整理する。既存の連携仕様そのものは未確認。"
            ),
        },
        {
            "id": "event_api", "state": "verify",
            "label": "業務イベント・API・定期連携",
            "description": "イベント、API、定期連携のどれを使うか、更新頻度、失敗時の再送・重複防止をPoCで確認する。",
        },
        {
            "id": "integration", "state": "proposed",
            "label": "業務連携・アプリケーション層",
            "description": "テナントID・利用者ロールを確認して入出力を制御し、既存画面またはAPIへ候補と根拠を返す。",
        },
        {
            "id": "data_layer", "state": "proposed",
            "label": "Autonomous AI Database",
            "description": "対象レコード、参照情報、判定結果、利用者の採否・修正、監査に必要な記録をテナント単位で管理する。",
        },
        {
            "id": "ai_processing", "state": "proposed",
            "label": modality["label"],
            "description": modality["processing_note"],
        },
        {
            "id": "human_action", "state": "proposed",
            "label": "人による確認・既存画面/APIへの組込み",
            "description": "担当者が候補・根拠・適用範囲を確認し、採用・修正・保留・エスカレーションを選択する。",
        },
        {
            "id": "record_and_audit", "state": "proposed",
            "label": "判断記録・監査",
            "description": "最終判断、修正理由、例外、参照範囲を既存業務の記録と評価台帳へ残し、再評価に利用する。",
        },
    ]
    controls = [
        {
            "id": "identity_tenant", "state": "verify",
            "label": "ID・権限・テナント境界",
            "description": "テナント識別子の取得元、利用者ロール、検索・表示・ログ出力の分離を権限別テストで確認する。",
        },
        {
            "id": "data_quality", "state": "verify",
            "label": "データ品質・連携条件",
            "description": "キー項目、欠損、更新遅延、対象範囲、異常・正解データの定義をPoC開始前に確認する。",
        },
        {
            "id": "fallback", "state": "proposed",
            "label": "根拠不足時のフォールバック",
            "description": "根拠不足、低信頼、権限外、連携失敗時は自動更新を行わず、既存業務フローまたは担当者へ戻す。",
        },
        {
            "id": "operations", "state": "verify",
            "label": "運用・品質評価",
            "description": "応答時間、処理量、処理単価、利用者の修正率、例外率、モデル・文書更新の責任分界を確認する。",
        },
    ]
    demo_scene = {
        "state": "proposed",
        "title": f"概念ワイヤーフレーム：{theme}",
        "event_label": "業務イベント・対象レコード",
        "event_detail": "既存画面またはAPIから、PoCで許可された対象範囲だけを受け付ける。",
        "candidate_label": modality["candidate_label"],
        "candidate_detail": f"{theme}に関する候補を表示する。自動確定は行わない。",
        "evidence_label": modality["evidence_label"],
        "evidence_detail": "候補の理由、参照範囲、適用できない条件を利用者が確認できる形で示す。",
        "human_label": "担当者の確認・判断",
        "human_detail": "採用・修正・保留・エスカレーションを選び、判断理由を記録する。",
        "record_label": "既存業務システム・評価記録",
        "record_detail": "確定結果と例外を既存フローへ戻し、PoC評価に必要な記録を残す。",
        "fallback_label": "根拠不足・低信頼・権限外 → 既存フローへ戻す",
    }
    return {
        "schema_version": TECHNICAL_PROPOSAL_SCHEMA_VERSION,
        "canonical_poc": canonical_poc,
        "modality": modality,
        "nodes": nodes,
        "controls": controls,
        "demo_scene": demo_scene,
    }


def normalize_technical_proposal(raw: object, assessment: dict) -> dict:
    """レビューJSONに含まれる状態付き技術提案の契約を検証する。"""
    if not isinstance(raw, dict) or _front_text(raw.get("schema_version"), 8) != TECHNICAL_PROPOSAL_SCHEMA_VERSION:
        return {}
    canonical = raw.get("canonical_poc")
    modality = raw.get("modality")
    nodes = raw.get("nodes")
    controls = raw.get("controls")
    demo = raw.get("demo_scene")
    if not all(isinstance(value, dict) for value in (canonical, modality, demo)):
        return {}
    pocs = priority_pocs_for(assessment)
    primary = pocs[0] if pocs else {}
    expected_id = _front_text(primary.get("use_case_id"), 8)
    expected_theme = normalized_use_case_label(primary.get("theme"))
    row = {
        "priority": _front_text(canonical.get("priority"), 8),
        "use_case_id": _front_text(canonical.get("use_case_id"), 8),
        "theme": _front_text(canonical.get("theme"), 72),
        "state": _front_text(canonical.get("state"), 12),
    }
    if (not row["theme"] or row["state"] != "proposed"
            or (expected_id and row["use_case_id"] != expected_id)
            or (expected_theme and normalized_use_case_label(row["theme"]) != expected_theme)):
        return {}
    modality_row = {key: _front_text(modality.get(key), 160) for key in (
        "kind", "label", "candidate_label", "evidence_label", "processing_note",
    )}
    if not all(modality_row.values()) or modality_row["kind"] not in {
        "rag", "anomaly_ml", "optimization", "document_understanding", "generative_assist",
    }:
        return {}
    if not isinstance(nodes, list) or len(nodes) != len(TECHNICAL_PROPOSAL_NODE_IDS):
        return {}
    normalized_nodes: list[dict[str, str]] = []
    for expected_node_id, node in zip(TECHNICAL_PROPOSAL_NODE_IDS, nodes):
        if not isinstance(node, dict):
            return {}
        row = {key: _front_text(node.get(key), 180) for key in ("id", "state", "label", "description")}
        if (row["id"] != expected_node_id or row["state"] not in TECHNICAL_PROPOSAL_STATE_VALUES
                or not row["label"] or not row["description"]):
            return {}
        normalized_nodes.append(row)
    if not isinstance(controls, list) or len(controls) != len(TECHNICAL_PROPOSAL_CONTROL_IDS):
        return {}
    normalized_controls: list[dict[str, str]] = []
    for expected_control_id, control in zip(TECHNICAL_PROPOSAL_CONTROL_IDS, controls):
        if not isinstance(control, dict):
            return {}
        row = {key: _front_text(control.get(key), 180) for key in ("id", "state", "label", "description")}
        if (row["id"] != expected_control_id or row["state"] not in TECHNICAL_PROPOSAL_STATE_VALUES
                or not row["label"] or not row["description"]):
            return {}
        normalized_controls.append(row)
    demo_row = {key: _front_text(demo.get(key), 180) for key in (
        "state", "title", "event_label", "event_detail", "candidate_label", "candidate_detail",
        "evidence_label", "evidence_detail", "human_label", "human_detail", "record_label",
        "record_detail", "fallback_label",
    )}
    if demo_row["state"] != "proposed" or not all(value for key, value in demo_row.items() if key != "state"):
        return {}
    return {
        "schema_version": TECHNICAL_PROPOSAL_SCHEMA_VERSION,
        "canonical_poc": {
            "priority": _front_text(canonical.get("priority"), 8),
            "use_case_id": _front_text(canonical.get("use_case_id"), 8),
            "theme": _front_text(canonical.get("theme"), 72),
            "state": "proposed",
        },
        "modality": modality_row,
        "nodes": normalized_nodes,
        "controls": normalized_controls,
        "demo_scene": demo_row,
    }


def technical_proposal_for(assessment: dict) -> dict:
    """有効な保存済み提案を優先し、旧JSONからは安全な派生データを返す。"""
    normalized = normalize_technical_proposal(assessment.get("technical_proposal"), assessment)
    return normalized or build_default_technical_proposal(assessment)


def materialize_poc_portfolio(assessment: dict, front_matter: dict | None = None) -> dict:
    """最終JSONと全描画経路を、同じ3件の優先PoCへ収束させる。

    初期LLM応答、15件候補の優先化、構造化分析は別工程で生成される。そのままでは
    表示名と表示順だけで結合され、別テーマの理由・検証条件が混ざり得る。ここで
    ``use_case_id`` をキーにしたポートフォリオを確定し、旧形式の各派生表も同じ
    内容へ更新する。
    """
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    portfolio = poc_portfolio_for(assessment, front)
    items = portfolio.get("items", [])
    if len(items) != 3:
        assessment["technical_proposal"] = build_default_technical_proposal(assessment)
        return portfolio
    assessment["poc_portfolio"] = portfolio
    assessment["poc_recommendations"] = [item.copy() for item in items]
    prioritization = front.get("use_case_prioritization") if isinstance(front, dict) else None
    if isinstance(prioritization, dict):
        prioritization["selection_logic"] = [
            {
                "use_case_id": item["use_case_id"],
                "theme": item["theme"],
                "why_now": item["reason"],
                "proof_needed": item["first_step"],
                "depends_on": item["depends_on"],
                "basis": item["basis"],
            }
            for item in items
        ]
    assessment["consulting_front_matter"] = front
    # 既存の詳細がIDまたはテーマで対応しない場合は、poc_logic_details_forが
    # 安全な共通フォールバックを生成する。別テーマの詳細を表示し続けない。
    assessment["poc_logic_details"] = poc_logic_details_for(assessment)
    # 構成図・デモ用の状態付きデータも、この時点で確定したP1へ追随させる。
    # ただし、承認済みJSONに有効な提案詳細がある場合は上書きしない。描画だけの
    # 再 materialize により、人手レビュー済みの設計内容を失わないためである。
    assessment["technical_proposal"] = technical_proposal_for(assessment)
    return portfolio

def industry_value_story_for(assessment: dict, front_matter: dict) -> dict:
    """業界・サービス固有の論点を、AI支援から事業価値まで一枚で追うデータへ整える。

    新しい会社固有の事実を生成せず、既に根拠区分付きで作成した前半データから
    派生させる。従って入力業界・サービス・公開調査結果に応じて内容が変わり、
    特定顧客向けの固定文言にはならない。
    """
    service = front_matter["service_model"]
    diagnosis = front_matter["operating_diagnosis"]
    context = front_matter["decision_context"]
    business_value = assessment.get("business_value") if isinstance(assessment.get("business_value"), dict) else {}
    impact_areas = [item for item in business_value.get("impact_areas", []) if isinstance(item, dict)]
    points = [_front_text(item, 84) for item in assessment.get("assessment_points", []) if _front_text(item, 84)]

    drivers: list[dict[str, str]] = []
    for index, item in enumerate(diagnosis.get("issue_tree", [])):
        if not isinstance(item, dict) or len(drivers) >= 3:
            continue
        label = _front_text(item.get("issue"), 38)
        detail = _front_text(item.get("business_effect") or item.get("validation_question"), 82)
        if label and detail:
            drivers.append({
                "label": label,
                "detail": detail,
                "basis": _front_basis(item.get("basis")) or "分析仮説",
            })
    while len(drivers) < 3:
        index = len(drivers)
        drivers.append({
            "label": f"確認すべき論点 {index + 1}",
            "detail": points[index] if index < len(points) else "対象業務・利用者・データを確認し、事業化の判断条件を定義する。",
            "basis": "要確認",
        })

    value_chain = [item for item in service.get("value_chain", []) if isinstance(item, dict)]
    chain_labels = "・".join(_front_text(item.get("stage"), 18) for item in value_chain[:2] if _front_text(item.get("stage"), 18))
    data_labels = "・".join(_front_text(item.get("data"), 24) for item in value_chain[:2] if _front_text(item.get("data"), 24))
    flow = [
        {"number": "1", "title": "業務・サービスの変化", "detail": chain_labels or "利用者・業務の判断点を捉える"},
        {"number": "2", "title": "判断に使うデータ", "detail": data_labels or "業務実績・マスタ・文書・ログを整理する"},
        {"number": "3", "title": "AIによる支援", "detail": "根拠と候補を提示し、人が採用・修正・例外判断を行う"},
        {"number": "4", "title": "顧客・事業価値", "detail": "品質、生産性、継続利用・横展開の成果を測定する"},
    ]

    outcomes: list[dict[str, str]] = []
    for item in impact_areas[:3]:
        label = _front_text(item.get("area"), 32)
        detail = _front_text(item.get("expected_impact") or item.get("ai_enabled"), 86)
        if label and detail:
            outcomes.append({"label": label, "detail": detail})
    while len(outcomes) < 3:
        outcomes.append({
            "label": ("利用者価値", "業務価値", "事業価値")[len(outcomes)],
            "detail": "対象KPI・比較条件・責任者をPoC開始前に合意し、実測で判断する。",
        })
    return {
        "headline": _front_text(service.get("takeaway") or assessment.get("executive_summary"), 130),
        "drivers": drivers,
        "flow": flow,
        "outcomes": outcomes,
        "decision_question": _front_text(context.get("decision_question"), 105),
        "in_scope": "・".join(_front_text(item, 36) for item in context.get("in_scope", [])[:3]),
        "success_definition": _front_text(context.get("success_definition"), 120),
    }


def _business_value_model_role(assessment: dict) -> str:
    """組織の立場を返す。未指定は ``unknown`` として安全側へ倒す。

    ``provider`` / ``mixed`` を入力または承認済みJSONで明示した場合だけ、
    MRR・ARPA等を含む提供者側の事業化レーンを適用する。案件種別を推測して
    提供者と扱うと、事業会社・運用会社にも不適切な商品化KPIを表示し得るため。
    """
    raw = _front_text(
        assessment.get("business_model_role") or assessment.get("organization_role"), 24,
    ).lower()
    aliases = {
        "provider": "provider", "isv": "provider", "saas": "provider", "software_provider": "provider",
        "operator": "operator", "enterprise": "operator", "customer": "operator",
        "mixed": "mixed", "unknown": "unknown",
        "提供者": "provider", "サービス提供者": "provider", "事業者": "provider",
        "導入企業": "operator", "事業会社": "operator", "運用者": "operator",
        "両方": "mixed", "不明": "unknown",
    }
    return aliases.get(raw, "unknown")


def _business_value_model_basis(items: list[dict]) -> str:
    """既存の根拠区分を優先し、なければ仮説として明示する。"""
    for item in items:
        basis = _front_basis(item.get("basis"))
        if basis:
            return basis
    return "分析仮説"


def _business_value_model_area_for_dimension(areas: list[dict], dimension: str, offset: int) -> dict:
    """既存の価値領域から、四つの運用価値を無理なく説明するための素材を選ぶ。"""
    keywords = {
        "品質": ("品質", "精度", "顧客", "体験", "正確", "対応"),
        "時間": ("時間", "工数", "生産性", "効率", "速度", "対応"),
        "リスク": ("リスク", "統制", "例外", "安全", "損失", "監査"),
        "定着": ("定着", "利用", "継続", "採用", "顧客", "体験"),
    }
    terms = keywords.get(dimension, ())
    for item in areas:
        text = " ".join(str(item.get(key, "")) for key in ("area", "business_rationale", "ai_enabled", "expected_impact"))
        if any(term in text for term in terms):
            return item
    return areas[offset % len(areas)] if areas else {}


def normalize_business_value_model(raw: object) -> dict:
    """二層の価値モデルJSONを正規化する。

    ``provider`` は標準機能・有償オプション・個別適用という商品化仮説、
    ``customer_operation`` は品質・時間・リスク・定着という導入先の実測領域を
    分けて保持する。すべての項目に根拠区分を付け、数値目標はこの契約では持たない。
    """
    if not isinstance(raw, dict) or _front_text(raw.get("schema_version"), 8) != BUSINESS_VALUE_MODEL_SCHEMA_VERSION:
        return {}
    role = _front_text(raw.get("organization_role"), 24).lower()
    if role not in BUSINESS_VALUE_MODEL_ROLES:
        return {}
    provider_raw = raw.get("provider")
    operation_raw = raw.get("customer_operation")
    bridge_raw = raw.get("bridge")
    if not all(isinstance(value, dict) for value in (provider_raw, operation_raw, bridge_raw)):
        return {}

    provider = {
        "applicable": provider_raw.get("applicable"),
        "label": _front_text(provider_raw.get("label"), 56),
        "positioning": _front_text(provider_raw.get("positioning"), 150),
        "segments": _front_items(
            provider_raw.get("segments"), count=3,
            fields={"segment": 46, "job_to_be_done": 94, "basis": 12}, basis_fields=("basis",),
        ),
        "packaging": _front_items(
            provider_raw.get("packaging"), count=3,
            fields={"offer": 30, "theme": 72, "commercial_hypothesis": 100, "proof_needed": 100, "basis": 12},
            basis_fields=("basis",),
        ),
        "metrics": _front_items(
            provider_raw.get("metrics"), count=4,
            fields={"metric": 34, "decision_use": 90, "baseline_status": 12, "basis": 12},
            basis_fields=("basis",),
        ),
    }
    if not isinstance(provider["applicable"], bool) or not provider["label"] or not provider["positioning"]:
        return {}
    if not all((provider["segments"], provider["packaging"], provider["metrics"])):
        return {}
    if any(item["baseline_status"] not in BUSINESS_VALUE_MODEL_BASELINE_VALUES for item in provider["metrics"]):
        return {}

    outcomes = _front_items(
        operation_raw.get("outcomes"), count=4,
        fields={"dimension": 16, "value": 104, "measure": 84, "basis": 12}, basis_fields=("basis",),
    )
    operation = {
        "label": _front_text(operation_raw.get("label"), 56),
        "positioning": _front_text(operation_raw.get("positioning"), 150),
        "outcomes": outcomes,
    }
    if (not operation["label"] or not operation["positioning"]
            or [item["dimension"] for item in outcomes] != list(BUSINESS_VALUE_MODEL_DIMENSIONS)):
        return {}

    bridge = {
        "customer_to_adoption": _front_text(bridge_raw.get("customer_to_adoption"), 108),
        "adoption_to_product": _front_text(bridge_raw.get("adoption_to_product"), 108),
        "product_to_business": _front_text(bridge_raw.get("product_to_business"), 108),
        "decision": _front_text(bridge_raw.get("decision"), 126),
        "basis": _front_basis(bridge_raw.get("basis")),
    }
    if not all(bridge.values()):
        return {}
    return {
        "schema_version": BUSINESS_VALUE_MODEL_SCHEMA_VERSION,
        "organization_role": role,
        "provider": provider,
        "customer_operation": operation,
        "bridge": bridge,
    }


def fallback_business_value_model(assessment: dict, front_matter: dict | None = None) -> dict:
    """既存のアセスメントから、数値を新設せず二層の価値モデルを導く。

    商品化方式や商用KPIは「現状の事実」ではなく、PoCで判断する仮説として表示する。
    導入企業のみを評価する案件では ``business_model_role=operator`` を指定でき、
    提供者側の表現を「該当する場合」に落とせる。
    """
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    portfolio = priority_pocs_for(assessment, front)
    role = _business_value_model_role(assessment)
    business_value = assessment.get("business_value") if isinstance(assessment.get("business_value"), dict) else {}
    areas = [item for item in business_value.get("impact_areas", []) if isinstance(item, dict)]
    issue_tree = front.get("operating_diagnosis", {}).get("issue_tree", []) if isinstance(front, dict) else []
    issue_tree = [item for item in issue_tree if isinstance(item, dict)]
    basis = _business_value_model_basis(issue_tree)

    while len(portfolio) < 3:
        number = len(portfolio) + 1
        portfolio.append({
            "priority": f"P{number}", "theme": f"優先テーマ {number}",
            "first_step": "対象業務、データ、比較条件、責任者をPoC開始前に合意する。", "basis": basis,
        })
    packaging_labels = ("標準機能候補", "有償オプション候補", "個別適用候補")
    packaging: list[dict[str, str]] = []
    segments: list[dict[str, str]] = []
    for index, item in enumerate(portfolio[:3]):
        theme = _front_text(item.get("theme") or f"優先テーマ {index + 1}", 72)
        item_basis = _front_basis(item.get("basis")) or basis
        packaging.append({
            "offer": packaging_labels[index],
            "theme": theme,
            "commercial_hypothesis": "導入先で再現性と運用条件を確認し、共通設定で提供できる範囲を判断する。",
            "proof_needed": _front_text(item.get("first_step") or "代表データ・利用者・比較条件を合意してPoCで確認する。", 100),
            "basis": item_basis,
        })
        segments.append({
            "segment": f"{theme}を必要とする利用組織",
            "job_to_be_done": "対象業務の判断・確認を、根拠付きで速く一貫して行える状態を目指す。",
            "basis": item_basis,
        })

    metrics = [
        {"metric": "MRR・ARPA", "decision_use": "AI機能の提供形態ごとの単価・契約価値への寄与を確認する。", "baseline_status": "要確認", "basis": "分析仮説"},
        {"metric": "継続利用・定着", "decision_use": "利用率、提案採用率、継続利用を通じて導入先の受容性を判断する。", "baseline_status": "要確認", "basis": "分析仮説"},
        {"metric": "サポート原価", "decision_use": "問い合わせ・例外対応・運用工数への影響を確認し、提供原価を見積もる。", "baseline_status": "要確認", "basis": "分析仮説"},
        {"metric": "横展開・スケール", "decision_use": "設定・導入・権限設計の再利用性から展開速度と対象範囲を判断する。", "baseline_status": "要確認", "basis": "分析仮説"},
    ]
    measure_by_dimension = {
        "品質": "正確性・修正率・根拠確認率を、現行運用と同じ対象範囲で比較する。",
        "時間": "処理時間・確認工数・待機時間を、現行運用と同じ対象範囲で比較する。",
        "リスク": "例外検知・見逃し・差戻しを、対象業務の定義に沿って確認する。",
        "定着": "利用率・採用率・手動復帰を、対象利用者と期間を定めて確認する。",
    }
    outcomes: list[dict[str, str]] = []
    for index, dimension in enumerate(BUSINESS_VALUE_MODEL_DIMENSIONS):
        area = _business_value_model_area_for_dimension(areas, dimension, index)
        value = _front_text(area.get("ai_enabled") or area.get("expected_impact") or area.get("business_rationale"), 104)
        if not value:
            value = "対象業務・利用者・データを特定し、AI支援が業務判断に与える変化をPoCで確認する。"
        outcomes.append({
            "dimension": dimension,
            "value": value,
            "measure": measure_by_dimension[dimension],
            "basis": basis,
        })
    provider_applicable = role in {"provider", "mixed"}
    provider_label = (
        "提供者側：AIサービスの事業化" if provider_applicable
        else "提供者側（提供形態を確認後に評価）：AIサービスの事業仮説"
    )
    model = {
        "schema_version": BUSINESS_VALUE_MODEL_SCHEMA_VERSION,
        "organization_role": role,
        "provider": {
            "applicable": provider_applicable,
            "label": provider_label,
            "positioning": "導入先の実測結果を、標準機能・有償オプション・個別適用の判断と、利用・継続・運用原価・横展開の事業KPIへ接続する。",
            "segments": segments,
            "packaging": packaging,
            "metrics": metrics,
        },
        "customer_operation": {
            "label": "導入先側：業務・利用価値",
            "positioning": "品質・時間・リスク・定着を、対象業務・利用者・比較条件を定めたPoCで実測し、AI支援を業務に組み込めるか判断する。",
            "outcomes": outcomes,
        },
        "bridge": {
            "customer_to_adoption": "導入先の業務KPIを、現行運用と同じ対象範囲・比較条件で実測する。",
            "adoption_to_product": "利用・採用・手動復帰の証跡から、共通設定で提供できる機能範囲を切り出す。",
            "product_to_business": "商品化後は利用・継続・原価・横展開を事業KPIとして継続測定する。",
            "decision": "PoC終了時に、標準機能化・有償オプション・個別適用・見送りのいずれかを決定する。",
            "basis": "分析仮説",
        },
    }
    normalized = normalize_business_value_model(model)
    return normalized or {}


def business_value_model_for(assessment: dict, front_matter: dict | None = None) -> dict:
    """明示された二層モデルを優先し、旧JSONでは安全な仮説モデルを返す。"""
    normalized = normalize_business_value_model(assessment.get("business_value_model"))
    return normalized or fallback_business_value_model(assessment, front_matter)


def materialize_business_value_model(assessment: dict, front_matter: dict | None = None) -> dict:
    """JSONレビューと描画を同じ二層価値モデルへ固定する。"""
    model = business_value_model_for(assessment, front_matter)
    if model:
        assessment["business_value_model"] = model
    return model

def assessment_story_for(assessment: dict) -> dict:
    """既存呼び出しとの互換性を保ちつつ、詳細な前半の構造化データを提供する。"""
    business_value = assessment.get("business_value") if isinstance(assessment.get("business_value"), dict) else {}
    areas = [item for item in business_value.get("impact_areas", []) if isinstance(item, dict)][:3]
    points = [str(item) for item in assessment.get("assessment_points", []) if str(item).strip()][:3]
    use_cases = [item for item in assessment.get("use_cases", []) if isinstance(item, dict)][:15]
    front_matter = consulting_front_matter_for(assessment)
    pocs = priority_pocs_for(assessment, front_matter)
    return {
        "summary": str(assessment.get("executive_summary") or "対象サービスの業務・データ・利用者を基に、AI実装の優先テーマとPoCの判断条件を整理します。"),
        "areas": areas, "points": points, "pocs": pocs, "use_cases": use_cases,
        "service_name": str(assessment.get("service_name") or "対象サービス"),
        "company_name": str(assessment.get("company_name") or "ご提案先企業様"),
        "front_matter": front_matter,
        "industry_value_story": industry_value_story_for(assessment, front_matter),
        "business_value_model": business_value_model_for(assessment, front_matter),
        "poc_selection_scorecard": poc_selection_scorecard_for(assessment, front_matter),
        "poc_priority_decision": (
            normalize_poc_priority_decision(assessment.get("poc_priority_decision"), assessment, front_matter)
            or priority_decision_for(assessment, front_matter)
        ),
        "poc_start_readiness": (
            normalize_poc_start_readiness(assessment.get("poc_start_readiness"), assessment, front_matter)
            or poc_start_readiness_for(assessment, front_matter)
        ),
    }
