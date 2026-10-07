"""固定済みAIユースケース・カタログの読込と適用契約。

過去の凍結レビューJSONから再利用するのは、サービス別の15件カタログだけである。
過去のPoC順位、調査、実装詳細は入力に含まれていても採用せず、今回の入力・調査で
後段を作り直す。ただし軽量カタログJSONに明示した ``poc_selection`` は、合意済みの
P1〜P3の *ID・順序だけ* を固定するための新規入力契約として受け付ける。
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Any

from ai_assess_runtime.poc_contract import normalized_use_case_label
from ai_assess_runtime.review_snapshot import canonical_sha256, file_sha256


USE_CASE_CATALOG_FORMAT = "ai-assess/use-case-catalog-v1"
USE_CASE_CATALOG_REUSE_FORMAT = "ai-assess/use-case-catalog-reuse/v1"
USE_CASE_CATALOG_SCHEMA_VERSION = "1"
FIXED_POC_SELECTION_FORMAT = "ai-assess/fixed-poc-selection-v1"
FIXED_POC_SELECTION_SCHEMA_VERSION = "1"

_REQUIRED_CASE_FIELDS = (
    "coverage_area",
    "use_case",
    "ai_technology",
    "description",
)
_DERIVED_POC_FIELDS = (
    "poc_recommendations",
    "poc_logic_details",
    "consulting_front_matter",
    "poc_portfolio",
    "poc_charters",
    "poc_selection_scorecard",
    "poc_priority_decision",
    "poc_start_readiness",
    "multitenant_governance",
    "technical_proposal",
    "poc_measurement_design",
    "business_value_model",
    "ai_product_business_impact",
    "customer_priority_binding",
    "fixed_poc_selection_binding",
)


def _required_text(value: object, label: str) -> str:
    text = str(value or "").strip()
    if not text:
        raise ValueError(f"ユースケースカタログの{label}が空です。")
    return text


def _case_number(value: object, label: str) -> int:
    try:
        number = int(str(value).strip())
    except (TypeError, ValueError) as error:
        raise ValueError(f"ユースケースカタログの{label}は1〜15の整数で指定してください。") from error
    if not 1 <= number <= 15:
        raise ValueError(f"ユースケースカタログの{label}は1〜15の範囲で指定してください。")
    return number


def _use_case_id(group_index: int, number: int) -> str:
    return f"UC{number:02d}" if group_index == 1 else f"S{group_index:02d}-UC{number:02d}"


def _normalize_cases(raw_cases: object, group_index: int) -> list[dict[str, object]]:
    if not isinstance(raw_cases, list) or len(raw_cases) != 15:
        raise ValueError("各サービスのユースケースカタログは15件必要です。")
    by_number: dict[int, dict[str, object]] = {}
    coverage_keys: set[str] = set()
    title_keys: set[str] = set()
    for raw_case in raw_cases:
        if not isinstance(raw_case, dict):
            raise ValueError("ユースケースカタログの各候補はオブジェクトで指定してください。")
        number = _case_number(raw_case.get("no"), "no")
        if number in by_number:
            raise ValueError("ユースケースカタログのnoが重複しています。")
        case = {"no": number, "use_case_id": _use_case_id(group_index, number)}
        for field in _REQUIRED_CASE_FIELDS:
            case[field] = _required_text(raw_case.get(field), field)
        provided_id = str(raw_case.get("use_case_id") or "").strip()
        if provided_id and provided_id != case["use_case_id"]:
            raise ValueError(
                "ユースケースカタログのuse_case_idがnoと一致しません: "
                f"no={number}、期待値={case['use_case_id']}"
            )
        coverage_key = normalized_use_case_label(case["coverage_area"])
        title_key = normalized_use_case_label(case["use_case"])
        if not coverage_key:
            raise ValueError("ユースケースカタログのcoverage_areaが空です。")
        if coverage_key in coverage_keys:
            raise ValueError("ユースケースカタログのcoverage_areaが重複しています。")
        if title_key in title_keys:
            raise ValueError("ユースケースカタログのuse_caseが重複しています。")
        coverage_keys.add(coverage_key)
        title_keys.add(title_key)
        by_number[number] = case
    if set(by_number) != set(range(1, 16)):
        raise ValueError("ユースケースカタログのnoは1〜15の連番である必要があります。")
    return [by_number[number] for number in range(1, 16)]


def _normalize_groups(raw_groups: object) -> list[dict[str, object]]:
    if not isinstance(raw_groups, list) or not 1 <= len(raw_groups) <= 3:
        raise ValueError("ユースケースカタログには1〜3件のサービス別一覧が必要です。")
    groups: list[dict[str, object]] = []
    service_keys: set[str] = set()
    for group_index, raw_group in enumerate(raw_groups, 1):
        if not isinstance(raw_group, dict):
            raise ValueError("サービス別ユースケースカタログはオブジェクトで指定してください。")
        service_name = _required_text(raw_group.get("service_name"), "service_name")
        service_type = _required_text(raw_group.get("service_type"), "service_type")
        service_key = normalized_use_case_label(service_name)
        if service_key in service_keys:
            raise ValueError("ユースケースカタログのservice_nameが重複しています。")
        service_keys.add(service_key)
        groups.append({
            "service_name": service_name,
            "service_type": service_type,
            "use_cases": _normalize_cases(raw_group.get("use_cases"), group_index),
        })
    return groups


def _normalize_poc_selection(
    raw_selection: object,
    groups: list[dict[str, object]],
) -> dict[str, object] | None:
    """軽量カタログの明示P1〜P3を、不変IDへ厳密に正規化する。

    これは過去の ``poc_recommendations`` を転記する仕組みではない。再利用を許可
    するのは、一覧のアイコンと後続詳細ページを同じ候補へ結び付けるための
    ``priority / use_case_id / use_case_no / theme / detail_slide_order`` だけである。
    優先理由、実装詳細、定量評価は今回の入力と調査で生成し直す。
    """
    if raw_selection is None:
        return None
    if not isinstance(raw_selection, dict):
        raise ValueError("poc_selectionはオブジェクトで指定してください。")
    if str(raw_selection.get("format") or "").strip() != FIXED_POC_SELECTION_FORMAT:
        raise ValueError(
            "poc_selection.formatは"
            f"{FIXED_POC_SELECTION_FORMAT}で指定してください。"
        )
    if str(raw_selection.get("schema_version") or "").strip() != FIXED_POC_SELECTION_SCHEMA_VERSION:
        raise ValueError(
            "poc_selection.schema_versionは"
            f"{FIXED_POC_SELECTION_SCHEMA_VERSION}で指定してください。"
        )
    raw_items = raw_selection.get("items")
    if not isinstance(raw_items, list) or len(raw_items) != 3:
        raise ValueError("poc_selection.itemsはP1〜P3の3件で指定してください。")
    primary_cases = groups[0].get("use_cases") if groups and isinstance(groups[0], dict) else None
    if not isinstance(primary_cases, list):
        raise ValueError("poc_selectionに対応する代表サービスの15件カタログがありません。")
    case_by_id = {
        str(case.get("use_case_id") or ""): case
        for case in primary_cases if isinstance(case, dict)
    }
    items: list[dict[str, object]] = []
    seen_ids: set[str] = set()
    for index, raw_item in enumerate(raw_items, 1):
        if not isinstance(raw_item, dict):
            raise ValueError("poc_selection.itemsの各項目はオブジェクトで指定してください。")
        priority = str(raw_item.get("priority") or "").strip()
        if priority != f"P{index}":
            raise ValueError("poc_selection.itemsはP1、P2、P3の順に指定してください。")
        use_case_id = str(raw_item.get("use_case_id") or "").strip()
        case = case_by_id.get(use_case_id)
        if case is None or use_case_id in seen_ids:
            raise ValueError("poc_selection.use_case_idは代表サービスの重複しない候補IDで指定してください。")
        use_case_no = _case_number(raw_item.get("use_case_no"), "poc_selection.use_case_no")
        expected_no = int(case["no"])
        if use_case_no != expected_no:
            raise ValueError(
                "poc_selection.use_case_noがuse_case_idと一致しません: "
                f"{use_case_id} / 期待値={expected_no}"
            )
        theme = _required_text(raw_item.get("theme"), "poc_selection.theme")
        if normalized_use_case_label(theme) != normalized_use_case_label(case["use_case"]):
            raise ValueError(
                "poc_selection.themeが15件カタログのユースケース名と一致しません: "
                f"{use_case_id}"
            )
        detail_slide_order = raw_item.get("detail_slide_order")
        if isinstance(detail_slide_order, bool) or detail_slide_order != index:
            raise ValueError("poc_selection.detail_slide_orderはP1〜P3に対応する1、2、3で指定してください。")
        items.append({
            "priority": priority,
            "use_case_id": use_case_id,
            "use_case_no": str(expected_no),
            "theme": str(case["use_case"]),
            "detail_slide_order": index,
        })
        seen_ids.add(use_case_id)
    return {
        "format": FIXED_POC_SELECTION_FORMAT,
        "schema_version": FIXED_POC_SELECTION_SCHEMA_VERSION,
        "selection_reason": str(raw_selection.get("selection_reason") or "").strip(),
        "items": items,
    }


def _groups_from_payload(payload: dict[str, object]) -> tuple[object, str, str]:
    """受け取った2形式から、サービス群・元形式・会社名を取り出す。"""
    source_format = str(payload.get("format") or "").strip()
    if source_format == USE_CASE_CATALOG_FORMAT:
        raw_groups = payload.get("service_use_case_groups")
        if raw_groups is None:
            raw_groups = [{
                "service_name": payload.get("service_name"),
                "service_type": payload.get("service_type"),
                "use_cases": payload.get("use_cases"),
            }]
        return raw_groups, source_format, str(payload.get("company_name") or "").strip()

    assessment = payload.get("assessment")
    if not isinstance(assessment, dict):
        raise ValueError(
            "ユースケースカタログJSONは、過去のassessment JSONまたは"
            f"format={USE_CASE_CATALOG_FORMAT}である必要があります。"
        )
    raw_groups = assessment.get("service_use_case_groups")
    if raw_groups is None:
        raw_groups = [{
            "service_name": assessment.get("service_name"),
            "service_type": assessment.get("service_genre"),
            "use_cases": assessment.get("use_cases"),
        }]
    return raw_groups, source_format, str(assessment.get("company_name") or "").strip()


def _primary_cases_match(groups: list[dict[str, object]], raw_cases: object) -> bool:
    """凍結レビューJSON内の二重表現が同じ15件を示すか確認する。"""
    if raw_cases is None:
        return True
    try:
        normalized = _normalize_cases(raw_cases, 1)
    except ValueError:
        return False
    return normalized == groups[0]["use_cases"]


def load_use_case_catalog(path: Path) -> dict[str, object]:
    """過去の凍結JSONまたは軽量カタログJSONを安全に正規化して読み込む。"""
    source_file = Path(path)
    if not source_file.is_file():
        raise ValueError(f"ユースケースカタログJSONが見つかりません: {source_file}")
    try:
        payload = json.loads(source_file.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValueError(f"ユースケースカタログJSONを読み込めません: {source_file}") from error
    if not isinstance(payload, dict):
        raise ValueError("ユースケースカタログJSONの最上位はオブジェクトである必要があります。")
    raw_groups, source_format, company_name = _groups_from_payload(payload)
    if not source_format:
        raise ValueError("ユースケースカタログJSONのformatがありません。")
    groups = _normalize_groups(raw_groups)
    source_assessment = payload.get("assessment")
    if isinstance(source_assessment, dict) and not _primary_cases_match(
            groups, source_assessment.get("use_cases")):
        raise ValueError(
            "過去assessment JSONのassessment.use_casesと"
            "service_use_case_groups[0].use_casesが一致しません。"
        )
    # 凍結assessment JSON内の過去P1〜P3を暗黙に復元しない。明示的な固定選定は
    # 軽量カタログ形式でだけ受け付け、利用者が意図して指定した場合に限る。
    poc_selection = _normalize_poc_selection(
        payload.get("poc_selection") if source_format == USE_CASE_CATALOG_FORMAT else None,
        groups,
    )
    catalog_sha256 = canonical_sha256(groups)
    return {
        "format": USE_CASE_CATALOG_FORMAT,
        "schema_version": USE_CASE_CATALOG_SCHEMA_VERSION,
        "company_name": company_name,
        "service_use_case_groups": groups,
        "use_cases": copy.deepcopy(groups[0]["use_cases"]),
        "poc_selection": copy.deepcopy(poc_selection),
        "provenance": {
            "format": USE_CASE_CATALOG_REUSE_FORMAT,
            "schema_version": USE_CASE_CATALOG_SCHEMA_VERSION,
            "source_file": str(source_file),
            "source_sha256": file_sha256(source_file),
            "source_format": source_format,
            "catalog_sha256": catalog_sha256,
            "service_count": len(groups),
            "service_names": [str(group["service_name"]) for group in groups],
            "poc_selection": copy.deepcopy(poc_selection),
            "poc_selection_sha256": canonical_sha256(poc_selection) if poc_selection else "",
            "poc_selection_count": len(poc_selection["items"]) if poc_selection else 0,
        },
    }


def _service_identity_key(value: object) -> str:
    # 正式名称と略称の照合では、登録商標・商標表示の有無だけを製品差と
    # みなさない。一般語を削るのではなく、表示上の商標記号だけを除外する。
    return normalized_use_case_label(value).translate(
        str.maketrans("", "", "®™℠©"),
    ).removesuffix("等")


def _service_names_match(current_name: object, source_name: object) -> bool:
    current_key = _service_identity_key(current_name)
    source_key = _service_identity_key(source_name)
    if not current_key or not source_key:
        return False
    if current_key == source_key:
        return True
    # 表紙では略称・読み仮名の片方だけになることがあるため、十分な長さの
    # 部分一致だけを許可する。短い一般名詞だけでの誤結合は防ぐ。
    return min(len(current_key), len(source_key)) >= 4 and (
        current_key in source_key or source_key in current_key
    )


def apply_use_case_catalog(assessment: dict[str, object], catalog: dict[str, object]) -> dict[str, object]:
    """今回の分析結果へ固定15件を適用し、過去PoC由来の値をすべて破棄する。"""
    groups = catalog.get("service_use_case_groups")
    if not isinstance(groups, list) or not groups or not isinstance(groups[0], dict):
        raise ValueError("適用するユースケースカタログが不正です。")
    source_service = str(groups[0].get("service_name") or "").strip()
    current_service = str(assessment.get("service_name") or "").strip()
    if not _service_names_match(current_service, source_service):
        raise ValueError(
            "入力の対象サービスとユースケースカタログの代表サービスが一致しません: "
            f"入力={current_service or '（未設定）'} / カタログ={source_service or '（未設定）'}"
        )
    result = copy.deepcopy(assessment)
    # 同一製品の短縮名と正式名が一致した場合、固定カタログの代表サービス名を
    # 以降の単一表示名として採用する。
    result["service_name"] = source_service
    fixed_groups = copy.deepcopy(groups)
    result["service_use_case_groups"] = fixed_groups
    result["use_cases"] = copy.deepcopy(fixed_groups[0]["use_cases"])
    for field in _DERIVED_POC_FIELDS:
        result.pop(field, None)
    # 後段のfront matterで、今回の入力・調査に基づくP1〜P3を再確定する。
    result["poc_recommendations"] = []
    return result
