"""定量効果、責任境界、PoC測定、原典照合を固定するドメイン契約。"""
import copy
from dataclasses import dataclass
from enum import Enum
import json
import re
from ai_assess_runtime.http_safety import public_https_url
from ai_assess_runtime.target_metric import (
    _monetary_target_parts,
    _normalize_monetary_target,
    _normalize_percentage_target,
    _percentage_range_parts,
    _verified_management_amount_target,
)
from ai_assess_runtime.poc_contract import (
    _business_value_model_role,
    _front_text,
    canonical_use_case_id,
    consulting_front_matter_for,
    materialize_business_value_model,
    materialize_poc_decision_data,
    materialize_poc_portfolio,
    materialize_poc_priority_decision,
    materialize_poc_selection_scorecard,
    materialize_poc_start_readiness,
    normalized_use_case_label,
    poc_logic_details_for,
    priority_pocs_for,
    technical_proposal_for,
)
POC_MEASUREMENT_DESIGN_SCHEMA_VERSION = "5"
AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION = "1"
AI_PRODUCT_BUSINESS_IMPACT_PROVIDER_CATEGORIES = (
    "revenue_growth", "profitability", "recurring_revenue",
)
AI_PRODUCT_BUSINESS_IMPACT_OPERATOR_CATEGORIES = (
    "revenue_growth", "profitability", "scalability_resilience",
)
LEGACY_POC_MEASUREMENT_DESIGN_SCHEMA_VERSIONS = {"2", "3", "4"}
BUSINESS_EFFECT_DISPLAY_LAYERS = {"customer_outcome_kpi", "provider_business_kpi"}
ALL_MEASUREMENT_DISPLAY_LAYERS = {"product_kpi", *BUSINESS_EFFECT_DISPLAY_LAYERS}
def is_safe_public_https_url(value: object) -> bool:
    """Webリサーチの取得対象を公開HTTPS URLへ限定する。"""
    # 候補一覧のフィルタではDNSアクセスを発生させず、実取得直前に
    # safe_fetch_public_httpsが解決先IPまで再検証する。
    return public_https_url(value)
def quantitative_priority_pocs(assessment: dict) -> list[dict[str, str]]:
    """Return the exact P1-P3 identity used by quantitative generation."""
    raw = priority_pocs_for(assessment)
    if len(raw) != 3:
        return []
    rows: list[dict[str, str]] = []
    for index, item in enumerate(raw, 1):
        if not isinstance(item, dict):
            return []
        priority = _front_text(item.get("priority") or f"P{index}", 8)
        use_case_id = _front_text(
            item.get("use_case_id") or canonical_use_case_id(item.get("use_case_no"))
            or f"UC{index:02d}", 12,
        )
        theme = _front_text(item.get("theme"), 72)
        if priority != f"P{index}" or not use_case_id or not theme:
            return []
        rows.append({"priority": priority, "use_case_id": use_case_id, "theme": theme})
    if len({row["use_case_id"] for row in rows}) != 3:
        return []
    return rows
def _quantitative_owner_issues(value: object, path: str) -> list[str]:
    if not isinstance(value, dict):
        return [f"{path}は責任境界オブジェクトである必要があります。"]
    if str(value.get("scope") or "") not in {"customer", "provider", "shared"}:
        return [f"{path}.scopeはcustomer/provider/sharedのいずれかです。"]
    if not str(value.get("role") or "").strip():
        return [f"{path}.roleがありません。"]
    if str(value.get("status") or "") not in {"confirmed", "confirm"}:
        return [f"{path}.statusはconfirmed/confirmのいずれかです。"]
    return []
def _quantitative_layer_issues(value: object, path: str, *, require_numeric: bool) -> list[str]:
    if not isinstance(value, dict):
        return [f"{path}はKPIオブジェクトである必要があります。"]
    required = ("kpi", "target", "baseline_definition", "comparison_condition", "formula",
                "metric_owner", "attribution_level", "causal_link")
    missing = [field for field in required if not str(value.get(field) or "").strip()]
    issues = [f"{path}の必須項目が不足しています: {', '.join(missing)}"] if missing else []
    issues.extend(_quantitative_owner_issues(value.get("metric_owner"), f"{path}.metric_owner"))
    if str(value.get("attribution_level") or "") not in {"direct", "contributory", "enabling"}:
        issues.append(f"{path}.attribution_levelはdirect/contributory/enablingのいずれかです。")
    if require_numeric and not _is_business_effect_metric(value.get("target")):
        issues.append(f"{path}.targetにはPoCで測定可能な数値目標が必要です。")
    return issues
def _is_variable_measurement_formula(value: object) -> bool:
    """Return True for a formula that depends on measured inputs, not a fixed claim."""
    text = re.sub(r"\s+", "", str(value or ""))
    if not text:
        return False
    # ``(PoC実績値-基準値)/基準値×100%`` の 100% は、効果目標ではなく
    # 比率を百分率へ換算する定数である。これだけは除外して判定する一方、
    # ``実績値×20%`` や ``実績値×500万円`` のような未検証の効果値は拒否する。
    claim_text = re.sub(r"(?:×|x|X|\*)100(?:%|％)", "", text)
    if _quantitative_tokens(claim_text):
        return False
    variable_markers = (
        "実績値", "基準値", "現状値", "測定値", "入力値", "対象値", "対象件数",
        "対象時間", "利用件数", "利用者数", "単価", "分母", "分子", "改善率",
        "採用率", "利用率", "PoCで確定", "PoCで算定", "PoCで測定",
    )
    return any(marker in text for marker in variable_markers)
def _ai_product_business_categories(assessment: dict) -> tuple[str, str, str]:
    """Return company-level KPI categories without inferring customer operations.
    ISV/SaaS providers and mixed businesses are assessed on product revenue,
    profitability and recurring revenue.  An operator's third lane is business
    capacity/resilience instead of SaaS retention.  In either case the owner is
    always the assessed company; the assessed company's customers only appear
    as causal evidence, never as the headline metric owner.
    """
    role = _business_value_model_role(assessment)
    if role in {"provider", "mixed"}:
        return AI_PRODUCT_BUSINESS_IMPACT_PROVIDER_CATEGORIES
    return AI_PRODUCT_BUSINESS_IMPACT_OPERATOR_CATEGORIES
AI_PRODUCT_BUSINESS_KPI_KEYWORDS: dict[str, tuple[str, ...]] = {
    "revenue_growth": (
        "売上", "ARR", "MRR", "ARPA", "受注", "有償", "契約単価", "契約売上",
    ),
    "profitability": (
        "利益", "粗利", "原価", "提供原価", "限界利益", "営業利益",
    ),
    "recurring_revenue": (
        "継続収益", "NRR", "ネット収益継続", "純収益継続", "Net Revenue Retention",
        "更新率", "解約率", "アップセル", "ARR", "継続売上",
    ),
    "scalability_resilience": (
        "処理能力", "生産能力", "事業継続", "売上機会", "損失", "供給能力",
    ),
}
AI_PRODUCT_BUSINESS_LEADING_KPI_KEYWORDS: dict[str, tuple[str, ...]] = {
    "revenue_growth": (
        "有償", "付帯率", "採用率", "商談", "受注", "契約", "ARPA", "販売",
    ),
    "profitability": (
        "標準運用", "標準化", "標準提供", "標準処理", "提供原価", "原価",
        "コスト", "粗利", "自動化", "運用率", "サポート工数",
    ),
    "recurring_revenue": (
        "継続利用", "更新", "解約", "アップセル", "利用率", "アクティブ率",
        "アクティブ契約率", "NRR", "継続契約",
    ),
    "scalability_resilience": (
        "処理能力", "供給能力", "生産性", "標準運用", "自動化", "稼働率", "復旧", "継続率",
    ),
}
AI_PRODUCT_CUSTOMER_OPERATION_KPI_TERMS = (
    "顧客売上", "顧客利益", "顧客工数", "顧客作業", "利用企業売上", "利用企業利益",
    "荷主工数", "現場工数", "倉庫工数", "作業時間",
    "判断時間", "問い合わせ対応時間", "回答時間", "ピッキング", "誤出荷",
    "在庫差異", "在庫回転", "倉庫回転", "欠品率", "棚卸差異", "出荷遅延",
)
AI_PRODUCT_NRR_KPI_TERMS = (
    "NRR", "ネット収益継続", "純収益継続", "Net Revenue Retention",
)
AI_PRODUCT_NRR_MAX_PERCENT = 300.0
def _contains_ai_product_kpi_keyword(value: object, keywords: tuple[str, ...]) -> bool:
    """KPI語彙を英字の大小を区別せず照合する。"""
    text = str(value or "").casefold()
    return any(keyword.casefold() in text for keyword in keywords)
def _is_nrr_company_kpi(category: str, kpi: object) -> bool:
    """100%超が定義上あり得る純収益継続率だけを識別する。"""
    return bool(
        category == "recurring_revenue"
        and _contains_ai_product_kpi_keyword(kpi, AI_PRODUCT_NRR_KPI_TERMS)
    )
def _ai_product_headline_percentage_target(
        value: object, category: str, kpi: object) -> tuple[str, tuple[float, float] | None]:
    """会社KPIの割合を正規化し、NRRだけ100%超の計画目標を許容する。"""
    normalized = _normalize_percentage_target(value)
    standard_parts = _percentage_range_parts(normalized)
    if standard_parts:
        return normalized, standard_parts
    if not _is_nrr_company_kpi(category, kpi):
        return normalized, None
    text = re.sub(r"\s+", "", str(value or "")).replace("％", "%")
    match = re.fullmatch(
        r"(\d+(?:\.\d+)?)(?:[〜～~\-–—](\d+(?:\.\d+)?))?%",
        text,
    )
    if not match:
        return normalized, None
    low = float(match.group(1))
    high = float(match.group(2) or match.group(1))
    if not 0 < low <= high <= AI_PRODUCT_NRR_MAX_PERCENT:
        return normalized, None
    def display(number: float) -> str:
        return str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".")
    canonical = (
        f"{display(low)}%" if low == high
        else f"{display(low)}〜{display(high)}%"
    )
    return canonical, (low, high)
def _ai_product_company_metric_issues(category: str, kpi: object,
                                      leading_kpi: object) -> list[str]:
    """Reject customer-operation metrics from the company-business P6 layer."""
    issues: list[str] = []
    kpi_text = str(kpi or "").strip()
    leading_text = str(leading_kpi or "").strip()
    if not _contains_ai_product_kpi_keyword(
        kpi_text, AI_PRODUCT_BUSINESS_KPI_KEYWORDS.get(category, ()),
    ):
        issues.append("会社レベルの売上・利益・継続収益/事業成長KPIではありません。")
    if any(token in kpi_text for token in AI_PRODUCT_CUSTOMER_OPERATION_KPI_TERMS):
        issues.append("顧客・現場側の業務KPIを会社事業KPIへ転用しています。")
    if not _contains_ai_product_kpi_keyword(
        leading_text, AI_PRODUCT_BUSINESS_LEADING_KPI_KEYWORDS.get(category, ()),
    ):
        issues.append("会社が直接管理する事業先行KPIではありません。")
    if any(token in leading_text for token in AI_PRODUCT_CUSTOMER_OPERATION_KPI_TERMS):
        issues.append("顧客・現場側の業務KPIを会社事業の先行KPIへ転用しています。")
    return issues
def _normalize_ai_product_business_impact_candidate_unchecked(
        value: object, assessment: dict) -> dict:
    """Normalize the P6 company-business contract while preserving repairable issues."""
    if not isinstance(value, dict):
        return {}
    expected_categories = _ai_product_business_categories(assessment)
    raw_items = value.get("items")
    if not isinstance(raw_items, list) or len(raw_items) != 3:
        return {}
    category_order = {category: index for index, category in enumerate(expected_categories)}
    items = [copy.deepcopy(item) for item in raw_items if isinstance(item, dict)]
    if len(items) != 3:
        return {}
    # If the model respected the category vocabulary, make the display order
    # deterministic.  Unknown or duplicate categories remain invalid.
    if {str(item.get("category") or "") for item in items} == set(expected_categories):
        items.sort(key=lambda item: category_order[str(item.get("category"))])
    default_labels = {
        "revenue_growth": "売上・ARRの拡大",
        "profitability": "粗利・営業利益の改善",
        "recurring_revenue": "継続収益・NRRの向上",
        "scalability_resilience": "事業処理能力・継続性の向上",
    }
    default_roles = {
        "revenue_growth": "サービス事業責任者・営業責任者",
        "profitability": "サービス事業責任者・財務責任者",
        "recurring_revenue": "サービス事業責任者・カスタマーサクセス責任者",
        "scalability_resilience": "事業責任者・業務責任者",
    }
    normalized_items: list[dict] = []
    for index, item in enumerate(items, 1):
        category = str(item.get("category") or "").strip()
        kpi = _front_text(item.get("kpi"), 38)
        raw_target = item.get("headline_metric") or item.get("target")
        target, range_parts = _ai_product_headline_percentage_target(
            raw_target, category, kpi,
        )
        if not range_parts:
            target = _normalize_monetary_target(raw_target)
        amount_parts = _monetary_target_parts(target)
        requested_claim = str(item.get("management_target_claim_id") or "").strip()
        target_claim = _verified_management_amount_target(
            value.get("management_targets"), target,
            AI_PRODUCT_BUSINESS_KPI_KEYWORDS.get(category, ()), requested_claim,
        ) if amount_parts else None
        leading = item.get("leading_indicator") if isinstance(item.get("leading_indicator"), dict) else {}
        leading_target = _normalize_percentage_target(leading.get("target"))
        leading_parts = _percentage_range_parts(leading_target)
        formula = _front_text(item.get("formula"), 170)
        gate = item.get("decision_gate") if isinstance(item.get("decision_gate"), dict) else {}
        if (range_parts or amount_parts) and _quantitative_tokens(target) - _quantitative_tokens(formula):
            formula = f"{kpi or default_labels.get(category, '事業KPI')}を同一基準で比較し、{target}の改善を判定する。"
        go_condition = _front_text(gate.get("go_condition"), 150)
        if (range_parts or amount_parts) and _quantitative_tokens(target) - _quantitative_tokens(go_condition):
            go_condition = f"{kpi or default_labels.get(category, '事業KPI')}が{target}の事業化基準を達成すること。"
        leading_kpi = _front_text(leading.get("kpi"), 34)
        leading_formula = _front_text(leading.get("formula"), 150)
        if leading_parts and _quantitative_tokens(leading_target) - _quantitative_tokens(leading_formula):
            leading_formula = f"{leading_kpi or '事業先行KPI'}を同一基準で比較し、{leading_target}の達成を判定する。"
        owner = item.get("metric_owner") if isinstance(item.get("metric_owner"), dict) else {}
        owner_scope = str(owner.get("scope") or "").strip()
        # Provider/company are harmless vocabulary aliases.  Customer/shared
        # scopes are deliberately preserved so that contract validation rejects
        # them instead of silently laundering them into assessed-company KPIs.
        if owner_scope in {"provider", "company", "assessed company"}:
            owner_scope = "assessed_company"
        normalized_items.append({
            "business_kpi_id": f"B{index:02d}",
            "category": category,
            "label": _front_text(item.get("label"), 30) or default_labels.get(category, "事業インパクト"),
            "headline_metric": target,
            "kpi": kpi,
            "detail": _front_text(item.get("detail"), 72),
            "baseline_definition": _front_text(item.get("baseline_definition"), 155),
            "comparison_condition": _front_text(item.get("comparison_condition"), 155),
            "formula": formula,
            "target_rationale": _front_text(item.get("target_rationale"), 155),
            "customer_value_signal": _front_text(item.get("customer_value_signal"), 72),
            "causal_chain": _front_text(item.get("causal_chain"), 60),
            "leading_indicator": {
                "kpi": leading_kpi,
                "target": leading_target,
                "formula": leading_formula,
            },
            "metric_owner": {
                "scope": owner_scope,
                "role": _front_text(owner.get("role"), 60) or default_roles.get(category, "事業責任者"),
                "status": str(owner.get("status") or "confirm")
                if str(owner.get("status") or "confirm") in {"confirmed", "confirm"} else "confirm",
            },
            "decision_gate": {
                "measurement_period": _front_text(gate.get("measurement_period"), 24),
                "go_condition": go_condition,
                "stop_condition": _front_text(gate.get("stop_condition"), 145),
                "decision_owner": _front_text(gate.get("decision_owner"), 42)
                or default_roles.get(category, "事業責任者"),
            },
            "basis_type": "verified_management_target" if amount_parts else "decision_threshold",
            "management_target_claim_id": str(target_claim.get("claim_id") or "") if target_claim else requested_claim,
            "estimate_id": f"BI{index:02d}",
            "estimate_low": range_parts[0] if range_parts else (amount_parts[0] if amount_parts else None),
            "estimate_high": range_parts[1] if range_parts else (amount_parts[1] if amount_parts else None),
            "estimate_unit": "%" if range_parts else (amount_parts[2] if amount_parts else ""),
        })
    normalized = {
        "schema_version": AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION,
        "status": "decision_thresholds",
        "organization_role": _business_value_model_role(assessment),
        "lead": _front_text(value.get("lead"), 220),
        "management_targets": copy.deepcopy(value.get("management_targets", [])),
        "items": normalized_items,
    }
    return normalized
def normalize_ai_product_business_impact_candidate_with_issues(
        value: object, assessment: dict) -> tuple[dict, list[str]]:
    """Return the normalized candidate and exact contract issues for LLM repair.
    The public strict normalizer intentionally returns an empty object for an
    invalid contract.  Generation retries, however, need the normalized object
    to retain deterministic repairs and to report the actual semantic defect
    instead of validating the unnormalized model response.
    """
    normalized = _normalize_ai_product_business_impact_candidate_unchecked(
        value, assessment,
    )
    if not normalized:
        return {}, ai_product_business_impact_contract_issues(value, assessment)
    return normalized, ai_product_business_impact_contract_issues(
        normalized, assessment,
    )
def normalize_ai_product_business_impact_candidate(
        value: object, assessment: dict) -> dict:
    """Normalize the P6 company-business contract without accepting defects."""
    normalized, issues = normalize_ai_product_business_impact_candidate_with_issues(
        value, assessment,
    )
    return normalized if normalized and not issues else {}
def ai_product_business_impact_contract_issues(
        value: object, assessment: dict) -> list[str]:
    """Validate that P6 is company-owned and independent of P1/P2/P3."""
    if not isinstance(value, dict):
        return ["AI製品事業インパクトがJSONオブジェクトではありません。"]
    issues: list[str] = []
    if str(value.get("schema_version") or "") != AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION:
        issues.append("schema_versionが不正です。")
    if str(value.get("status") or "") != "decision_thresholds":
        issues.append("statusはdecision_thresholdsである必要があります。")
    expected_categories = _ai_product_business_categories(assessment)
    items = value.get("items")
    if not isinstance(items, list) or len(items) != 3:
        return issues + ["会社レベルの事業KPIは3件必要です。"]
    categories = [str(item.get("category") or "") for item in items if isinstance(item, dict)]
    if tuple(categories) != tuple(expected_categories):
        issues.append("事業KPIカテゴリが売上・利益・継続収益/事業継続性の順と一致しません。")
    for index, item in enumerate(items, 1):
        path = f"items[{index}]"
        if not isinstance(item, dict):
            issues.append(f"{path}がJSONオブジェクトではありません。")
            continue
        for forbidden in ("priority", "use_case_id", "theme"):
            if forbidden in item:
                issues.append(f"{path}.{forbidden}はP6の会社事業KPI契約に含めません。")
        required = (
            "business_kpi_id", "category", "label", "headline_metric", "kpi", "detail",
            "baseline_definition", "comparison_condition", "formula", "target_rationale",
            "customer_value_signal", "causal_chain", "basis_type", "estimate_id",
        )
        missing = [field for field in required if not str(item.get(field) or "").strip()]
        if missing:
            issues.append(f"{path}の必須項目が不足しています: {', '.join(missing)}")
        expected_id = f"B{index:02d}"
        if str(item.get("business_kpi_id") or "") != expected_id:
            issues.append(f"{path}.business_kpi_idは{expected_id}です。")
        category = str(item.get("category") or "")
        leading = item.get("leading_indicator")
        leading_kpi = leading.get("kpi") if isinstance(leading, dict) else ""
        for semantic_issue in _ai_product_company_metric_issues(
            category, item.get("kpi"), leading_kpi,
        ):
            issues.append(f"{path}: {semantic_issue}")
        target = str(item.get("headline_metric") or "")
        parts = _ai_product_headline_percentage_target(
            target, category, item.get("kpi"),
        )[1]
        amount_parts = _monetary_target_parts(target)
        target_claim = _verified_management_amount_target(
            value.get("management_targets"), target,
            AI_PRODUCT_BUSINESS_KPI_KEYWORDS.get(category, ()), item.get("management_target_claim_id"),
        ) if amount_parts else None
        if not parts and not target_claim:
            issues.append(
                f"{path}.headline_metricは0%超100%以下の割合目標、"
                "NRRの場合は300%以下の割合目標、または原典照合済み"
                "中期経営目標と一致する金額が必要です。"
            )
        elif (_quantitative_tokens(target) - _quantitative_tokens(item.get("formula"))
              or _quantitative_tokens(target) - _quantitative_tokens(
                  (item.get("decision_gate") or {}).get("go_condition")
                  if isinstance(item.get("decision_gate"), dict) else ""
              )):
            issues.append(f"{path}の数値目標が算定式またはGo条件へ継承されていません。")
        owner = item.get("metric_owner")
        if (not isinstance(owner, dict) or owner.get("scope") != "assessed_company"
                or not str(owner.get("role") or "").strip()
                or owner.get("status") not in {"confirmed", "confirm"}):
            issues.append(f"{path}.metric_ownerは評価対象企業の責任者である必要があります。")
        elif any(token in str(owner.get("role") or "") for token in ("顧客", "荷主", "利用企業")):
            issues.append(f"{path}.metric_owner.roleを顧客側の責任者にはできません。")
        if (not isinstance(leading, dict) or not str(leading.get("kpi") or "").strip()
                or not _percentage_range_parts(leading.get("target"))
                or _quantitative_tokens(leading.get("target")) - _quantitative_tokens(leading.get("formula"))):
            issues.append(f"{path}.leading_indicatorの数値・算定式が不足しています。")
        gate = item.get("decision_gate")
        if not isinstance(gate, dict) or not all(str(gate.get(field) or "").strip() for field in (
            "measurement_period", "go_condition", "stop_condition", "decision_owner",
        )):
            issues.append(f"{path}.decision_gateが不足しています。")
        elif any(token in str(gate.get("decision_owner") or "") for token in ("顧客", "荷主", "利用企業")):
            issues.append(f"{path}.decision_gate.decision_ownerを顧客側の責任者にはできません。")
        expected_basis = "verified_management_target" if target_claim else "decision_threshold"
        expected_parts = amount_parts if target_claim else ((*parts, "%") if parts else None)
        if item.get("basis_type") != expected_basis:
            issues.append(f"{path}.basis_typeは{expected_basis}です。")
        if target_claim and str(item.get("management_target_claim_id") or "") != str(target_claim.get("claim_id")):
            issues.append(f"{path}.management_target_claim_idが中期経営目標と一致しません。")
        if expected_parts and (
            item.get("estimate_low") != expected_parts[0]
            or item.get("estimate_high") != expected_parts[1]
            or item.get("estimate_unit") != expected_parts[2]
        ):
            issues.append(f"{path}のestimate値がheadline_metricと一致しません。")
    return issues
def ai_product_business_impact_for(assessment: dict) -> dict:
    """Return the validated company-level AI business impact contract."""
    return normalize_ai_product_business_impact_candidate(
        assessment.get("ai_product_business_impact"), assessment,
    )
def normalize_quantitative_hypothesis_candidate(
        value: object, priority_pocs: list[dict[str, str]]) -> tuple[dict, list[str]]:
    """Repair schema drift without inventing a business or product target.
    The executive headline is one customer/provider business KPI.  The product
    KPI remains a separately measured leading indicator.  Existing values may
    be copied between a layer and its top-level display fields, but a missing
    numeric value is never synthesized by Python.
    """
    if not isinstance(value, dict):
        return {}, []
    candidate = copy.deepcopy(value)
    hypotheses = candidate.get("hypotheses")
    if not isinstance(hypotheses, list) or len(hypotheses) != 3 or len(priority_pocs) != 3:
        return candidate, []
    changes: list[str] = []
    priorities = [str(item.get("priority") or "") for item in hypotheses if isinstance(item, dict)]
    if len(priorities) == 3 and set(priorities) == {"P1", "P2", "P3"}:
        sorted_hypotheses = sorted(hypotheses, key=lambda item: int(str(item.get("priority"))[1:]))
        if sorted_hypotheses != hypotheses:
            hypotheses[:] = sorted_hypotheses
            changes.append("hypotheses:canonical_priority_order")
    for index, item in enumerate(hypotheses, 1):
        if not isinstance(item, dict):
            continue
        expected = priority_pocs[index - 1]
        for field in ("priority", "use_case_id", "theme"):
            # 欠損だけを補う。別PoCの非空ID・名称を上書きして誤対応を
            # 隠すことはせず、後段の契約検証で拒否する。
            if not str(item.get(field) or "").strip():
                item[field] = str(expected[field])
                changes.append(f"hypotheses[{index}].{field}:canonical_binding")
        display_layer_name = str(item.get("display_kpi_layer") or "")
        business_numeric_layers = [
            name for name in BUSINESS_EFFECT_DISPLAY_LAYERS
            if isinstance(item.get(name), dict)
            and _is_business_effect_metric(item[name].get("target"))
        ]
        if display_layer_name not in BUSINESS_EFFECT_DISPLAY_LAYERS and len(business_numeric_layers) == 1:
            display_layer_name = business_numeric_layers[0]
            item["display_kpi_layer"] = display_layer_name
            changes.append(f"hypotheses[{index}].display_kpi_layer:business_layer_selected")
        display_layer = item.get(display_layer_name) if isinstance(item.get(display_layer_name), dict) else {}
        normalized_metric = _normalize_percentage_target(item.get("headline_metric"))
        normalized_layer_target = _normalize_percentage_target(display_layer.get("target"))
        if normalized_metric != str(item.get("headline_metric") or ""):
            item["headline_metric"] = normalized_metric
            changes.append(f"hypotheses[{index}].headline_metric:format_normalized")
        if normalized_layer_target != str(display_layer.get("target") or ""):
            display_layer["target"] = normalized_layer_target
            changes.append(f"hypotheses[{index}].{display_layer_name}.target:format_normalized")
        headline_metric = str(item.get("headline_metric") or "").strip()
        if not headline_metric and _is_business_effect_metric(display_layer.get("target")):
            headline_metric = str(display_layer["target"])
            item["headline_metric"] = headline_metric
            changes.append(f"hypotheses[{index}].headline_metric:copied_from_display_layer")
        elif headline_metric and not str(display_layer.get("target") or "").strip():
            display_layer["target"] = headline_metric
            changes.append(f"hypotheses[{index}].{display_layer_name}.target:copied_from_headline")
        if not str(item.get("kpi") or "").strip() and str(display_layer.get("kpi") or "").strip():
            item["kpi"] = str(display_layer["kpi"])
            changes.append(f"hypotheses[{index}].kpi:copied_from_display_layer")
        elif str(item.get("kpi") or "").strip() and not str(display_layer.get("kpi") or "").strip():
            display_layer["kpi"] = str(item["kpi"])
            changes.append(f"hypotheses[{index}].{display_layer_name}.kpi:copied_from_headline")
        expected_scope = "provider" if display_layer_name == "provider_business_kpi" else "shared"
        expected_attribution = "enabling" if display_layer_name == "provider_business_kpi" else "contributory"
        expected_beneficiary = "provider" if display_layer_name == "provider_business_kpi" else "shared"
        if display_layer_name in BUSINESS_EFFECT_DISPLAY_LAYERS:
            display_layer["attribution_level"] = expected_attribution
            layer_owner = display_layer.get("metric_owner") if isinstance(display_layer.get("metric_owner"), dict) else {}
            if display_layer_name == "provider_business_kpi":
                layer_owner["scope"] = "provider"
                layer_owner.setdefault("role", "サービス事業責任者")
            else:
                if str(layer_owner.get("scope") or "") not in {"customer", "shared"}:
                    layer_owner["scope"] = expected_scope
                layer_owner.setdefault("role", "顧客業務責任者・サービス責任者")
            layer_owner.setdefault("status", "confirm")
            display_layer["metric_owner"] = layer_owner
            item[display_layer_name] = display_layer
            item["metric_owner"] = copy.deepcopy(layer_owner)
            item["attribution_level"] = expected_attribution
            item["beneficiary"] = expected_beneficiary
        kpi = str(item.get("kpi") or "").strip()
        product = item.get("product_kpi") if isinstance(item.get("product_kpi"), dict) else {}
        product_target = _normalize_percentage_target(product.get("target"))
        if product_target != str(product.get("target") or ""):
            product["target"] = product_target
            changes.append(f"hypotheses[{index}].product_kpi.target:format_normalized")
        product["attribution_level"] = "direct"
        product_owner = product.get("metric_owner") if isinstance(product.get("metric_owner"), dict) else {}
        product_owner.update({"scope": "provider"})
        product_owner.setdefault("role", "製品・AI機能の測定責任者")
        product_owner.setdefault("status", "confirm")
        product["metric_owner"] = product_owner
        if not str(product.get("baseline_definition") or "").strip() and str(
            item.get("baseline_definition") or ""
        ).strip():
            product["baseline_definition"] = str(item["baseline_definition"])
        if not str(product.get("comparison_condition") or "").strip() and str(
            item.get("comparison_condition") or ""
        ).strip():
            product["comparison_condition"] = str(item["comparison_condition"])
        if not str(product.get("causal_link") or "").strip():
            product["causal_link"] = "AI機能の利用有無が製品KPIを直接変える。"
        headline_tokens = _quantitative_tokens(headline_metric)
        product_tokens = _quantitative_tokens(product.get("target"))
        if headline_tokens and not headline_tokens <= _quantitative_tokens(item.get("formula")):
            item["formula"] = f"{kpi or '事業KPI'}を同一条件で比較し、{headline_metric}の改善を判定する。"
            changes.append(f"hypotheses[{index}].formula:headline_metric_inherited")
        if headline_tokens and display_layer_name in BUSINESS_EFFECT_DISPLAY_LAYERS \
                and not headline_tokens <= _quantitative_tokens(display_layer.get("formula")):
            display_layer["formula"] = f"{kpi or '事業KPI'}を同一条件で比較し、{headline_metric}の改善を判定する。"
            item[display_layer_name] = display_layer
            changes.append(f"hypotheses[{index}].{display_layer_name}.formula:headline_metric_inherited")
        if product_tokens and not product_tokens <= _quantitative_tokens(product.get("formula")):
            product["formula"] = (
                f"{str(product.get('kpi') or '製品KPI')}を同一条件で比較し、"
                f"{str(product.get('target') or '')}の改善を判定する。"
            )
            changes.append(f"hypotheses[{index}].product_kpi.formula:product_metric_inherited")
        item["product_kpi"] = product
        for layer_name, scope, attribution, role, formula in (
            (
                "customer_outcome_kpi", "customer", "contributory", "顧客業務責任者",
                "顧客成果変化率＝（顧客PoC実績値－顧客基準実績値）÷顧客基準実績値×100%",
            ),
            (
                "provider_business_kpi", "provider", "enabling", "サービス事業責任者",
                "提供者成果変化率＝（提供者PoC実績値－提供者基準実績値）÷提供者基準実績値×100%",
            ),
        ):
            layer = item.get(layer_name) if isinstance(item.get(layer_name), dict) else {}
            if layer_name == display_layer_name:
                continue
            if _quantitative_tokens(layer.get("target")) or not any(
                label in re.sub(r"\s+", "", str(layer.get("target") or ""))
                for label in ("PoCで確定", "PoCで算定", "PoCで測定")
            ):
                layer["target"] = "PoCで確定"
                changes.append(f"hypotheses[{index}].{layer_name}.target:unsupported_number_removed")
            if not _is_variable_measurement_formula(layer.get("formula")):
                layer["formula"] = formula
                changes.append(f"hypotheses[{index}].{layer_name}.formula:variable_formula_applied")
            layer["attribution_level"] = attribution
            owner = layer.get("metric_owner") if isinstance(layer.get("metric_owner"), dict) else {}
            owner.update({"scope": scope})
            owner.setdefault("role", role)
            owner.setdefault("status", "confirm")
            layer["metric_owner"] = owner
            item[layer_name] = layer
        gate = item.get("poc_gate") if isinstance(item.get("poc_gate"), dict) else {}
        gate.setdefault("measurement_period", "要確認")
        if headline_tokens and not headline_tokens <= _quantitative_tokens(gate.get("business_go_condition")):
            gate["business_go_condition"] = f"事業KPIを同一条件で比較し、{headline_metric}を達成する。"
            changes.append(f"hypotheses[{index}].poc_gate.business_go_condition:headline_metric_inherited")
        if product_tokens and not product_tokens <= _quantitative_tokens(gate.get("leading_kpi_condition")):
            gate["leading_kpi_condition"] = (
                f"製品先行KPIを同一条件で比較し、{str(product.get('target') or '')}を達成する。"
            )
            changes.append(f"hypotheses[{index}].poc_gate.leading_kpi_condition:product_metric_inherited")
        combined_tokens = _quantitative_tokens(gate.get("go_condition"))
        if (headline_tokens and not headline_tokens <= combined_tokens) or (
                product_tokens and not product_tokens <= combined_tokens):
            gate["go_condition"] = (
                f"事業効果{headline_metric}、製品先行KPI{str(product.get('target') or '')}をともに達成する。"
            )
            changes.append(f"hypotheses[{index}].poc_gate.go_condition:dual_metric_inherited")
        gate.setdefault("stop_condition", "品質・安全性または運用負荷が許容条件を外れた場合は停止する。")
        gate.setdefault("decision_owner", "要確認")
        item["poc_gate"] = gate
    return candidate, changes
def _unverified_downstream_kpi_issues(value: object, path: str) -> list[str]:
    """Prevent unsupported downstream business numbers in LLM decision thresholds."""
    if not isinstance(value, dict):
        return [f"{path}はKPIオブジェクトである必要があります。"]
    target = re.sub(r"\s+", "", str(value.get("target") or ""))
    formula = str(value.get("formula") or "")
    if _quantitative_tokens(target):
        return [f"{path}.targetに未検証の固定数値を設定できません。PoCで確定または実績変数式にしてください。"]
    target_is_pending = any(label in target for label in ("PoCで確定", "PoCで算定", "PoCで測定"))
    target_is_formula = _is_variable_measurement_formula(target)
    issues = []
    if not (target_is_pending or target_is_formula):
        issues.append(f"{path}.targetはPoCで確定または実績変数式にしてください。")
    if not _is_variable_measurement_formula(formula):
        issues.append(f"{path}.formulaは固定数値ではなく、顧客・提供者の実績値を用いる変数式にしてください。")
    return issues
def quantitative_hypothesis_contract_issues(
        value: object, priority_pocs: list[dict[str, str]] | None = None) -> list[str]:
    """Validate the business-effect headline and its product leading KPI."""
    if not isinstance(value, dict):
        return ["JSONオブジェクトではありません。"]
    hypotheses = value.get("hypotheses")
    if not isinstance(hypotheses, list) or len(hypotheses) != 3:
        return ["hypothesesは必ず3件必要です。"]
    required = (
        "priority", "use_case_id", "theme", "business_outcome", "headline_metric", "kpi",
        "baseline_definition", "comparison_condition", "detail", "formula", "poc_gate",
        "assumption", "beneficiary", "metric_owner", "attribution_level", "causal_link",
        "product_kpi", "customer_outcome_kpi", "provider_business_kpi",
    )
    issues: list[str] = []
    kpis: set[str] = set()
    product_kpis: set[str] = set()
    for index, item in enumerate(hypotheses, 1):
        if not isinstance(item, dict):
            issues.append(f"hypotheses[{index}]がオブジェクトではありません。")
            continue
        missing = [field for field in required if not str(item.get(field) or "").strip()]
        if missing:
            issues.append(f"hypotheses[{index}]の必須項目が不足しています: {', '.join(missing)}")
            continue
        if priority_pocs:
            expected = priority_pocs[index - 1]
            if str(item.get("priority")) != expected["priority"]:
                issues.append(f"hypotheses[{index}].priorityが選定PoCと一致しません。")
            if str(item.get("use_case_id")) != expected["use_case_id"]:
                issues.append(f"hypotheses[{index}].use_case_idが選定PoCと一致しません。")
            if normalized_use_case_label(item.get("theme")) != normalized_use_case_label(expected["theme"]):
                issues.append(f"hypotheses[{index}].themeが選定PoCと一致しません。")
        kpi = re.sub(r"\s+", "", str(item["kpi"]))
        if kpi in kpis:
            issues.append(f"hypotheses[{index}]の事業効果KPIが他項目と重複しています。")
        kpis.add(kpi)
        headline_tokens = _quantitative_tokens(item["headline_metric"])
        formula_tokens = _quantitative_tokens(item["formula"])
        gate_tokens = _quantitative_tokens(json.dumps(item["poc_gate"], ensure_ascii=False))
        range_match = re.fullmatch(
            r"\s*(\d+(?:\.\d+)?)\s*(?:[〜～~\-–—]\s*(\d+(?:\.\d+)?)\s*)?[％%]\s*",
            str(item["headline_metric"]),
        )
        if not range_match or not _is_business_effect_metric(item["headline_metric"]):
            issues.append(
                f"hypotheses[{index}]のheadline_metricは0%超100%以下の割合または割合レンジである必要があります。"
            )
        else:
            low = float(range_match.group(1))
            high = float(range_match.group(2) or range_match.group(1))
            if not (0 < low <= high <= 100):
                issues.append(
                    f"hypotheses[{index}]のheadline_metricは0%超100%以下かつ下限<=上限である必要があります。"
                )
        if not headline_tokens:
            issues.append(f"hypotheses[{index}]のheadline_metricに数値と単位がありません。")
        elif not headline_tokens <= formula_tokens:
            issues.append(f"hypotheses[{index}]の目標値がformulaへ同じ表現で継承されていません。")
        elif not headline_tokens <= gate_tokens:
            issues.append(f"hypotheses[{index}]の目標値がpoc_gateへ同じ表現で継承されていません。")
        if str(item.get("beneficiary") or "") not in {"customer", "provider", "shared"}:
            issues.append(f"hypotheses[{index}].beneficiaryはcustomer/provider/sharedのいずれかです。")
        issues.extend(_quantitative_owner_issues(item.get("metric_owner"), f"hypotheses[{index}].metric_owner"))
        display_layer_name = str(item.get("display_kpi_layer") or "")
        if display_layer_name not in BUSINESS_EFFECT_DISPLAY_LAYERS:
            issues.append(
                f"hypotheses[{index}].display_kpi_layerはcustomer_outcome_kpiまたは"
                "provider_business_kpiである必要があります。"
            )
        product = item.get("product_kpi")
        issues.extend(_quantitative_layer_issues(product, f"hypotheses[{index}].product_kpi", require_numeric=True))
        customer = item.get("customer_outcome_kpi")
        provider = item.get("provider_business_kpi")
        issues.extend(_quantitative_layer_issues(
            customer, f"hypotheses[{index}].customer_outcome_kpi",
            require_numeric=display_layer_name == "customer_outcome_kpi",
        ))
        issues.extend(_quantitative_layer_issues(
            provider, f"hypotheses[{index}].provider_business_kpi",
            require_numeric=display_layer_name == "provider_business_kpi",
        ))
        if display_layer_name != "customer_outcome_kpi":
            issues.extend(_unverified_downstream_kpi_issues(
                customer, f"hypotheses[{index}].customer_outcome_kpi",
            ))
        if display_layer_name != "provider_business_kpi":
            issues.extend(_unverified_downstream_kpi_issues(
                provider, f"hypotheses[{index}].provider_business_kpi",
            ))
        if isinstance(product, dict):
            product_kpi_name = re.sub(r"\s+", "", str(product.get("kpi") or ""))
            if product_kpi_name in product_kpis:
                issues.append(f"hypotheses[{index}]の製品先行KPIが他項目と重複しています。")
            product_kpis.add(product_kpi_name)
            product_owner = product.get("metric_owner") if isinstance(product.get("metric_owner"), dict) else {}
            if str(product_owner.get("scope") or "") != "provider":
                issues.append(f"hypotheses[{index}].product_kpi.metric_owner.scopeはproviderです。")
            if str(product.get("attribution_level") or "") != "direct":
                issues.append(f"hypotheses[{index}].product_kpi.attribution_levelはdirectです。")
        top_owner = item.get("metric_owner") if isinstance(item.get("metric_owner"), dict) else {}
        if isinstance(customer, dict):
            customer_owner = customer.get("metric_owner") if isinstance(customer.get("metric_owner"), dict) else {}
            if str(customer_owner.get("scope") or "") not in {"customer", "shared"}:
                issues.append(f"hypotheses[{index}].customer_outcome_kpiの責任範囲はcustomer/sharedです。")
            if str(customer.get("attribution_level") or "") != "contributory":
                issues.append(f"hypotheses[{index}].customer_outcome_kpi.attribution_levelはcontributoryです。")
        if isinstance(provider, dict):
            provider_owner = provider.get("metric_owner") if isinstance(provider.get("metric_owner"), dict) else {}
            if str(provider_owner.get("scope") or "") != "provider":
                issues.append(f"hypotheses[{index}].provider_business_kpi.metric_owner.scopeはproviderです。")
            if str(provider.get("attribution_level") or "") != "enabling":
                issues.append(f"hypotheses[{index}].provider_business_kpi.attribution_levelはenablingです。")
        display_layer = item.get(display_layer_name) if isinstance(item.get(display_layer_name), dict) else {}
        if display_layer:
            if str(display_layer.get("target") or "").strip() != str(item.get("headline_metric") or "").strip():
                issues.append(f"hypotheses[{index}]のheadline_metricは表示する事業KPIのtargetと一致させてください。")
            if re.sub(r"\s+", "", str(display_layer.get("kpi") or "")) != kpi:
                issues.append(f"hypotheses[{index}]のkpiは表示する事業KPIのkpiと一致させてください。")
            expected_scope = {"customer", "shared"} if display_layer_name == "customer_outcome_kpi" else {"provider"}
            expected_attribution = "contributory" if display_layer_name == "customer_outcome_kpi" else "enabling"
            expected_beneficiaries = {"customer", "shared"} if display_layer_name == "customer_outcome_kpi" else {"provider"}
            if (str(top_owner.get("scope") or "") not in expected_scope
                    or str(item.get("attribution_level") or "") != expected_attribution
                    or str(item.get("beneficiary") or "") not in expected_beneficiaries):
                issues.append(f"hypotheses[{index}]の事業効果KPIの責任・帰属・受益者が表示層と一致しません。")
        gate = item.get("poc_gate")
        gate_required = (
            "measurement_period", "go_condition", "business_go_condition",
            "leading_kpi_condition", "stop_condition", "decision_owner",
        )
        if not isinstance(gate, dict) or not all(str(gate.get(field) or "").strip() for field in gate_required):
            issues.append(
                f"hypotheses[{index}].poc_gateの測定期間・事業効果条件・製品先行KPI条件・"
                "停止条件・判定責任者が不足しています。"
            )
        elif isinstance(product, dict):
            product_tokens = _quantitative_tokens(product.get("target"))
            if not product_tokens:
                issues.append(f"hypotheses[{index}].product_kpi.targetに数値目標がありません。")
            elif not product_tokens <= _quantitative_tokens(product.get("formula")):
                issues.append(f"hypotheses[{index}]の製品先行KPI値がproduct_kpi.formulaへ継承されていません。")
            elif not product_tokens <= _quantitative_tokens(gate.get("leading_kpi_condition")):
                issues.append(f"hypotheses[{index}]の製品先行KPI値がleading_kpi_conditionへ継承されていません。")
            elif not product_tokens <= gate_tokens:
                issues.append(f"hypotheses[{index}]の製品先行KPI値がpoc_gateへ継承されていません。")
        if display_layer and headline_tokens and not headline_tokens <= _quantitative_tokens(display_layer.get("formula")):
            issues.append(f"hypotheses[{index}]の事業効果目標が表示KPIのformulaへ継承されていません。")
        if isinstance(gate, dict) and headline_tokens \
                and not headline_tokens <= _quantitative_tokens(gate.get("business_go_condition")):
            issues.append(f"hypotheses[{index}]の事業効果目標がbusiness_go_conditionへ継承されていません。")
    return issues
def materialize_assessment_decision_contract(assessment: dict, front_matter: dict | None = None) -> dict:
    """通常生成・既存JSON再利用で共通の意思決定データを確定する。
    描画時だけに派生させると、レビューJSONとPPTXの内容がずれる。優先PoC、
    評価仮説・開始ゲート、PoCチャーター、統制、二層価値、技術提案をここで同じ
    入力から materialize し、旧JSONを再利用する場合にも不足項目を安全な
    ``confirm`` として可視化する。
    """
    front = front_matter if isinstance(front_matter, dict) else consulting_front_matter_for(assessment)
    assessment["consulting_front_matter"] = front
    priority_decision = materialize_poc_priority_decision(assessment, front)
    portfolio = materialize_poc_portfolio(assessment, front)
    decision_data = materialize_poc_decision_data(assessment, front)
    business_value_model = materialize_business_value_model(assessment, front)
    assessment["technical_proposal"] = technical_proposal_for(assessment)
    # 優先テーマごとの実現ロジックもレビューJSONへ凍結する。描画時だけ
    # フォールバックすると、異常検知・最適化・RAGの違いをJSONレビューで
    # 確認できず、PPTXとの内容差分が生じるためである。
    assessment["poc_logic_details"] = poc_logic_details_for(assessment)
    # LLMで作成した数値案は、顧客実績や外部効果事例ではなくPoCの判定基準として
    # 専用契約へ凍結する。既存のv2契約を上書きせず、旧JSONだけ安全なKPI設計へ戻す。
    measurement_design = poc_measurement_design_for(assessment)
    # A numeric contract that exists but no longer matches P1-P3 must remain
    # visible to JSON validation.  Replacing it with a generic pre-PoC page
    # silently discards the very business-effect targets the review approved.
    if measurement_design.get("status") != "invalid_numeric_contract":
        assessment["poc_measurement_design"] = measurement_design
    scorecard = materialize_poc_selection_scorecard(assessment, front)
    start_readiness = materialize_poc_start_readiness(assessment, front)
    result = {
        "poc_priority_decision": priority_decision,
        "poc_portfolio": portfolio,
        **decision_data,
        "business_value_model": business_value_model,
        "technical_proposal": assessment["technical_proposal"],
        "poc_measurement_design": assessment["poc_measurement_design"],
    }
    if scorecard:
        result["poc_selection_scorecard"] = scorecard
    if start_readiness:
        result["poc_start_readiness"] = start_readiness
    return result
def _customer_measurement_text(value: object, limit: int = 180) -> str:
    """内部の根拠区分を、顧客向けのPoC目標表現へ正規化する。"""
    text = _front_text(value, limit)
    replacements = (
        ("生成AIによるPoC効果仮説", "AI導入のPoC目標"),
        ("PoC効果仮説", "PoC目標"),
        ("効果仮説", "目標効果"),
        ("PoC仮説", "PoC目標"),
        ("仮説値", "目標値"),
        ("仮説", "目標"),
        ("公開実績ではありません", "PoCで実測して本番化を判断します"),
        ("公開実績ではない", "PoCで実測する"),
    )
    for before, after in replacements:
        text = text.replace(before, after)
    return re.sub(r"\s+", " ", text).strip(" 。")
def _normalize_measurement_owner(value: object) -> dict:
    if not isinstance(value, dict):
        return {}
    scope = str(value.get("scope") or "")
    role = _customer_measurement_text(value.get("role"), 58)
    status = str(value.get("status") or "")
    if scope not in {"customer", "provider", "shared"} or not role or status not in {"confirmed", "confirm"}:
        return {}
    return {"scope": scope, "role": role, "status": status}
def _normalize_measurement_kpi_layer(value: object, *, numeric_target: bool) -> dict:
    if not isinstance(value, dict):
        return {}
    required = ("kpi", "target", "baseline_definition", "comparison_condition", "formula", "causal_link")
    if not all(str(value.get(field) or "").strip() for field in required):
        return {}
    owner = _normalize_measurement_owner(value.get("metric_owner"))
    attribution = str(value.get("attribution_level") or "")
    target = _customer_measurement_text(value.get("target"), 28)
    if (not owner or attribution not in {"direct", "contributory", "enabling"}
            or (numeric_target and not _is_business_effect_metric(target))):
        return {}
    return {
        "kpi": _customer_measurement_text(value.get("kpi"), 48),
        "target": target,
        "baseline_definition": _customer_measurement_text(value.get("baseline_definition"), 92),
        "comparison_condition": _customer_measurement_text(value.get("comparison_condition"), 88),
        "formula": _customer_measurement_text(value.get("formula"), 96),
        "metric_owner": owner,
        "attribution_level": attribution,
        "causal_link": _customer_measurement_text(value.get("causal_link"), 96),
    }
def _legacy_measurement_design_to_v5(value: dict, assessment: dict | None) -> dict:
    """Upgrade legacy designs without moving a number to a different KPI layer.
    Schema v4 already contains the three responsibility layers, but its
    headline is the product KPI.  Preserve that meaning and mark it as a legacy
    product headline.  Schema v2/v3 is upgraded only when every row maps
    bijectively to the selected P1-P3; positional remapping remains forbidden.
    """
    if str(value.get("schema_version") or "") == "4":
        upgraded = copy.deepcopy(value)
        upgraded["schema_version"] = POC_MEASUREMENT_DESIGN_SCHEMA_VERSION
        upgraded["migration_status"] = "schema4_product_headline"
        return upgraded
    pocs = quantitative_priority_pocs(assessment or {}) if isinstance(assessment, dict) else []
    items = value.get("items")
    if len(pocs) != 3 or not isinstance(items, list) or len(items) != 3:
        return {}
    matched: list[dict] = []
    used: set[int] = set()
    for poc in pocs:
        candidates: list[tuple[int, dict]] = []
        for index, item in enumerate(items):
            if index in used or not isinstance(item, dict):
                continue
            explicit = (
                str(item.get("priority") or "") == poc["priority"]
                and str(item.get("use_case_id") or "") == poc["use_case_id"]
                and normalized_use_case_label(item.get("theme")) == normalized_use_case_label(poc["theme"])
            )
            labels = [item.get("business_outcome")]
            if isinstance(item.get("supporting_use_cases"), list):
                labels.extend(item["supporting_use_cases"])
            named = any(
                normalized_use_case_label(label) == normalized_use_case_label(poc["theme"])
                for label in labels if str(label or "").strip()
            )
            if explicit or named:
                candidates.append((index, item))
        if len(candidates) != 1:
            return {}
        index, item = candidates[0]
        used.add(index)
        metric = _customer_measurement_text(item.get("headline_metric"), 28)
        kpi = _customer_measurement_text(item.get("kpi"), 48)
        formula = _customer_measurement_text(item.get("formula"), 96)
        old_gate = _customer_measurement_text(item.get("poc_gate"), 96)
        if not metric or not kpi or not formula or not old_gate:
            return {}
        owner = {"scope": "provider", "role": "製品・AI機能の測定責任者（要確認）", "status": "confirm"}
        customer_owner = {"scope": "customer", "role": "対象業務の責任者（要確認）", "status": "confirm"}
        provider_owner = {"scope": "provider", "role": "サービス事業責任者（要確認）", "status": "confirm"}
        matched.append({
            "priority": poc["priority"], "use_case_id": poc["use_case_id"], "theme": poc["theme"],
            "business_outcome": _customer_measurement_text(item.get("business_outcome"), 48),
            "headline_metric": metric, "kpi": kpi,
            "detail": _customer_measurement_text(item.get("detail"), 120),
            "baseline_definition": _customer_measurement_text(item.get("baseline_definition"), 92),
            "comparison_condition": _customer_measurement_text(item.get("comparison_condition"), 88),
            "formula": formula,
            "poc_gate": {
                "measurement_period": "旧契約の比較期間をPoC開始前に確認",
                "go_condition": old_gate,
                "stop_condition": "品質・安全性・運用負荷の許容条件をPoC開始前に確認",
                "decision_owner": "要確認",
            },
            "target_rationale": _customer_measurement_text(item.get("target_rationale"), 110),
            "display_kpi_layer": "product_kpi",
            "beneficiary": "shared", "metric_owner": owner,
            "attribution_level": "direct",
            "causal_link": f"{poc['theme']}の機能KPIを実測し、顧客業務成果と提供者事業成果への寄与を段階的に確認する。",
            "product_kpi": {
                "kpi": kpi, "target": metric,
                "baseline_definition": _customer_measurement_text(item.get("baseline_definition"), 92),
                "comparison_condition": _customer_measurement_text(item.get("comparison_condition"), 88),
                "formula": formula, "metric_owner": owner, "attribution_level": "direct",
                "causal_link": "AI機能の利用有無を同一条件で比較し、製品KPIへの直接効果を確認する。",
            },
            "customer_outcome_kpi": {
                "kpi": _customer_measurement_text(item.get("business_outcome"), 48) or "顧客業務成果",
                "target": "PoCで確定", "baseline_definition": "対象業務の現状実績をPoC開始前に確定",
                "comparison_condition": "製品KPI達成時の対象業務を現行運用と比較",
                "formula": "対象業務の実績値を用いてPoCで確定", "metric_owner": customer_owner,
                "attribution_level": "contributory", "causal_link": "製品KPI改善が顧客業務成果へ寄与するかを検証する。",
            },
            "provider_business_kpi": {
                "kpi": "利用定着・継続利用・運用原価", "target": "PoCで確定",
                "baseline_definition": "提供者側の現状実績をPoC開始前に確定",
                "comparison_condition": "PoC利用群と現行提供条件を比較",
                "formula": "提供者側の実績値を用いてPoCで確定", "metric_owner": provider_owner,
                "attribution_level": "enabling", "causal_link": "顧客成果と利用定着を提供者事業KPIへ接続して確認する。",
            },
            "basis_type": "legacy_migrated",
        })
    return {
        "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
        "status": "decision_thresholds", "migration_status": "legacy_migrated",
        "lead": _customer_measurement_text(value.get("lead"), 220),
        "management_targets": value.get("management_targets", []), "items": matched,
    }
def normalize_poc_measurement_design(value: object, assessment: dict | None = None) -> dict:
    """Normalize schema v5 and safely read older schemas without value reassignment."""
    if not isinstance(value, dict):
        return {}
    status = str(value.get("status") or "").strip()
    items = value.get("items")
    if not isinstance(items, list):
        return {}
    if status == "pre_poc":
        required = ("dimension", "kpi", "baseline", "comparison", "formula", "data")
        normalized_items = []
        for item in items:
            if not isinstance(item, dict) or not all(str(item.get(key) or "").strip() for key in required):
                return {}
            normalized_items.append({key: _front_text(item.get(key), 180) for key in required})
        if len(normalized_items) != 4:
            return {}
        result = {"schema_version": "1", "status": "pre_poc", "items": normalized_items}
        if str(value.get("migration_reason") or "").strip():
            result["migration_reason"] = _front_text(value.get("migration_reason"), 160)
        return result
    schema_version = str(value.get("schema_version") or "")
    if schema_version in LEGACY_POC_MEASUREMENT_DESIGN_SCHEMA_VERSIONS:
        upgraded = _legacy_measurement_design_to_v5(value, assessment)
        if not upgraded:
            return {}
        return normalize_poc_measurement_design(upgraded, assessment)
    if (schema_version != POC_MEASUREMENT_DESIGN_SCHEMA_VERSION
            or status not in {"decision_thresholds", "external_verified"} or len(items) != 3):
        return {}
    expected_pocs = quantitative_priority_pocs(assessment) if isinstance(assessment, dict) else []
    common_required = (
        "priority", "use_case_id", "theme", "business_outcome", "headline_metric", "kpi", "detail",
        "baseline_definition", "comparison_condition", "formula", "poc_gate", "target_rationale",
        "beneficiary", "metric_owner", "attribution_level", "causal_link", "product_kpi",
        "customer_outcome_kpi", "provider_business_kpi",
    )
    evidence_required = (
        "claim_id", "claim_status", "source_id", "source_url", "source_title", "source_locator",
        "source_metric", "evidence_excerpt",
    )
    normalized_items: list[dict] = []
    claim_ids: set[str] = set(); source_ids: set[str] = set(); use_case_ids: set[str] = set()
    for index, item in enumerate(items, 1):
        required = common_required + (evidence_required if status == "external_verified" else ())
        if not isinstance(item, dict) or not all(str(item.get(key) or "").strip() for key in required):
            return {}
        priority = _front_text(item.get("priority"), 8)
        use_case_id = _front_text(item.get("use_case_id"), 12)
        theme = _front_text(item.get("theme"), 72)
        if priority != f"P{index}" or not use_case_id or use_case_id in use_case_ids or not theme:
            return {}
        if expected_pocs:
            expected = expected_pocs[index - 1]
            if (priority != expected["priority"] or use_case_id != expected["use_case_id"]
                    or normalized_use_case_label(theme) != normalized_use_case_label(expected["theme"])):
                return {}
        owner = _normalize_measurement_owner(item.get("metric_owner"))
        beneficiary = str(item.get("beneficiary") or "")
        attribution = str(item.get("attribution_level") or "")
        display_kpi_layer = str(item.get("display_kpi_layer") or "product_kpi")
        if display_kpi_layer not in ALL_MEASUREMENT_DISPLAY_LAYERS:
            return {}
        product = _normalize_measurement_kpi_layer(item.get("product_kpi"), numeric_target=True)
        customer = _normalize_measurement_kpi_layer(
            item.get("customer_outcome_kpi"),
            numeric_target=display_kpi_layer == "customer_outcome_kpi",
        )
        provider = _normalize_measurement_kpi_layer(
            item.get("provider_business_kpi"),
            numeric_target=display_kpi_layer == "provider_business_kpi",
        )
        gate = item.get("poc_gate")
        if (not owner or beneficiary not in {"customer", "provider", "shared"}
                or not product or not customer or not provider or not isinstance(gate, dict)):
            return {}
        if (product["metric_owner"].get("scope") != "provider"
                or product.get("attribution_level") != "direct"
                or customer["metric_owner"].get("scope") not in {"customer", "shared"}
                or customer.get("attribution_level") != "contributory"
                or provider["metric_owner"].get("scope") != "provider"
                or provider.get("attribution_level") != "enabling"):
            return {}
        display_layer = {
            "product_kpi": product,
            "customer_outcome_kpi": customer,
            "provider_business_kpi": provider,
        }[display_kpi_layer]
        expected_owner_scopes = {
            "product_kpi": {"provider"},
            "customer_outcome_kpi": {"customer", "shared"},
            "provider_business_kpi": {"provider"},
        }[display_kpi_layer]
        expected_attribution = {
            "product_kpi": "direct",
            "customer_outcome_kpi": "contributory",
            "provider_business_kpi": "enabling",
        }[display_kpi_layer]
        expected_beneficiaries = {
            "product_kpi": {"customer", "provider", "shared"},
            "customer_outcome_kpi": {"customer", "shared"},
            "provider_business_kpi": {"provider"},
        }[display_kpi_layer]
        if (owner.get("scope") not in expected_owner_scopes
                or attribution != expected_attribution
                or beneficiary not in expected_beneficiaries):
            return {}
        if status == "decision_thresholds":
            for layer_name, layer in (
                    ("customer_outcome_kpi", customer),
                    ("provider_business_kpi", provider)):
                if layer_name != display_kpi_layer and _unverified_downstream_kpi_issues(layer, layer_name):
                    return {}
        normalized_gate = {
            "measurement_period": _customer_measurement_text(gate.get("measurement_period"), 42),
            "go_condition": _customer_measurement_text(gate.get("go_condition"), 92),
            "stop_condition": _customer_measurement_text(gate.get("stop_condition"), 92),
            "decision_owner": _customer_measurement_text(gate.get("decision_owner"), 58),
        }
        if not all(normalized_gate.values()):
            return {}
        if display_kpi_layer in BUSINESS_EFFECT_DISPLAY_LAYERS:
            normalized_gate.update({
                "business_go_condition": _customer_measurement_text(
                    gate.get("business_go_condition"), 92,
                ),
                "leading_kpi_condition": _customer_measurement_text(
                    gate.get("leading_kpi_condition"), 92,
                ),
            })
            if not normalized_gate["business_go_condition"] or not normalized_gate["leading_kpi_condition"]:
                return {}
        headline = _customer_measurement_text(item.get("headline_metric"), 28)
        kpi = _customer_measurement_text(item.get("kpi"), 48)
        headline_tokens = _quantitative_tokens(headline)
        product_tokens = _quantitative_tokens(product.get("target"))
        if (not _is_business_effect_metric(headline) or not headline_tokens
                or headline != display_layer["target"] or kpi != display_layer["kpi"]
                or not headline_tokens <= _quantitative_tokens(item.get("formula"))
                or not headline_tokens <= _quantitative_tokens(display_layer.get("formula"))
                or not headline_tokens <= _quantitative_tokens(normalized_gate["go_condition"])
                or not product_tokens
                or not product_tokens <= _quantitative_tokens(product.get("formula"))
                or not product_tokens <= _quantitative_tokens(normalized_gate["go_condition"])):
            return {}
        if display_kpi_layer in BUSINESS_EFFECT_DISPLAY_LAYERS and (
                not headline_tokens <= _quantitative_tokens(normalized_gate["business_go_condition"])
                or not product_tokens <= _quantitative_tokens(normalized_gate["leading_kpi_condition"])):
            return {}
        expected_basis = "verified_external_benchmark" if status == "external_verified" else str(item.get("basis_type") or "")
        if status == "external_verified" and expected_basis != "verified_external_benchmark":
            return {}
        if status == "decision_thresholds" and expected_basis not in {"decision_threshold", "legacy_migrated"}:
            return {}
        normalized_item = {
            "priority": priority, "use_case_id": use_case_id, "theme": theme,
            "business_outcome": _customer_measurement_text(item.get("business_outcome"), 48),
            "headline_metric": headline, "kpi": kpi,
            "detail": _customer_measurement_text(item.get("detail"), 130),
            "baseline_definition": _customer_measurement_text(item.get("baseline_definition"), 92),
            "comparison_condition": _customer_measurement_text(item.get("comparison_condition"), 88),
            "formula": _customer_measurement_text(item.get("formula"), 96),
            "poc_gate": normalized_gate,
            "target_rationale": _customer_measurement_text(item.get("target_rationale"), 110),
            "display_kpi_layer": display_kpi_layer,
            "beneficiary": beneficiary, "metric_owner": owner, "attribution_level": attribution,
            "causal_link": _customer_measurement_text(item.get("causal_link"), 110),
            "product_kpi": product, "customer_outcome_kpi": customer,
            "provider_business_kpi": provider, "basis_type": expected_basis,
        }
        if status == "decision_thresholds" and str(item.get("estimate_id") or "").strip():
            estimate_id = _front_text(item.get("estimate_id"), 20)
            low, high, unit = item.get("estimate_low"), item.get("estimate_high"), str(item.get("estimate_unit") or "")
            if (not re.fullmatch(r"E\d{2}", estimate_id) or not isinstance(low, (int, float))
                    or not isinstance(high, (int, float)) or not (0 < float(low) <= float(high) <= 100)
                    or unit != "%"):
                return {}
            normalized_item.update({"estimate_id": estimate_id, "estimate_low": float(low),
                                    "estimate_high": float(high), "estimate_unit": unit})
        if status == "external_verified":
            source_metric_tokens = _quantitative_tokens(item.get("source_metric"))
            excerpt_tokens = _quantitative_tokens(item.get("evidence_excerpt"))
            downstream_tokens = (
                _quantitative_tokens(customer.get("target"))
                | _quantitative_tokens(provider.get("target"))
            )
            claim_id = _front_text(item.get("claim_id"), 40); source_id = _front_text(item.get("source_id"), 20)
            if (not headline_tokens <= source_metric_tokens or not headline_tokens <= excerpt_tokens
                    or not downstream_tokens <= source_metric_tokens
                    or not downstream_tokens <= excerpt_tokens
                    or str(item.get("claim_status")) != "verified_external" or claim_id in claim_ids
                    or source_id in source_ids or not source_id.startswith("R")
                    or not is_safe_public_https_url(item.get("source_url"))):
                return {}
            claim_ids.add(claim_id); source_ids.add(source_id)
            normalized_item.update({
                "claim_id": claim_id, "claim_status": "verified_external", "source_id": source_id,
                "source_url": str(item.get("source_url")), "source_title": _front_text(item.get("source_title"), 75),
                "source_locator": _front_text(item.get("source_locator"), 120),
                "source_metric": _front_text(item.get("source_metric"), 160),
                "evidence_excerpt": _front_text(item.get("evidence_excerpt"), 260),
            })
        use_case_ids.add(use_case_id); normalized_items.append(normalized_item)
    management_targets = []
    for target in value.get("management_targets", []) if isinstance(value.get("management_targets"), list) else []:
        if not isinstance(target, dict) or str(target.get("source_id") or "") != "M1":
            continue
        label, target_text = _front_text(target.get("label"), 45), _front_text(target.get("target"), 45)
        if label and target_text:
            normalized_target = {
                "label": label, "target": target_text,
                "delta": _front_text(target.get("delta") or target.get("comparison"), 35),
                "source_id": "M1",
            }
            # External verification is re-run after normalization.  Preserve
            # the exact provenance fields rather than reducing them to labels.
            for field, limit in (
                ("claim_id", 40), ("claim_status", 30), ("source_url", 500),
                ("source_locator", 160), ("source_metric", 220),
                ("evidence_excerpt", 420),
            ):
                if str(target.get(field) or "").strip():
                    normalized_target[field] = _front_text(target.get(field), limit)
            management_targets.append(normalized_target)
    result = {
        "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION, "status": status,
        "lead": _customer_measurement_text(value.get("lead"), 220),
        "management_targets": management_targets[:3], "items": normalized_items,
    }
    if str(value.get("migration_status") or "") in {"legacy_migrated", "schema4_product_headline"}:
        result["migration_status"] = str(value.get("migration_status"))
    return result
def poc_measurement_design_from_quantitative_hypothesis(
        value: object, assessment: dict | None = None) -> dict:
    """生成した数値案を、実績主張ではなくPoC判定基準として凍結する。"""
    if not isinstance(value, dict) or value.get("evidence_mode") not in {"hypothesis", "llm_estimate"}:
        return {}
    benchmarks = value.get("benchmarks")
    scenarios = value.get("value_scenarios")
    if not isinstance(benchmarks, list) or not isinstance(scenarios, list) or len(benchmarks) != 3 or len(scenarios) != 3:
        return {}
    items = []
    for benchmark, scenario in zip(benchmarks, scenarios):
        if not isinstance(benchmark, dict) or not isinstance(scenario, dict):
            return {}
        if not all(str(benchmark.get(field) or "").strip() for field in (
            "priority", "use_case_id", "theme", "beneficiary", "metric_owner", "attribution_level",
            "causal_link", "product_kpi", "customer_outcome_kpi", "provider_business_kpi",
        )):
            return {}
        if any(
            str(scenario.get(field) or "") and str(scenario.get(field)) != str(benchmark.get(field))
            for field in ("priority", "use_case_id", "theme")
        ):
            return {}
        items.append({
            "priority": benchmark.get("priority"),
            "use_case_id": benchmark.get("use_case_id"),
            "theme": benchmark.get("theme"),
            "business_outcome": benchmark.get("use_case"),
            "headline_metric": benchmark.get("headline_metric"),
            "kpi": benchmark.get("kpi"),
            "detail": benchmark.get("detail"),
            "baseline_definition": benchmark.get("baseline_definition"),
            "comparison_condition": benchmark.get("comparison_condition"),
            "formula": scenario.get("formula"),
            "poc_gate": scenario.get("poc_gate"),
            "target_rationale": benchmark.get("source_title"),
            "display_kpi_layer": benchmark.get("display_kpi_layer"),
            "beneficiary": benchmark.get("beneficiary"),
            "metric_owner": benchmark.get("metric_owner"),
            "attribution_level": benchmark.get("attribution_level"),
            "causal_link": benchmark.get("causal_link"),
            "product_kpi": benchmark.get("product_kpi"),
            "customer_outcome_kpi": benchmark.get("customer_outcome_kpi"),
            "provider_business_kpi": benchmark.get("provider_business_kpi"),
            "basis_type": "decision_threshold",
            "estimate_id": benchmark.get("estimate_id"),
            "estimate_low": benchmark.get("estimate_low"),
            "estimate_high": benchmark.get("estimate_high"),
            "estimate_unit": benchmark.get("estimate_unit"),
        })
    return normalize_poc_measurement_design({
        "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
        "status": "decision_thresholds",
        "lead": (
            "対象サービスの業務データと優先PoCを踏まえ、AI実装で目指す売上・利益・生産性・損失抑制の定量目標を具体化します。"
            "同じ母集団と比較期間でPoC実測し、達成度を本番化と横展開の判断に用います。"
        ),
        "management_targets": value.get("management_targets", []),
        "items": items,
    }, assessment)
def poc_measurement_design_from_verified_evidence(
        value: object, research: object, assessment: dict | None = None, *,
        require_fetched: bool = False) -> dict:
    """原典照合済みの3件だけを、定量効果ページの再現可能な契約へ凍結する。"""
    if not isinstance(value, dict) or value.get("evidence_mode") != "external_verified":
        return {}
    if not isinstance(research, dict):
        return {}
    benchmarks = value.get("benchmarks")
    scenarios = value.get("value_scenarios")
    if (not isinstance(benchmarks, list) or not isinstance(scenarios, list)
            or len(benchmarks) != 3 or len(scenarios) != 3):
        return {}
    selected = quantitative_priority_pocs(assessment or {}) if isinstance(assessment, dict) else []
    # Schema v4 cannot present a customer business effect as if it were a
    # provider-controlled product KPI.  External evidence is displayable only
    # when it already binds to P1-P3 and supplies the three responsibility
    # layers; otherwise it remains calibration evidence for the LLM estimate.
    if len(selected) != 3:
        return {}
    sources = _research_sources_by_id(research)
    approved_ids = _research_approved_quantitative_source_ids(research)
    benchmark_source_ids = [str(item.get("source_id") or "") for item in benchmarks if isinstance(item, dict)]
    if len(benchmark_source_ids) != 3 or len(set(benchmark_source_ids)) != 3:
        return {}
    items: list[dict] = []
    used_scenario_indexes: set[int] = set()
    for selected_poc, benchmark in zip(selected, benchmarks):
        if not _quantitative_claim_is_verifiable(
                benchmark, sources, approved_ids, require_fetched=require_fetched):
            return {}
        if (str(benchmark.get("priority") or "") != selected_poc["priority"]
                or str(benchmark.get("use_case_id") or "") != selected_poc["use_case_id"]
                or normalized_use_case_label(benchmark.get("theme")) != normalized_use_case_label(selected_poc["theme"])
                or not all(isinstance(benchmark.get(field), dict) for field in (
                    "metric_owner", "product_kpi", "customer_outcome_kpi", "provider_business_kpi",
                ))):
            return {}
        source_id = str(benchmark.get("source_id") or "")
        metric = re.sub(r"\s+", "", str(benchmark.get("headline_metric") or ""))
        use_case = re.sub(r"\s+", "", str(benchmark.get("use_case") or ""))
        matches = [
            (index, scenario) for index, scenario in enumerate(scenarios)
            if isinstance(scenario, dict)
            and scenario.get("source_ids") == [source_id]
            and re.sub(r"\s+", "", str(scenario.get("benchmark") or "")) == metric
            and re.sub(r"\s+", "", str(scenario.get("use_case") or "")) == use_case
        ]
        if len(matches) != 1 or matches[0][0] in used_scenario_indexes:
            return {}
        scenario_index, scenario = matches[0]
        used_scenario_indexes.add(scenario_index)
        metric_tokens = _quantitative_tokens(benchmark.get("headline_metric"))
        if (not _is_business_effect_metric(benchmark.get("headline_metric"))
                or not metric_tokens <= _quantitative_tokens(scenario.get("formula"))
                or not metric_tokens <= _quantitative_tokens(scenario.get("poc_gate"))):
            return {}
        source = sources[source_id]
        items.append({
            "claim_id": benchmark.get("claim_id"),
            "claim_status": benchmark.get("claim_status"),
            "priority": selected_poc["priority"],
            "use_case_id": selected_poc["use_case_id"],
            "theme": selected_poc["theme"],
            "business_outcome": benchmark.get("use_case"),
            "headline_metric": benchmark.get("headline_metric"),
            "kpi": benchmark.get("kpi"),
            "detail": benchmark.get("detail"),
            "baseline_definition": "同一KPIの現状値（代表範囲・比較期間を固定）",
            "comparison_condition": "現行運用とAI支援を同一母集団・期間で比較",
            "formula": scenario.get("formula"),
            "poc_gate": scenario.get("poc_gate"),
            "target_rationale": f"{benchmark.get('source_title')}の公開効果を参照水準として採用",
            "display_kpi_layer": "product_kpi",
            "beneficiary": benchmark.get("beneficiary"),
            "metric_owner": benchmark.get("metric_owner"),
            "attribution_level": benchmark.get("attribution_level"),
            "causal_link": benchmark.get("causal_link"),
            "product_kpi": benchmark.get("product_kpi"),
            "customer_outcome_kpi": benchmark.get("customer_outcome_kpi"),
            "provider_business_kpi": benchmark.get("provider_business_kpi"),
            "basis_type": "verified_external_benchmark",
            "source_id": source_id,
            "source_url": benchmark.get("source_url"),
            "source_title": benchmark.get("source_title"),
            "source_locator": benchmark.get("source_locator"),
            "source_metric": benchmark.get("source_metric"),
            "evidence_excerpt": benchmark.get("evidence_excerpt"),
        })
    management_targets = extract_verified_management_targets_from_midterm_plan(
        _midterm_plan_source(research), max_targets=3,
    )
    return normalize_poc_measurement_design({
        "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
        "status": "external_verified",
        "lead": (
            "中期経営計画の到達目標と、対象業務に近い公開事例の効果水準を接続し、"
            "AI実装で狙う事業インパクトを具体化します。対象サービスでは同じKPI定義・母集団・比較期間で実測し、"
            "算定結果を本番化と横展開の判断に用います。"
        ),
        "management_targets": management_targets,
        "items": items,
    }, assessment)
def apply_quantitative_research_outcome(
        assessment: dict, research: dict, executive_evidence: object, *,
        llm_estimate: object | None = None,
        llm_model_id: str = "",
        llm_generation_log: list[dict] | None = None,
        missing_evidence: list[str] | None = None) -> dict:
    """Legacy compatibility helper for previously generated estimate contracts.
    The normal generation workflow no longer calls this function. New runs use
    ``resolve_quantitative_display_contract`` and
    ``materialize_quantitative_display_contract`` so unsupported numeric claims
    cannot bypass the single evidence gate. This helper remains only for tests
    and migration code that must understand an already frozen legacy payload.
    """
    evidence = executive_evidence if isinstance(executive_evidence, dict) else {}
    external_generation_audit = (
        copy.deepcopy(evidence.get("generation_audit"))
        if isinstance(evidence.get("generation_audit"), dict) else {}
    )
    if external_generation_audit:
        # 採用claimだけでなく、判定時に提示した全較正根拠と
        # プロンプト＋入力の安定hashをレビュー用JSONに残す。
        research["external_quantitative_generation_audit"] = external_generation_audit
    benchmarks = evidence.get("benchmarks") if isinstance(evidence.get("benchmarks"), list) else []
    benchmarks = [item for item in benchmarks if isinstance(item, dict)]
    used_source_ids = sorted({
        str(item.get("source_id")) for item in benchmarks
        if re.fullmatch(r"R\d+", str(item.get("source_id") or ""))
    })
    approved_claim_ids = [
        str(item.get("claim_id")) for item in benchmarks
        if str(item.get("claim_id") or "")
    ]
    research["quantitative_evidence"] = {
        "schema_version": QUANTITATIVE_EVIDENCE_SCHEMA_VERSION,
        "status": "sufficient" if len(benchmarks) == 3 and len(used_source_ids) == 3 else (
            "partial" if benchmarks else "insufficient"
        ),
        "approved_source_ids": used_source_ids,
        "reason": (
            "公開資料の個別監査と独立根拠監査を通過した数値効果だけを保持します。"
            if benchmarks else
            "対象業務へ直接対応する原典照合済みの定量効果を確認できませんでした。"
        ),
    }
    if evidence and benchmarks:
        assessment["executive_evidence"] = evidence
    measurement_design = {}
    if len(benchmarks) == 3 and len(used_source_ids) == 3 and len(set(approved_claim_ids)) == 3:
        measurement_design = poc_measurement_design_from_verified_evidence(evidence, research, assessment)
    if measurement_design:
        assessment["poc_measurement_design"] = measurement_design
        research["poc_measurement_design"] = measurement_design
        research["poc_measurement_design_audit"] = {
            "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
            "status": "external_verified",
            "basis_type": "verified_external_benchmark",
            "generation_source": "external_verified",
            "claim_ids": approved_claim_ids,
            "approved_source_ids": used_source_ids,
            "input_source_ids": external_generation_audit.get("input_source_ids", []),
            "calibration_source_ids": external_generation_audit.get("calibration_source_ids", []),
            "base_prompt_input_sha256": external_generation_audit.get("base_prompt_input_sha256", ""),
            "model_id": external_generation_audit.get("model_id", ""),
            "poc_bindings": [
                {field: str(item.get(field) or "") for field in ("priority", "use_case_id", "theme")}
                for item in measurement_design.get("items", []) if isinstance(item, dict)
            ],
            "reason": "3件の公開定量根拠を原文照合し、対応する対象サービス換算式・PoC判定基準へ固定しました。",
        }
        research["llm_quantitative_estimate"] = {
            "schema_version": "1",
            "status": "not_used",
            "display_mode": "external_verified",
            "model_id": "",
            "estimate_count": 0,
            "estimate_ids": [],
            "customer_actuals_assumed": False,
            "reason": "原典照合済みの外部定量効果が3件そろったため、LLM独自試算は表示に使用していません。",
        }
        research["quantitative_analysis"] = {
            "status": "verified_external_benchmarks",
            "final_display_mode": "external_verified",
            "required_claim_count": 3,
            "approved_claim_count": 3,
            "approved_claim_ids": approved_claim_ids,
            "approved_source_ids": used_source_ids,
            "llm_estimate_count": 0,
            "llm_model_id": "",
            "fallback_triggered": False,
            "customer_actuals_assumed": False,
            "display_contract": "verified_external_benchmark",
            "stopped_before_output": False,
            "unsupported_claims_omitted": False,
            "reason": "優先PoCまたは同じ業務KPIに対応する公開定量事例を3件確認しました。",
        }
        return {
            "status": "external_verified", "measurement_design": measurement_design,
            "approved_claim_ids": approved_claim_ids, "approved_source_ids": used_source_ids,
        }
    estimate = llm_estimate if isinstance(llm_estimate, dict) else {}
    estimate_design = poc_measurement_design_from_quantitative_hypothesis(estimate, assessment)
    external_status = (
        "contract_incomplete" if len(benchmarks) >= 3
        else "partial" if benchmarks else "insufficient"
    )
    if not estimate_design:
        # A failed PoC-level estimate must not block the independent P6 company
        # business case.  Preserve the failure in the audit log and render a
        # non-numeric measurement contract for the downstream PoC pages.  This
        # follows the evidence policy: unsupported PoC effects are omitted, not
        # repaired into plausible-looking claims.
        fallback_design = {
            "schema_version": "1", "status": "pre_poc",
            "items": poc_measurement_design_items(),
        }
        assessment["poc_measurement_design"] = fallback_design
        research["poc_measurement_design"] = copy.deepcopy(fallback_design)
        research["poc_measurement_design_audit"] = {
            "schema_version": "1",
            "status": "pre_poc",
            "basis_type": "measurement_contract",
            "generation_source": "unsupported_numeric_claims_omitted",
            "model_id": llm_model_id,
            "approved_source_ids": used_source_ids,
            "claim_ids": approved_claim_ids,
            "reason": (
                "外部原典3件または有効なPoC定量試算を確認できなかったため、"
                "個別PoCの未根拠数値を表示せず、現状値・比較条件・算定式・必要データを合意します。"
            ),
        }
        research["llm_quantitative_estimate"] = {
            "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
            "status": "failed_omitted",
            "display_mode": "measurement_design",
            "model_id": llm_model_id,
            "prompt_contract_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
            "estimate_count": 0,
            "estimate_ids": [],
            "customer_actuals_assumed": False,
            "calibration_source_ids": used_source_ids,
            "generation_attempts": copy.deepcopy(llm_generation_log or []),
            "reason": "PoC別定量試算が表示契約を満たさないため、数値主張を省略しました。",
        }
        research["quantitative_analysis"] = {
            "status": "measurement_design_only",
            "final_display_mode": "pre_poc",
            "external_evidence_status": external_status,
            "required_claim_count": 3,
            "approved_claim_count": len(benchmarks),
            "approved_claim_ids": approved_claim_ids,
            "approved_source_ids": used_source_ids,
            "llm_estimate_count": 0,
            "llm_estimate_ids": [],
            "llm_model_id": llm_model_id,
            "fallback_triggered": True,
            "customer_actuals_assumed": False,
            "display_contract": "pre_poc_measurement_design",
            "missing_evidence": [str(item) for item in (missing_evidence or []) if str(item).strip()],
            "stopped_before_output": False,
            "unsupported_claims_omitted": True,
            "reason": (
                "個別PoCの定量効果は未根拠のまま表示せず、別契約の会社事業KPIと、"
                "PoCで数値を確定する測定設計を使用します。"
            ),
        }
        return {
            "status": "measurement_design_only",
            "measurement_design": fallback_design,
            "approved_claim_ids": approved_claim_ids,
            "approved_source_ids": used_source_ids,
            "estimate_ids": [],
        }
    estimate_ids = [
        str(item.get("estimate_id") or "")
        for item in estimate_design.get("items", []) if isinstance(item, dict)
    ]
    displayed_metrics = [
        str(item.get("headline_metric") or "")
        for item in estimate_design.get("items", []) if isinstance(item, dict)
    ]
    display_kpi_layers = [
        str(item.get("display_kpi_layer") or "")
        for item in estimate_design.get("items", []) if isinstance(item, dict)
    ]
    product_leading_metrics = [
        str((item.get("product_kpi") or {}).get("target") or "")
        for item in estimate_design.get("items", [])
        if isinstance(item, dict) and isinstance(item.get("product_kpi"), dict)
    ]
    if len(estimate_ids) != 3 or len(set(estimate_ids)) != 3 or not all(estimate_ids):
        raise RuntimeError("LLM定量効果試算の識別子が3件で一意ではありません。")
    assessment["poc_measurement_design"] = estimate_design
    research["poc_measurement_design"] = estimate_design
    research["poc_measurement_design_audit"] = {
        "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
        "status": "decision_thresholds",
        "basis_type": "decision_threshold",
        "generation_source": "llm_estimate",
        "model_id": llm_model_id,
        "estimate_ids": estimate_ids,
        "displayed_metrics": displayed_metrics,
        "display_kpi_layers": display_kpi_layers,
        "product_leading_metrics": product_leading_metrics,
        "poc_bindings": [
            {field: str(item.get(field) or "") for field in ("priority", "use_case_id", "theme")}
            for item in estimate_design.get("items", []) if isinstance(item, dict)
        ],
        "external_verified_count": len(benchmarks),
        "calibration_source_ids": used_source_ids,
        "external_input_source_ids": external_generation_audit.get("input_source_ids", []),
        "external_calibration_source_ids": external_generation_audit.get("calibration_source_ids", []),
        "external_base_prompt_input_sha256": external_generation_audit.get("base_prompt_input_sha256", ""),
        "claim_ids": approved_claim_ids,
        "approved_source_ids": used_source_ids,
        "reason": (
            "外部原典の定量効果が3件に満たないため、対象サービス、優先PoC、利用可能データを基に"
            "LLMが3件のPoC効果目標レンジを設定しました。"
        ),
    }
    estimate_rows = []
    for item in estimate_design.get("items", []):
        if not isinstance(item, dict):
            continue
        estimate_rows.append({
            "estimate_id": str(item.get("estimate_id") or ""),
            "priority": str(item.get("priority") or ""),
            "use_case_id": str(item.get("use_case_id") or ""),
            "theme": str(item.get("theme") or ""),
            "business_outcome": str(item.get("business_outcome") or ""),
            "headline_metric": str(item.get("headline_metric") or ""),
            "estimate_low": item.get("estimate_low"),
            "estimate_high": item.get("estimate_high"),
            "estimate_unit": str(item.get("estimate_unit") or ""),
            "kpi": str(item.get("kpi") or ""),
            "baseline_definition": str(item.get("baseline_definition") or ""),
            "comparison_condition": str(item.get("comparison_condition") or ""),
            "formula": str(item.get("formula") or ""),
            "poc_gate": copy.deepcopy(item.get("poc_gate") or {}),
            "assumption": str(item.get("target_rationale") or ""),
            "display_kpi_layer": str(item.get("display_kpi_layer") or ""),
            "beneficiary": str(item.get("beneficiary") or ""),
            "metric_owner": copy.deepcopy(item.get("metric_owner") or {}),
            "attribution_level": str(item.get("attribution_level") or ""),
            "causal_link": str(item.get("causal_link") or ""),
            "product_kpi": copy.deepcopy(item.get("product_kpi") or {}),
            "customer_outcome_kpi": copy.deepcopy(item.get("customer_outcome_kpi") or {}),
            "provider_business_kpi": copy.deepcopy(item.get("provider_business_kpi") or {}),
        })
    research["llm_quantitative_estimate"] = {
        "schema_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
        "status": "generated",
        "display_mode": "llm_estimate",
        "model_id": llm_model_id,
        "prompt_contract_version": POC_MEASUREMENT_DESIGN_SCHEMA_VERSION,
        "estimate_count": len(estimate_rows),
        "estimate_ids": estimate_ids,
        "customer_actuals_assumed": False,
        "calibration_source_ids": used_source_ids,
        "generation_attempts": copy.deepcopy(llm_generation_log or []),
        "estimates": estimate_rows,
    }
    research["quantitative_analysis"] = {
        "status": "llm_estimated_targets",
        "final_display_mode": "llm_estimate",
        "external_evidence_status": external_status,
        "required_claim_count": 3,
        "approved_claim_count": len(benchmarks),
        "approved_claim_ids": approved_claim_ids,
        "approved_source_ids": used_source_ids,
        "llm_estimate_count": len(estimate_rows),
        "llm_estimate_ids": estimate_ids,
        "llm_model_id": llm_model_id,
        "fallback_triggered": True,
        "fallback_reason": "外部原典の3件表示契約（件数・原文照合・数値継承）が成立しないため",
        "customer_actuals_assumed": False,
        "display_contract": "poc_business_effect_thresholds",
        "missing_evidence": [str(item) for item in (missing_evidence or []) if str(item).strip()],
        "stopped_before_output": False,
        "unsupported_claims_omitted": False,
        "reason": (
            "確認済みの外部根拠は監査ログに保持し、顧客実績を事実として補わず、"
            "LLMが生成した3件の事業効果レンジと製品先行KPIをPoC判定目標として表示します。"
        ),
    }
    return {
        "status": "llm_estimate", "measurement_design": estimate_design,
        "approved_claim_ids": approved_claim_ids, "approved_source_ids": used_source_ids,
        "estimate_ids": estimate_ids,
    }
def poc_measurement_design_for(assessment: dict) -> dict:
    """レビューJSONに凍結された定量設計を優先し、旧JSONだけ安全な設計へ戻す。"""
    raw = assessment.get("poc_measurement_design")
    normalized = normalize_poc_measurement_design(raw, assessment)
    if normalized:
        return normalized
    if isinstance(raw, dict) and str(raw.get("status") or "") in {
            "decision_thresholds", "external_verified"}:
        return {
            "schema_version": str(raw.get("schema_version") or ""),
            "status": "invalid_numeric_contract",
            "reason": "定量設計が優先PoC・責任境界・事業効果表示契約と一致しません。",
        }
    return {"schema_version": "1", "status": "pre_poc", "items": poc_measurement_design_items()}
def poc_measurement_design_items() -> list[dict[str, str]]:
    """外部根拠がない段階で、数値を創作せずに合意できる定量KPI契約を返す。"""
    return [
        {
            "dimension": "品質",
            "kpi": "合格率・正答率・手戻り率",
            "baseline": "現行運用で同一基準により判定した結果",
            "comparison": "同じ対象範囲・期間・判定者でAI利用時と比較",
            "formula": "品質改善率 ＝（AI利用時の合格率 − 現行合格率）÷ 現行合格率",
            "data": "評価対象、正解／判定結果、修正履歴、判定者",
        },
        {
            "dimension": "時間",
            "kpi": "処理時間・判断時間・待ち時間",
            "baseline": "現行運用の開始から完了までの実測時間",
            "comparison": "同一業務・件数・難易度でAI利用有無を比較",
            "formula": "時間削減率 ＝（現行平均時間 − AI利用平均時間）÷ 現行平均時間",
            "data": "開始／完了時刻、対象件数、担当者、例外処理時間",
        },
        {
            "dimension": "リスク",
            "kpi": "見逃し率・誤検知率・損失回避",
            "baseline": "現行運用で発生した見逃し・誤判定・損失",
            "comparison": "同じ定義と観測期間でAI支援時の結果を比較",
            "formula": "損失回避額 ＝ 対象件数 × 一件当たり実績損失 × 回避率",
            "data": "事象定義、発生件数、確定結果、実績損失、対応履歴",
        },
        {
            "dimension": "利用・定着",
            "kpi": "採用率・継続利用率・手動復帰率",
            "baseline": "対象利用者と現行機能の利用・手動処理状況",
            "comparison": "同じ利用者群で試行期間中の利用行動を比較",
            "formula": "継続利用率 ＝ 継続利用者数 ÷ 対象利用者数",
            "data": "対象利用者、利用ログ、採用／修正／復帰、継続期間",
        },
    ]
QUANTITATIVE_EVIDENCE_SCHEMA_VERSION = "1"
QUANTITATIVE_DISPLAY_CONTRACT_SCHEMA_VERSION = "1"
class QuantitativeEvidenceStatus(str, Enum):
    """原典照合の充足状況。
    ``contract_incomplete`` は候補数だけはそろっていても、原文一致、出典、
    P1-P3との結合、または責任境界のいずれかが表示契約を満たさない状態を表す。
    """
    VERIFIED = "verified"
    PARTIAL = "partial"
    INSUFFICIENT = "insufficient"
    CONTRACT_INCOMPLETE = "contract_incomplete"
class QuantitativeDisplayMode(str, Enum):
    """PPTXへ許可する定量表示モード。"""
    EXTERNAL_VERIFIED = "external_verified"
    MEASUREMENT_DESIGN_ONLY = "measurement_design_only"
@dataclass(frozen=True)
class QuantitativeDisplayContract:
    """定量根拠、測定設計、表示可否を一体で返す不変の判定結果。
    可変な辞書を直接保持すると、描画や監査ログ作成時の副作用で判定結果が
    書き換わり得るため、測定設計は正規化済みJSON文字列として保持する。
    ``measurement_design`` はアクセスのたびに新しい辞書を返す。
    """
    schema_version: str
    evidence_status: QuantitativeEvidenceStatus
    display_mode: QuantitativeDisplayMode
    numeric_claims_allowed: bool
    _measurement_design_json: str
    approved_claim_ids: tuple[str, ...]
    approved_source_ids: tuple[str, ...]
    missing_evidence: tuple[str, ...]
    reason: str
    @property
    def measurement_design(self) -> dict:
        return json.loads(self._measurement_design_json)
    def as_dict(self) -> dict:
        """レビューJSONと監査ログへ保存できる標準形を返す。"""
        return {
            "schema_version": self.schema_version,
            "evidence_status": self.evidence_status.value,
            "display_mode": self.display_mode.value,
            "numeric_claims_allowed": self.numeric_claims_allowed,
            "measurement_design": self.measurement_design,
            "approved_claim_ids": list(self.approved_claim_ids),
            "approved_source_ids": list(self.approved_source_ids),
            "missing_evidence": list(self.missing_evidence),
            "reason": self.reason,
        }
def _frozen_measurement_design(value: dict) -> str:
    """安定したJSONへ変換し、呼出元の辞書から判定結果を切り離す。"""
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
def _pre_poc_measurement_design() -> dict:
    """外部定量根拠が不足する場合の、数値主張を含まない測定契約。"""
    return {
        "schema_version": "1",
        "status": "pre_poc",
        "items": poc_measurement_design_items(),
    }
def resolve_quantitative_display_contract(
        assessment: dict, research: dict, executive_evidence: object, *,
        missing_evidence: list[str] | None = None) -> QuantitativeDisplayContract:
    """外部根拠・測定設計・表示可否を一度だけ判定する。
    数値表示を許可するのは、独立した3件の公開根拠が既存の原文照合契約と
    P1-P3結合契約をすべて満たす場合だけである。候補が不足、部分的、または
    契約不整合の場合は、数値を補完せずに測定設計へ切り替える。この関数は
    ``assessment`` と ``research`` を変更しない。
    """
    evidence = executive_evidence if isinstance(executive_evidence, dict) else {}
    benchmarks = evidence.get("benchmarks") if isinstance(evidence.get("benchmarks"), list) else []
    benchmarks = [copy.deepcopy(item) for item in benchmarks if isinstance(item, dict)]
    sources = _research_sources_by_id(research)
    approved_source_pool = _research_approved_quantitative_source_ids(research)
    verified_rows = [
        item for item in benchmarks
        if _quantitative_claim_is_verifiable(item, sources, approved_source_pool)
    ]
    approved_source_ids = tuple(sorted({
        str(item.get("source_id") or "") for item in verified_rows
        if re.fullmatch(r"R\d+", str(item.get("source_id") or ""))
    }))
    approved_claim_ids = tuple(
        str(item.get("claim_id") or "") for item in verified_rows
        if str(item.get("claim_id") or "").strip()
    )
    verified_design: dict = {}
    if (len(verified_rows) == 3 and len(approved_source_ids) == 3
            and len(approved_claim_ids) == 3 and len(set(approved_claim_ids)) == 3):
        verified_design = poc_measurement_design_from_verified_evidence(
            evidence, research, assessment, require_fetched=True,
        )
    if verified_design:
        return QuantitativeDisplayContract(
            schema_version=QUANTITATIVE_DISPLAY_CONTRACT_SCHEMA_VERSION,
            evidence_status=QuantitativeEvidenceStatus.VERIFIED,
            display_mode=QuantitativeDisplayMode.EXTERNAL_VERIFIED,
            numeric_claims_allowed=True,
            _measurement_design_json=_frozen_measurement_design(verified_design),
            approved_claim_ids=approved_claim_ids,
            approved_source_ids=approved_source_ids,
            missing_evidence=(),
            reason=(
                "独立した3件の公開定量根拠について、出典、原文数値、P1-P3結合、"
                "責任境界を確認したため数値表示を許可します。"
            ),
        )
    missing = tuple(dict.fromkeys(
        str(item).strip() for item in (missing_evidence or []) if str(item).strip()
    ))
    if not verified_rows:
        status = QuantitativeEvidenceStatus.INSUFFICIENT
    elif len(verified_rows) < 3:
        status = QuantitativeEvidenceStatus.PARTIAL
    else:
        status = QuantitativeEvidenceStatus.CONTRACT_INCOMPLETE
    return QuantitativeDisplayContract(
        schema_version=QUANTITATIVE_DISPLAY_CONTRACT_SCHEMA_VERSION,
        evidence_status=status,
        display_mode=QuantitativeDisplayMode.MEASUREMENT_DESIGN_ONLY,
        numeric_claims_allowed=False,
        _measurement_design_json=_frozen_measurement_design(_pre_poc_measurement_design()),
        approved_claim_ids=approved_claim_ids,
        approved_source_ids=approved_source_ids,
        missing_evidence=missing,
        reason=(
            "原典照合済みの独立した定量根拠が3件そろっていないため、"
            "未根拠の割合・金額を表示せず、現状値・比較条件・算定式・必要データを"
            "PoCで確定する測定設計を使用します。"
        ),
    )
def materialize_quantitative_display_contract(
        assessment: dict, research: dict,
        contract: QuantitativeDisplayContract) -> dict:
    """単一の判定結果から、表示用データと監査ログを同時に固定する。
    この関数だけが定量表示の主要キーを書き込むことで、assessment側の表示値と
    research側の監査結果が別々の条件分岐で食い違うことを防ぐ。
    """
    design = contract.measurement_design
    assessment["poc_measurement_design"] = copy.deepcopy(design)
    research["poc_measurement_design"] = copy.deepcopy(design)
    research["quantitative_display_contract"] = contract.as_dict()
    research["quantitative_evidence"] = {
        "schema_version": QUANTITATIVE_EVIDENCE_SCHEMA_VERSION,
        "status": (
            "sufficient" if contract.evidence_status is QuantitativeEvidenceStatus.VERIFIED
            else "partial" if contract.evidence_status in {
                QuantitativeEvidenceStatus.PARTIAL,
                QuantitativeEvidenceStatus.CONTRACT_INCOMPLETE,
            }
            else "insufficient"
        ),
        "approved_source_ids": list(contract.approved_source_ids),
        "reason": contract.reason,
    }
    generation_source = (
        "external_verified" if contract.numeric_claims_allowed
        else "unsupported_numeric_claims_omitted"
    )
    research["poc_measurement_design_audit"] = {
        "schema_version": str(design.get("schema_version") or "1"),
        "status": str(design.get("status") or "pre_poc"),
        "basis_type": (
            "verified_external_benchmark" if contract.numeric_claims_allowed
            else "measurement_contract"
        ),
        "generation_source": generation_source,
        "claim_ids": list(contract.approved_claim_ids),
        "approved_source_ids": list(contract.approved_source_ids),
        "poc_bindings": [
            {field: str(item.get(field) or "") for field in (
                "priority", "use_case_id", "theme",
            )}
            for item in design.get("items", []) if isinstance(item, dict)
        ] if contract.numeric_claims_allowed else [],
        "reason": contract.reason,
    }
    research["quantitative_analysis"] = {
        "status": (
            "verified_external_benchmarks" if contract.numeric_claims_allowed
            else "measurement_design_only"
        ),
        "final_display_mode": contract.display_mode.value,
        "external_evidence_status": contract.evidence_status.value,
        "required_claim_count": 3,
        "approved_claim_count": len(contract.approved_claim_ids),
        "approved_claim_ids": list(contract.approved_claim_ids),
        "approved_source_ids": list(contract.approved_source_ids),
        "customer_actuals_assumed": False,
        "display_contract": (
            "verified_external_benchmark" if contract.numeric_claims_allowed
            else "pre_poc_measurement_design"
        ),
        "missing_evidence": list(contract.missing_evidence),
        "stopped_before_output": False,
        "unsupported_claims_omitted": not contract.numeric_claims_allowed,
        "reason": contract.reason,
    }
    return {
        "status": contract.display_mode.value,
        "measurement_design": copy.deepcopy(design),
        "approved_claim_ids": list(contract.approved_claim_ids),
        "approved_source_ids": list(contract.approved_source_ids),
    }
def materialize_safe_reproducibility_v2_quantitative_contract(
        assessment: dict, research: dict, *, migrate_legacy: bool = False) -> dict:
    """Freeze external claims, PoC design and company planning estimates safely.
    External-result percentages still require three fetched source bodies.
    Separately, an audited ``ai_product_business_impact`` may contain three LLM
    planning ranges for the assessed company's own sales/profit/recurring-value
    decision.  It is retained only when model, assumptions, reason, source count
    and responsibility boundary are frozen in the research log.
    """
    if not isinstance(assessment, dict) or not isinstance(research, dict):
        raise TypeError("assessmentとresearchはdictで指定してください。")
    legacy_fields: list[str] = []
    impact_model = normalize_ai_product_business_impact_candidate(
        assessment.get("ai_product_business_impact"), assessment,
    )
    impact_audit = research.get("ai_product_business_impact_audit")
    impact_preserved = bool(
        impact_model
        and isinstance(impact_audit, dict)
        and _ai_product_business_impact_is_consistent(
            impact_model, research, assessment,
        )
        # 会社KPIを前提にするAI導入効果の計画試算は、そのKPI自体が現行の
        # 中計原典照合を通る場合だけ保持する。KPIが誤抽出と判明した後も
        # 派生した効果試算だけ残ることを防ぐ。
        and all(
            _management_target_is_verifiable(target, research, require_fetched=True)
            for target in impact_model.get("management_targets", [])
            if isinstance(target, dict)
        )
    )
    if impact_preserved and isinstance(impact_audit, dict):
        approved_ids = _ai_product_business_impact_external_source_ids(research)
        impact_audit.update({
            "evidence_mode": "llm_estimate",
            "estimate_classification": "planning_estimate",
            "display_label": "AI導入効果目標",
            "calculation_assumptions": ai_product_business_impact_calculation_assumptions(
                impact_model,
            ),
            "generation_reason": str(
                impact_audit.get("generation_reason")
                or impact_audit.get("reason")
                or "評価対象企業・製品全体のAI導入効果を事業計画として試算するため。"
            ),
            "external_evidence_count": len(approved_ids),
            "external_source_ids": approved_ids,
            "management_target_claim_ids": [
                str(target.get("claim_id") or "")
                for target in impact_model.get("management_targets", [])
                if isinstance(target, dict)
            ],
            "actual_result_claimed": False,
        })
        assessment["ai_product_business_impact"] = copy.deepcopy(impact_model)
        research["ai_product_business_impact"] = copy.deepcopy(impact_model)
        impact_preserved = _ai_product_business_impact_is_consistent(
            impact_model, research, assessment,
        )
    if not impact_preserved:
        if assessment.pop("ai_product_business_impact", None) is not None:
            legacy_fields.append("assessment.ai_product_business_impact")
        for field in ("ai_product_business_impact", "ai_product_business_impact_audit"):
            if research.pop(field, None) is not None:
                legacy_fields.append(f"research.{field}")
    if research.pop("llm_quantitative_estimate", None) is not None:
        legacy_fields.append("research.llm_quantitative_estimate")
    prior_analysis = research.get("quantitative_analysis")
    if isinstance(prior_analysis, dict) and (
            prior_analysis.get("final_display_mode") == "llm_estimate"
            or prior_analysis.get("status") == "llm_estimated_targets"
            or prior_analysis.get("unsupported_claims_omitted") is False):
        legacy_fields.append("research.quantitative_analysis")
    prior_audit = research.get("poc_measurement_design_audit")
    if isinstance(prior_audit, dict) and prior_audit.get("generation_source") in {
            "llm_estimate", "llm_decision_threshold"}:
        legacy_fields.append("research.poc_measurement_design_audit")
    prior_design = assessment.get("poc_measurement_design")
    if isinstance(prior_design, dict) and prior_design.get("status") == "decision_thresholds":
        legacy_fields.append("assessment.poc_measurement_design")
    sanitize_quantitative_evidence(
        assessment, research, migrate_legacy=migrate_legacy, require_fetched=True,
    )
    # 中計ページはexecutive_evidenceとは別にmanagement_targetsを保持する。
    # ここも原文取得済みの同じ契約で絞らないと、P4だけに旧数値が残り得る。
    for holder_name, holder in (("assessment", assessment), ("research", research)):
        analysis = holder.get("midterm_plan_analysis")
        if not isinstance(analysis, dict):
            continue
        raw_targets = analysis.get("management_targets")
        targets = raw_targets if isinstance(raw_targets, list) else []
        kept_targets = [
            target for target in targets
            if _management_target_is_verifiable(
                target, research, require_fetched=True,
            )
        ]
        if len(kept_targets) != len(targets):
            legacy_fields.append(f"{holder_name}.midterm_plan_analysis.management_targets")
        if isinstance(raw_targets, list):
            analysis["management_targets"] = kept_targets
        # 中計の数値は構造化management_targetsカードだけで表示する。
        # LLM要約へ同じ数字が流れ込むと、取得本文の数値を別の効果へ
        # 意味変更しても追跡不能になるため、自由文は非数値へ正規化する。
        if _is_business_effect_metric(analysis.get("plan_summary")):
            analysis["plan_summary"] = (
                "公開中期経営計画の重点方針、成長領域、実行条件を整理し、"
                "AI実装との接続点を確認します。"
            )
            legacy_fields.append(f"{holder_name}.midterm_plan_analysis.plan_summary")
        alignment_rows = analysis.get("ai_alignment")
        if isinstance(alignment_rows, list):
            replacements = {
                "plan_priority": "公開中期経営計画の重点方針",
                "ai_role": "対象業務の判断支援・標準化・説明力向上",
                "why_now": "重点施策の実行力を高め、利用定着と横展開につなげるため",
                "related_use_case": "優先AIユースケース",
            }
            for index, row in enumerate(alignment_rows):
                if not isinstance(row, dict):
                    continue
                for field, replacement in replacements.items():
                    if _is_business_effect_metric(row.get(field)):
                        row[field] = replacement
                        legacy_fields.append(
                            f"{holder_name}.midterm_plan_analysis.ai_alignment[{index}].{field}"
                        )
        if _is_business_effect_metric(analysis.get("caveat")):
            analysis["caveat"] = (
                "AI施策の効果と優先度は、対象データ、比較条件、責任者を合意し、"
                "PoCの実測結果で判断します。"
            )
            legacy_fields.append(f"{holder_name}.midterm_plan_analysis.caveat")
    prior_missing = (
        prior_analysis.get("missing_evidence")
        if isinstance(prior_analysis, dict)
        and isinstance(prior_analysis.get("missing_evidence"), list)
        else []
    )
    missing_evidence = [str(item) for item in prior_missing if str(item).strip()]
    if legacy_fields:
        missing_evidence.append(
            "旧JSONのLLM独自試算または未根拠の定量表示契約は、原典取得済みの根拠へ追跡できないため省略しました。"
        )
    contract = resolve_quantitative_display_contract(
        assessment,
        research,
        assessment.get("executive_evidence"),
        missing_evidence=missing_evidence,
    )
    evidence = assessment.get("executive_evidence")
    if isinstance(evidence, dict):
        if _is_business_effect_metric(evidence.get("plan_evidence_summary")):
            evidence["plan_evidence_summary"] = (
                "公開資料に記載された事業方針と、対象サービスでAIが支援する業務判断を接続して検討します。"
            )
            legacy_fields.append("assessment.executive_evidence.plan_evidence_summary")
        if _is_business_effect_metric(evidence.get("ai_necessity_analysis")):
            evidence["ai_necessity_analysis"] = (
                "対象サービスの業務・データ・利用者を踏まえ、判断支援、例外検知、説明支援を"
                "人の承認と組み合わせて検証し、実測結果から標準機能化の可否を判断します。"
            )
            legacy_fields.append("assessment.executive_evidence.ai_necessity_analysis")
    if not contract.numeric_claims_allowed:
        if isinstance(evidence, dict):
            if evidence.get("benchmarks") or evidence.get("value_scenarios"):
                legacy_fields.append("assessment.executive_evidence.partial_quantitative_claims")
            evidence["evidence_mode"] = "pre_poc"
            evidence["benchmarks"] = []
            evidence["value_scenarios"] = []
            # 部分根拠や旧LLM試算の値がP5等の自由文へ回り込まないよう、
            # 非数値の説明へ戻す。取得済み中計の構造化targetは別契約で保持する。
    outcome = materialize_quantitative_display_contract(assessment, research, contract)
    if impact_preserved:
        impact_audit = research.get("ai_product_business_impact_audit")
        if isinstance(impact_audit, dict):
            approved_ids = sorted(str(item) for item in outcome["approved_source_ids"])
            impact_audit["external_source_ids"] = approved_ids
            impact_audit["external_evidence_count"] = len(approved_ids)
        if not _ai_product_business_impact_is_consistent(
                assessment.get("ai_product_business_impact"), research, assessment):
            raise RuntimeError(
                "AI導入効果目標の計画試算と外部根拠件数を再現性契約へ固定できません。"
            )
    if legacy_fields:
        research["legacy_quantitative_migration_audit"] = {
            "schema_version": "1",
            "status": "unsupported_numeric_claims_omitted",
            "omitted_fields": sorted(set(legacy_fields)),
            "reason": (
                "外部実績claimとして追跡できない数値は省略しました。評価対象企業・"
                "製品全体のLLM計画試算は、独立した監査契約を満たす場合だけ保持します。"
            ),
        }
    return outcome
def quantitative_display_contract_issues(
        value: object, assessment: dict, research: dict) -> list[str]:
    """凍結JSONの単一定量表示契約と表示データの不一致を検出する。"""
    if not isinstance(value, dict):
        return ["research.quantitative_display_contractはオブジェクトである必要があります。"]
    issues: list[str] = []
    if str(value.get("schema_version") or "") != QUANTITATIVE_DISPLAY_CONTRACT_SCHEMA_VERSION:
        issues.append("quantitative_display_contract.schema_versionが不正です。")
    mode = str(value.get("display_mode") or "")
    evidence_status = str(value.get("evidence_status") or "")
    numeric_allowed = value.get("numeric_claims_allowed")
    design = value.get("measurement_design")
    if mode not in {item.value for item in QuantitativeDisplayMode}:
        issues.append("quantitative_display_contract.display_modeが不正です。")
    if evidence_status not in {item.value for item in QuantitativeEvidenceStatus}:
        issues.append("quantitative_display_contract.evidence_statusが不正です。")
    if not isinstance(numeric_allowed, bool):
        issues.append("quantitative_display_contract.numeric_claims_allowedは真偽値です。")
    if not isinstance(design, dict):
        issues.append("quantitative_display_contract.measurement_designがありません。")
        return issues
    if design != assessment.get("poc_measurement_design"):
        issues.append("quantitative_display_contractとassessment.poc_measurement_designが一致しません。")
    if numeric_allowed is True:
        if mode != QuantitativeDisplayMode.EXTERNAL_VERIFIED.value:
            issues.append("数値表示許可時のdisplay_modeはexternal_verifiedです。")
        if evidence_status != QuantitativeEvidenceStatus.VERIFIED.value:
            issues.append("数値表示許可時のevidence_statusはverifiedです。")
        if not _poc_measurement_design_is_verifiable(design, research):
            issues.append("数値表示契約は原典照合済みの測定設計と一致しません。")
    elif numeric_allowed is False:
        if mode != QuantitativeDisplayMode.MEASUREMENT_DESIGN_ONLY.value:
            issues.append("数値表示不可時のdisplay_modeはmeasurement_design_onlyです。")
        if str(design.get("status") or "") != "pre_poc":
            issues.append("数値表示不可時はpre_poc測定設計を使用してください。")
        if _quantitative_tokens(json.dumps(design.get("items") or [], ensure_ascii=False)):
            issues.append("measurement_design_onlyに固定の効果数値を含められません。")
    return issues
def _research_sources_by_id(research: object) -> dict[str, dict]:
    if not isinstance(research, dict):
        return {}
    sources = research.get("industry_sources")
    if not isinstance(sources, list):
        return {}
    result: dict[str, dict] = {}
    for index, source in enumerate(sources, 1):
        if isinstance(source, dict):
            source_id = str(source.get("id") or f"R{index}")
            result[source_id] = source
    return result
def _research_approved_quantitative_source_ids(research: object) -> set[str]:
    if not isinstance(research, dict):
        return set()
    quantitative = research.get("quantitative_evidence")
    if isinstance(quantitative, dict):
        return {
            str(source_id) for source_id in quantitative.get("approved_source_ids", [])
            if isinstance(source_id, str) and source_id.startswith("R")
        }
    audits = research.get("research_audit")
    if isinstance(audits, list) and audits:
        audit = audits[-1].get("audit") if isinstance(audits[-1], dict) else {}
        if isinstance(audit, dict) and audit.get("sufficient") is True:
            return {
                str(source_id) for source_id in audit.get("usable_source_ids", [])
                if isinstance(source_id, str) and source_id.startswith("R")
            }
    return set()
def _normalized_evidence_text(value: object) -> str:
    return re.sub(r"\s+", "", str(value or ""))
QUANTITATIVE_LITERAL_PATTERN = re.compile(
    r"(?:[$¥]\s*\d[\d,.]*\s*[KMBkmb]?|"
    r"\d[\d,.]*(?:\s*[〜～~\-–—→]\s*\d[\d,.]*)?\s*"
    r"(?:%|％|パーセント|割|倍|[xX]|万|億|円|ドル|時間|分|秒|週|日|月|年|"
    r"BPS|bps|pp|ポイント|pt|TPS|tps|req/s|件/時))"
)
def _quantitative_tokens(value: object) -> set[str]:
    return {token.replace(" ", "") for token in QUANTITATIVE_LITERAL_PATTERN.findall(str(value or ""))}
def _is_business_effect_metric(value: object) -> bool:
    """期間・件数だけを、改善効果の見出し数値として誤採用しない。"""
    text = re.sub(r"\s+", "", str(value or ""))
    if not _quantitative_tokens(text):
        return False
    if re.search(
            r"(?:%|％|パーセント|割|円|万円|億円|ドル|[$¥]|倍|[xX]|"
            r"BPS|bps|pp|ポイント|pt|TPS|tps|req/s|件/時)", text):
        return True
    # 「60分→15分」「3時間から1時間」のように、改善前後が同時に示される
    # 時間効果は採用する。一方、「12週間」のような検証期間だけは除外する。
    if re.search(r"\d[\d,.]*(?:時間|分|秒).*(?:→|から|より).*(?:\d[\d,.]*(?:時間|分|秒))", text):
        return True
    return bool(
        re.search(r"(?:削減|短縮|向上|改善|増加|減少|抑制|回避)", text)
        and re.search(r"\d[\d,.]*(?:時間|分|秒|件|人|社|拠点|店舗|台)", text)
    )
def _midterm_plan_source(research: object) -> dict:
    if not isinstance(research, dict):
        return {}
    plan = research.get("midterm_plan")
    if not isinstance(plan, dict) or plan.get("status") != "found":
        return {}
    if not is_safe_public_https_url(plan.get("url")):
        return {}
    return plan
def extract_verified_management_targets_from_midterm_plan(plan: object,
                                                          *, max_targets: int = 4) -> list[dict]:
    """公開中計の原文に明記された経営目標だけを、表示可能なclaimへ変換する。
    業界・顧客ごとの固定値は持たず、売上・利益・継続収益・人員など一般的な
    経営KPI語と、その直後にある数値表現を原文から抽出する。AIによる効果値では
    なく、経営計画が掲げる到達点として保持する。
    """
    if not isinstance(plan, dict) or plan.get("status") != "found" or not is_safe_public_https_url(plan.get("url")):
        return []
    excerpt = re.sub(r"\s+", " ", str(plan.get("excerpt") or "")).strip()
    if not excerpt:
        return []
    # 年次報告書のPDFには、中計の重要KPI表と個別事業のロードマップが同居する。
    # PDF抽出後の行順だけで全ページを走査すると、個別事業の数値を全社中計目標と
    # 誤認するため、まず中計の重要KPI節だけを対象にする。対象範囲を特定できない
    # 資料は、数値を表示せず方針だけを扱う（数値を推測しない）。
    section_match = re.search(
        r"(?:重要\s*経営\s*指標|業績\s*目標|経営\s*目標|財務\s*目標|数値\s*目標)",
        excerpt,
        re.IGNORECASE,
    )
    if section_match is None:
        return []
    target_scope = excerpt[section_match.start():section_match.start() + 3600]
    target_pattern = re.compile(
        r"(?P<label>(?:[^\s。！？]{1,24}\s+){0,2}[^\s。！？]{1,24})\s+"
        r"(?P<number>\d[\d,]*(?:\.\d+)?)\s*(?P<unit>億円|百万円|万円|円|%|％|人|件|社|拠点|店舗|台)"
    )
    kpi_tokens = ("売上", "利益", "収益", "MRR", "ARR", "EBITDA", "人員", "従業員", "顧客", "契約", "稼働", "シェア", "生産性", "EPS", "ROE", "ROIC")
    noise_tokens = {"表示", "計画", "目標値", "目標", "以上", "原典", "単体での表示"}
    targets: list[dict] = []
    seen: set[tuple[str, str]] = set()
    plan_years = [int(year) for year in re.findall(r"20(\d{2})", str(plan.get("title") or ""))]
    # 例: 中期経営計画(2024-2026)のKPIは、計画最終年度の2027年3月期で示される。
    allowed_period_years = {2000 + max(plan_years), 2001 + max(plan_years)} if plan_years else set()
    for match in target_pattern.finditer(target_scope):
        raw_label = re.sub(r"20\d{2}年(?:\d{1,2}月)?(?:期)?", "", match.group("label")).strip()
        if not any(token.lower() in raw_label.lower() for token in kpi_tokens):
            continue
        parts = [part for part in raw_label.split() if part not in noise_tokens]
        # KPI語を含む語から後ろを採用し、直前が「クラウドサービス」のような
        # 短い修飾語なら併せて残す。
        kpi_index = next((index for index, part in enumerate(parts)
                          if any(token.lower() in part.lower() for token in kpi_tokens)), len(parts) - 1)
        start = max(0, kpi_index - 1) if kpi_index == len(parts) - 1 and kpi_index > 0 and len(parts[kpi_index - 1]) <= 16 else kpi_index
        label = " ".join(parts[start:]).strip(" ：:|｜")
        if not label or len(label) > 28:
            continue
        # "売上高 EPS 16%" のように表の別列が連結したPDF抽出は、値と指標の
        # 対応を再構成できない。複数KPIや数値を含むラベルは不採用にする。
        kpi_count = sum(1 for token in kpi_tokens if token.lower() in label.lower())
        if kpi_count != 1 or _quantitative_tokens(label):
            continue
        value = f"{match.group('number')}{match.group('unit').replace('％', '%')}"
        key = (label.replace(" ", ""), value)
        if key in seen:
            continue
        seen.add(key)
        # PDF抽出では「2027年 売上高 6,200億円」の期間もlabel側に連結し得る。
        # 数値の直前まで含めて期間を拾い、指標名からは上で除去している。
        before = target_scope[max(0, match.start() - 220):match.end()]
        periods = re.findall(r"20\d{2}年(?:\d{1,2}月)?(?:期)?", before)
        period = periods[-1] if periods else ""
        period_years = [int(year[:4]) for year in re.findall(r"20\d{2}", period)]
        if allowed_period_years and (not period_years or period_years[-1] not in allowed_period_years):
            continue
        after = target_scope[match.end():match.end() + 90]
        comparison_match = re.search(
            r"(?:20\d{2}年(?:\d{1,2}月)?(?:期)?比\s*)?[+＋\-－]\s*\d[\d,]*(?:\.\d+)?\s*(?:%|％|人|件|社|拠点|店舗|台)",
            after,
        )
        comparison = comparison_match.group(0).replace(" ", "").replace("％", "%") if comparison_match else ""
        quote_end = match.end() + (comparison_match.end() if comparison_match else min(45, len(after)))
        evidence_excerpt = target_scope[max(0, match.start() - 18):min(len(target_scope), quote_end)].strip()
        source_metric = " ".join(part for part in (period, label, value, comparison) if part)
        targets.append({
            "claim_id": f"M1-{len(targets) + 1:02d}",
            "claim_status": "verified_external",
            "label": label,
            "target": value,
            "period": period,
            "comparison": comparison,
            "implication": "対象サービスのAI施策は、この経営目標への寄与仮説としてPoCで検証します。",
            "source_id": "M1",
            "source_url": str(plan.get("url") or ""),
            "source_locator": f"{period or '中期経営計画'}の{label}目標",
            "source_metric": source_metric,
            "evidence_excerpt": evidence_excerpt,
        })
        if len(targets) >= max_targets:
            break
    return targets
def materialize_midterm_plan_targets(assessment: dict, research: dict) -> list[dict]:
    """中計ページへ、原文照合済みの経営目標を決定的に補完する。"""
    plan = _midterm_plan_source(research)
    extracted = extract_verified_management_targets_from_midterm_plan(plan)
    analysis = assessment.get("midterm_plan_analysis")
    if not isinstance(analysis, dict):
        research_analysis = research.get("midterm_plan_analysis")
        analysis = copy.deepcopy(research_analysis) if isinstance(research_analysis, dict) else {}
    if not analysis.get("ai_alignment"):
        return []
    # LLMが再記述したclaimは、文字列が原典抜粋に含まれていても表の行・列や
    # 対象範囲まで照合できない。表示対象は上の決定的抽出結果だけに限定する。
    # 取れない場合は空配列にして、数値を伴わない方針スライドへ安全に縮退する。
    analysis["management_targets"] = extracted
    analysis.setdefault("source", {"title": plan.get("title", ""), "url": plan.get("url", "")})
    assessment["midterm_plan_analysis"] = analysis
    research["midterm_plan_analysis"] = copy.deepcopy(analysis)
    return list(analysis["management_targets"])
def _management_target_is_verifiable(
        target: object, research: object, *, require_fetched: bool = False) -> bool:
    """中計の数値目標も、URL・原典抜粋・表示値の対応を必須にする。"""
    if not isinstance(target, dict):
        return False
    plan = _midterm_plan_source(research)
    provenance = research.get("provenance") if isinstance(research, dict) else None
    enforce_fetched = require_fetched or (
        isinstance(provenance, dict)
        and provenance.get("format") == "ai-assess/reproducibility-v2"
    )
    if (not plan or (enforce_fetched and plan.get("fetch_status") != "fetched")
            or target.get("source_id") != "M1"
            or target.get("claim_status") != "verified_external"
            or not all(str(target.get(field) or "").strip() for field in (
                "claim_id", "source_url", "source_locator", "source_metric", "evidence_excerpt", "target",
            ))):
        return False
    if str(target.get("source_url")) != str(plan.get("url")):
        return False
    # PDFのテキスト抽出では表の行・列が連結することがある。URLと抜粋内の
    # 数字が一致するだけでは、別事業・別年度のKPIを中計目標として取り込めて
    # しまうため、同じ決定的抽出器が認めた「指標・値・期間」の組だけを通す。
    verified_targets = extract_verified_management_targets_from_midterm_plan(plan)
    target_key = tuple(_normalized_evidence_text(target.get(field)) for field in (
        "label", "target", "period",
    ))
    verified_keys = {
        tuple(_normalized_evidence_text(item.get(field)) for field in (
            "label", "target", "period",
        ))
        for item in verified_targets
    }
    if not target_key[0] or target_key not in verified_keys:
        return False
    target_tokens = _quantitative_tokens(target.get("target"))
    source_tokens = _quantitative_tokens(target.get("source_metric"))
    excerpt = _normalized_evidence_text(target.get("evidence_excerpt"))
    plan_excerpt = _normalized_evidence_text(plan.get("excerpt"))
    return bool(target_tokens and target_tokens <= source_tokens and len(excerpt) >= 8 and excerpt in plan_excerpt)
def _quantitative_claim_is_verifiable(
        benchmark: object, sources: dict[str, dict], approved_source_ids: set[str], *,
        require_fetched: bool = False) -> bool:
    if not isinstance(benchmark, dict):
        return False
    source_id = str(benchmark.get("source_id") or "")
    source = sources.get(source_id, {})
    # 検索結果のsnippetは原典候補の発見にだけ使い、数値claimの根拠にはしない。
    # fetch_statusがない旧レビューJSONだけは読み取り互換を保つが、新規調査で
    # 明示された状態は原文取得済み(fetched)以外をすべて不採用にする。
    fetch_status = str(source.get("fetch_status") or "").strip()
    if (require_fetched and fetch_status != "fetched") or (
            fetch_status and fetch_status != "fetched"):
        return False
    source_url = str(source.get("url") or "")
    claim_url = str(benchmark.get("source_url") or "")
    if (not source_id.startswith("R") or source_id not in approved_source_ids
            or not is_safe_public_https_url(source_url) or claim_url != source_url):
        return False
    if not all(str(benchmark.get(field) or "").strip() for field in (
        "claim_id", "claim_status", "source_locator", "source_metric", "evidence_excerpt", "headline_metric",
    )):
        return False
    if benchmark.get("claim_status") != "verified_external":
        return False
    headline_tokens = _quantitative_tokens(benchmark.get("headline_metric"))
    source_tokens = _quantitative_tokens(benchmark.get("source_metric"))
    excerpt = _normalized_evidence_text(benchmark.get("evidence_excerpt"))
    source_body = str(source.get("excerpt") or "")
    if require_fetched and not source_body.strip():
        return False
    # v2の原典claimは取得本文だけを照合する。検索snippetへのフォールバックは
    # fetch_statusを持たない既存v1 JSONの読み取り互換に限る。
    source_excerpt = _normalized_evidence_text(
        source_body if require_fetched else source_body or source.get("snippet")
    )
    return bool(headline_tokens and headline_tokens <= source_tokens and len(excerpt) >= 8 and excerpt in source_excerpt)
def _poc_measurement_design_is_verifiable(
        design: object, research: object, assessment: dict | None = None) -> bool:
    """定量ページの3件が、承認済み原典claimと一対一で再照合できるか確認する。"""
    normalized = normalize_poc_measurement_design(design, assessment)
    if normalized.get("status") != "external_verified" or not isinstance(research, dict):
        return False
    sources = _research_sources_by_id(research)
    approved_ids = _research_approved_quantitative_source_ids(research)
    items = normalized.get("items", [])
    if len(items) != 3:
        return False
    claim_ids = {str(item.get("claim_id") or "") for item in items}
    source_ids = {str(item.get("source_id") or "") for item in items}
    bindings = [
        {field: str(item.get(field) or "") for field in ("priority", "use_case_id", "theme")}
        for item in items if isinstance(item, dict)
    ]
    if len(claim_ids) != 3 or len(source_ids) != 3:
        return False
    research_provenance = research.get("provenance")
    require_fetched = bool(
        isinstance(research_provenance, dict)
        and research_provenance.get("format") == "ai-assess/reproducibility-v2"
    )
    if require_fetched and any(
            str(sources.get(source_id, {}).get("fetch_status") or "") != "fetched"
            for source_id in source_ids):
        return False
    if not all(_quantitative_claim_is_verifiable(
            item, sources, approved_ids, require_fetched=require_fetched,
    ) for item in items):
        return False
    if not all(
        _management_target_is_verifiable(target, research)
        for target in normalized.get("management_targets", [])
    ):
        return False
    audit = research.get("poc_measurement_design_audit")
    if not isinstance(audit, dict):
        return False
    return bool(
        audit.get("status") == "external_verified"
        and audit.get("basis_type") == "verified_external_benchmark"
        and set(str(item) for item in audit.get("claim_ids", [])) == claim_ids
        and set(str(item) for item in audit.get("approved_source_ids", [])) == source_ids
        and list(audit.get("poc_bindings") or []) == bindings
        and str(audit.get("reason") or "").strip()
    )
def _llm_quantitative_estimate_is_consistent(
        design: object, research: object, assessment: dict | None = None) -> bool:
    """LLM試算の表示値・監査ログ・最終採用モードが同じ3件を指すか確認する。"""
    normalized = normalize_poc_measurement_design(design, assessment)
    if normalized.get("status") != "decision_thresholds" or not isinstance(research, dict):
        return False
    items = normalized.get("items", [])
    if not isinstance(items, list) or len(items) != 3:
        return False
    estimate_ids = [str(item.get("estimate_id") or "") for item in items if isinstance(item, dict)]
    metrics = [str(item.get("headline_metric") or "") for item in items if isinstance(item, dict)]
    display_layers = [str(item.get("display_kpi_layer") or "") for item in items if isinstance(item, dict)]
    product_metrics = [
        str((item.get("product_kpi") or {}).get("target") or "")
        for item in items if isinstance(item, dict) and isinstance(item.get("product_kpi"), dict)
    ]
    bindings = [
        {field: str(item.get(field) or "") for field in ("priority", "use_case_id", "theme")}
        for item in items if isinstance(item, dict)
    ]
    if (len(estimate_ids) != 3 or len(set(estimate_ids)) != 3
            or not all(re.fullmatch(r"E\d{2}", value) for value in estimate_ids)):
        return False
    audit = research.get("poc_measurement_design_audit")
    estimate_log = research.get("llm_quantitative_estimate")
    analysis = research.get("quantitative_analysis")
    if not all(isinstance(value, dict) for value in (audit, estimate_log, analysis)):
        return False
    model_id = str(audit.get("model_id") or "")
    if not model_id:
        return False
    if not (
        audit.get("status") == "decision_thresholds"
        and audit.get("basis_type") == "decision_threshold"
        and audit.get("generation_source") == "llm_estimate"
        and list(audit.get("estimate_ids") or []) == estimate_ids
        and list(audit.get("displayed_metrics") or []) == metrics
        and list(audit.get("display_kpi_layers") or []) == display_layers
        and list(audit.get("product_leading_metrics") or []) == product_metrics
        and list(audit.get("poc_bindings") or []) == bindings
        and str(audit.get("reason") or "").strip()
    ):
        return False
    logged_rows = estimate_log.get("estimates")
    logged_ids = [
        str(item.get("estimate_id") or "") for item in logged_rows
        if isinstance(logged_rows, list) and isinstance(item, dict)
    ] if isinstance(logged_rows, list) else []
    logged_metrics = [
        str(item.get("headline_metric") or "") for item in logged_rows
        if isinstance(logged_rows, list) and isinstance(item, dict)
    ] if isinstance(logged_rows, list) else []
    logged_bindings = [
        {field: str(item.get(field) or "") for field in ("priority", "use_case_id", "theme")}
        for item in logged_rows if isinstance(logged_rows, list) and isinstance(item, dict)
    ] if isinstance(logged_rows, list) else []
    logged_layers = [
        str(item.get("display_kpi_layer") or "") for item in logged_rows
        if isinstance(logged_rows, list) and isinstance(item, dict)
    ] if isinstance(logged_rows, list) else []
    logged_product_metrics = [
        str((item.get("product_kpi") or {}).get("target") or "") for item in logged_rows
        if isinstance(logged_rows, list) and isinstance(item, dict)
        and isinstance(item.get("product_kpi"), dict)
    ] if isinstance(logged_rows, list) else []
    if not (
        estimate_log.get("status") == "generated"
        and estimate_log.get("display_mode") == "llm_estimate"
        and estimate_log.get("model_id") == model_id
        and estimate_log.get("estimate_count") == 3
        and list(estimate_log.get("estimate_ids") or []) == estimate_ids
        and logged_ids == estimate_ids
        and logged_metrics == metrics
        and logged_bindings == bindings
        and logged_layers == display_layers
        and logged_product_metrics == product_metrics
        and all(
            isinstance(item, dict)
            and item.get("display_kpi_layer") in BUSINESS_EFFECT_DISPLAY_LAYERS
            and isinstance(item.get("product_kpi"), dict)
            and isinstance(item.get(item.get("display_kpi_layer")), dict)
            and str(item.get("headline_metric") or "")
                == str(item[item["display_kpi_layer"]].get("target") or "")
            for item in logged_rows
        )
        and estimate_log.get("customer_actuals_assumed") is False
    ):
        return False
    return bool(
        analysis.get("status") == "llm_estimated_targets"
        and analysis.get("final_display_mode") == "llm_estimate"
        and analysis.get("llm_estimate_count") == 3
        and list(analysis.get("llm_estimate_ids") or []) == estimate_ids
        and analysis.get("llm_model_id") == model_id
        and analysis.get("display_contract") == "poc_business_effect_thresholds"
        and analysis.get("customer_actuals_assumed") is False
        and analysis.get("stopped_before_output") is False
    )
def _ai_product_business_impact_is_consistent(
        model: object, research: object, assessment: dict) -> bool:
    """Verify that the P6 company KPIs and their independent audit are identical."""
    normalized = normalize_ai_product_business_impact_candidate(model, assessment)
    if not normalized or not isinstance(research, dict):
        return False
    audit = research.get("ai_product_business_impact_audit")
    logged_model = research.get("ai_product_business_impact")
    if not isinstance(audit, dict) or not isinstance(logged_model, dict):
        return False
    logged_normalized = normalize_ai_product_business_impact_candidate(logged_model, assessment)
    if logged_normalized != normalized:
        return False
    items = normalized["items"]
    estimate_ids = [str(item["estimate_id"]) for item in items]
    business_kpi_ids = [str(item["business_kpi_id"]) for item in items]
    categories = [str(item["category"]) for item in items]
    displayed_targets = [str(item["headline_metric"]) for item in items]
    formulas = [str(item["formula"]) for item in items]
    base_consistent = bool(
        audit.get("schema_version") == AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION
        and audit.get("status") == "decision_thresholds"
        and audit.get("generation_source") in {"llm_decision_threshold", "llm_estimate"}
        and audit.get("display_contract") == "assessed_company_ai_business_impact"
        and str(audit.get("model_id") or "")
        and audit.get("prompt_contract_version") == AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION
        and audit.get("estimate_count") == 3
        and list(audit.get("estimate_ids") or []) == estimate_ids
        and list(audit.get("business_kpi_ids") or []) == business_kpi_ids
        and list(audit.get("categories") or []) == categories
        and list(audit.get("displayed_targets") or []) == displayed_targets
        and list(audit.get("formulas") or []) == formulas
        and audit.get("assessed_company_actuals_assumed") is False
        and audit.get("customer_poc_metrics_reused") is False
        and str(audit.get("reason") or "").strip()
    )
    if not base_consistent:
        return False
    provenance = research.get("provenance")
    require_planning_contract = bool(audit.get("evidence_mode") is not None or (
        isinstance(provenance, dict)
        and provenance.get("format") == "ai-assess/reproducibility-v2"))
    if not require_planning_contract:
        return True
    management_targets = (normalized.get("management_targets")
                          if isinstance(normalized.get("management_targets"), list) else [])
    if any(not _management_target_is_verifiable(
            target, research, require_fetched=True) for target in management_targets):
        return False
    expected_management_claim_ids = [str(target.get("claim_id") or "")
                                     for target in management_targets if isinstance(target, dict)]
    expected_assumptions = ai_product_business_impact_calculation_assumptions(normalized)
    source_ids = audit.get("external_source_ids")
    approved_ids = _ai_product_business_impact_external_source_ids(research)
    if not isinstance(source_ids, list):
        return False
    normalized_source_ids = sorted({str(source_id) for source_id in source_ids
                                    if isinstance(source_id, str)
                                    and re.fullmatch(r"R\d+", source_id)})
    external_count = audit.get("external_evidence_count")
    return bool(
        audit.get("evidence_mode") == "llm_estimate"
        and audit.get("estimate_classification") == "planning_estimate"
        and audit.get("display_label") == "AI導入効果目標"
        and isinstance(external_count, int)
        and not isinstance(external_count, bool)
        and external_count == len(normalized_source_ids)
        and normalized_source_ids == approved_ids
        and list(audit.get("management_target_claim_ids") or [])
        == expected_management_claim_ids
        and list(audit.get("calculation_assumptions") or []) == expected_assumptions
        and str(audit.get("generation_reason") or "").strip()
        and audit.get("actual_result_claimed") is False
    )
def ai_product_business_impact_calculation_assumptions(model: object) -> list[dict[str, str]]:
    """Return the exact, reviewable assumptions behind each planning range."""
    if not isinstance(model, dict):
        return []
    rows: list[dict[str, str]] = []
    for item in model.get("items", []) if isinstance(model.get("items"), list) else []:
        if not isinstance(item, dict):
            continue
        leading = item.get("leading_indicator")
        rows.append({
            "estimate_id": str(item.get("estimate_id") or ""),
            "business_kpi_id": str(item.get("business_kpi_id") or ""),
            "baseline_definition": str(item.get("baseline_definition") or ""),
            "comparison_condition": str(item.get("comparison_condition") or ""),
            "formula": str(item.get("formula") or ""),
            "target_rationale": str(item.get("target_rationale") or ""),
            "leading_indicator_formula": str(leading.get("formula") or "")
            if isinstance(leading, dict) else "",
        })
    return rows
def _ai_product_business_impact_external_source_ids(research: object) -> list[str]:
    """Return externally verified IDs relevant when the LLM estimate was made."""
    if not isinstance(research, dict):
        return []
    display = research.get("quantitative_display_contract")
    if isinstance(display, dict) and isinstance(display.get("approved_source_ids"), list):
        return sorted({str(source_id) for source_id in display["approved_source_ids"]
                       if isinstance(source_id, str) and re.fullmatch(r"R\d+", source_id)})
    return sorted(_research_approved_quantitative_source_ids(research))
def materialize_ai_product_business_impact_contract(
        assessment: dict, research: dict, model: object, *, model_id: str,
        analysis_log: list[dict] | None = None,
        generation_reason: str,
        external_source_ids: list[str] | tuple[str, ...] | None = None) -> dict:
    """Freeze three company planning targets separately from external results."""
    if not isinstance(assessment, dict) or not isinstance(research, dict):
        raise TypeError("assessmentとresearchはdictで指定してください。")
    normalized = normalize_ai_product_business_impact_candidate(model, assessment)
    issues = ai_product_business_impact_contract_issues(normalized or model, assessment)
    if not normalized or issues:
        raise ValueError(
            "AI導入効果目標の計画試算が契約を満たしません: " + " / ".join(issues)
        )
    model_id = str(model_id or "").strip()
    if not model_id:
        raise ValueError("AI導入効果目標には生成モデルIDが必要です。")
    generation_reason = str(generation_reason or "").strip()
    if not generation_reason:
        raise ValueError("AI導入効果目標には生成理由が必要です。")
    approved_ids = _ai_product_business_impact_external_source_ids(research)
    declared_ids = sorted({str(source_id) for source_id in (external_source_ids or approved_ids)
                           if isinstance(source_id, str) and re.fullmatch(r"R\d+", source_id)})
    if declared_ids != approved_ids:
        raise ValueError("外部定量根拠IDがresearch.quantitative_evidenceと一致しません。")
    items = normalized["items"]
    audit = {
        "schema_version": AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION,
        "status": "decision_thresholds",
        "generation_source": "llm_estimate",
        "evidence_mode": "llm_estimate",
        "estimate_classification": "planning_estimate",
        "display_label": "AI導入効果目標",
        "display_contract": "assessed_company_ai_business_impact",
        "model_id": model_id,
        "prompt_contract_version": AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION,
        "estimate_count": 3,
        "estimate_ids": [str(item["estimate_id"]) for item in items],
        "business_kpi_ids": [str(item["business_kpi_id"]) for item in items],
        "categories": [str(item["category"]) for item in items],
        "displayed_targets": [str(item["headline_metric"]) for item in items],
        "formulas": [str(item["formula"]) for item in items],
        "calculation_assumptions": ai_product_business_impact_calculation_assumptions(
            normalized,
        ),
        "generation_reason": generation_reason,
        "external_evidence_count": len(declared_ids),
        "external_source_ids": declared_ids,
        "management_target_claim_ids": [
            str(target.get("claim_id") or "")
            for target in normalized.get("management_targets", [])
            if isinstance(target, dict)
        ],
        "assessed_company_actuals_assumed": False,
        "customer_poc_metrics_reused": False,
        "actual_result_claimed": False,
        "reason": (
            "対象企業・製品の事業モデルと公開計画を基に、割合の計画試算、または原典照合済み"
            "中期経営目標への接続を作成しました。外部導入実績や達成済み効果とは区別します。"
        ),
    }
    assessment["ai_product_business_impact"] = copy.deepcopy(normalized)
    research["ai_product_business_impact"] = copy.deepcopy(normalized)
    research["ai_product_business_impact_audit"] = audit
    research["ai_product_business_impact_generation_log"] = copy.deepcopy(
        analysis_log if isinstance(analysis_log, list) else [])
    if not _ai_product_business_impact_is_consistent(normalized, research, assessment):
        raise RuntimeError("AI導入効果目標と監査ログの固定に失敗しました。")
    return copy.deepcopy(normalized)
def sanitize_quantitative_evidence(
        assessment: dict, research: dict, *, migrate_legacy: bool = False,
        require_fetched: bool = False) -> list[dict[str, str]]:
    """未裏付けの定量主張を描画データから除き、判断理由を監査台帳へ残す。"""
    evidence = assessment.get("executive_evidence")
    if not isinstance(evidence, dict):
        return []
    sources = _research_sources_by_id(research)
    approved_ids = _research_approved_quantitative_source_ids(research)
    prior_benchmarks = evidence.get("benchmarks") if isinstance(evidence.get("benchmarks"), list) else []
    prior_scenarios = evidence.get("value_scenarios") if isinstance(evidence.get("value_scenarios"), list) else []
    prior_targets = evidence.get("management_targets") if isinstance(evidence.get("management_targets"), list) else []
    mode = str(evidence.get("evidence_mode") or "").strip()
    kept = [
        benchmark for benchmark in prior_benchmarks
        if mode == "external_verified" and _quantitative_claim_is_verifiable(
            benchmark, sources, approved_ids, require_fetched=require_fetched,
        )
    ]
    audit_rows: list[dict[str, str]] = []
    for index, benchmark in enumerate(prior_benchmarks, 1):
        if benchmark not in kept:
            audit_rows.append({
                "claim_id": str(benchmark.get("claim_id") or f"legacy-Q{index:02d}") if isinstance(benchmark, dict) else f"legacy-Q{index:02d}",
                "decision": "omitted",
                "reason": "外部根拠のURL・承認済みsource_id・数値を含む原典抜粋のいずれかを確認できないため、定量主張として表示しません。",
            })
    kept_targets = [
        target for target in prior_targets
        if _management_target_is_verifiable(
            target, research, require_fetched=require_fetched,
        )
    ]
    for index, target in enumerate(prior_targets, 1):
        if target not in kept_targets:
            audit_rows.append({
                "claim_id": str(target.get("claim_id") or f"legacy-M{index:02d}") if isinstance(target, dict) else f"legacy-M{index:02d}",
                "decision": "omitted",
                "reason": "中期経営計画の数値は、原典URL・数値を含む抜粋・表示値の対応を確認できないため表示しません。",
            })
    evidence["management_targets"] = kept_targets
    legacy_quantitative = assessment.get("quantitative_analysis")
    if isinstance(legacy_quantitative, dict) and legacy_quantitative.get("metrics"):
        assessment["quantitative_analysis"] = {}
        audit_rows.append({
            "claim_id": "legacy-quantitative-analysis",
            "decision": "omitted",
            "reason": "旧quantitative_analysisはclaimごとの原典照合契約を持たないため、表示対象から除外しました。",
        })
    if kept:
        kept_source_ids = {str(item["source_id"]) for item in kept}
        evidence["evidence_mode"] = "external_verified"
        evidence["benchmarks"] = kept
        evidence["value_scenarios"] = [
            item for item in prior_scenarios if isinstance(item, dict)
            and isinstance(item.get("source_ids"), list)
            and set(str(source_id) for source_id in item["source_ids"]) <= kept_source_ids
        ]
        research["quantitative_evidence"] = {
            "schema_version": QUANTITATIVE_EVIDENCE_SCHEMA_VERSION,
            "status": "sufficient" if len(kept) >= 3 else "partial",
            "approved_source_ids": sorted(kept_source_ids),
            "reason": "独立した根拠監査を通過し、原典URL・数値表現・抜粋を確認できた主張だけを表示します。",
        }
    elif prior_benchmarks or mode == "hypothesis" or migrate_legacy:
        # 数字付きのPoC仮説は測定設計へ戻す。仮説を事実のように再利用できる
        # JSONへ残さず、値は再承認されるまでチャーターのconfirmで扱う。
        evidence["evidence_mode"] = "pre_poc"
        evidence["benchmarks"] = []
        evidence["value_scenarios"] = []
        research["quantitative_evidence"] = {
            "schema_version": QUANTITATIVE_EVIDENCE_SCHEMA_VERSION,
            "status": "insufficient",
            "approved_source_ids": [],
            "reason": "公開根拠または顧客承認済みの定量前提を確認できないため、数値目標はPoCの基準値・比較条件の合意後に設定します。",
        }
    effective_mode = str(evidence.get("evidence_mode") or "").strip()
    display_source_ids = {str(item.get("source_id")) for item in kept if isinstance(item, dict)}
    if kept_targets:
        display_source_ids.add("M1")
    display_sources: list[dict[str, str]] = []
    plan = _midterm_plan_source(research)
    if "M1" in display_source_ids and plan:
        display_sources.append({"id": "M1", "title": str(plan.get("title") or "中期経営計画"), "url": str(plan["url"])})
    for source_id in sorted(display_source_ids):
        source = sources.get(source_id)
        if source and is_safe_public_https_url(source.get("url")):
            display_sources.append({
                "id": source_id, "title": str(source.get("title") or source_id), "url": str(source["url"]),
            })
    if effective_mode == "pre_poc":
        # 根拠が揃わない旧JSONでは、数値を含む導入判断文・H系疑似ソースも残さない。
        # 何を測るかはPoCチャーターに保持し、目標値は開始ゲートの合意後に入力する。
        if not kept_targets:
            evidence["plan_evidence_summary"] = "公開資料に記載された事業方針と、対象サービスでAIが支援する業務判断を接続して検討します。"
            evidence["ai_necessity_analysis"] = (
                "対象サービスの業務・データ・利用者を踏まえ、固定ルールや担当者の経験に依存する判断を、"
                "根拠付きの候補提示と人の承認へ転換する余地を検討します。PoCでは対象範囲、データ品質、"
                "基準値、比較条件、停止条件、責任者を合意し、実測結果と横展開条件から標準化の可否を判断します。"
            )
        evidence["decision_message"] = (
            "代表対象と責任者を定め、基準値・比較条件・成功／停止条件を合意したテーマから段階的にPoCを開始し、"
            "実測結果と横展開条件を基に標準化の可否を判断する。"
        )
        evidence["caveat"] = (
            "効果率・期間・金額は、公開根拠または顧客承認済みの定量前提を確認したうえで提示し、"
            "PoCの実測結果を判断材料として扱う想定です。"
        )
        evidence["evidence_gap_message"] = "数値目標は、対象データ・基準値・比較条件の合意後にPoCで設定します。"
        if plan:
            display_sources = [{"id": "M1", "title": str(plan.get("title") or "中期経営計画"), "url": str(plan["url"])}]
    evidence["sources"] = display_sources
    if audit_rows:
        existing = research.get("quantitative_claim_audit")
        prior_audit = existing if isinstance(existing, list) else []
        research["quantitative_claim_audit"] = [*prior_audit, *audit_rows]
    return audit_rows
def displayed_quantitative_effect_issues(assessment: object, research: object) -> list[str]:
    """Reject effect numbers outside fetched claims or audited planning ranges.
    The prompt already asks the model not to invent metrics, but a frozen review
    contract must enforce that rule independently.  Company/product-level LLM
    ranges are allowed only inside the dedicated planning-estimate contract;
    they are never accepted as external results. Scores, page numbers, WBS
    periods and cost snapshots remain outside this assessment-text check.
    """
    if not isinstance(assessment, dict) or not isinstance(research, dict):
        return []
    display_contract = research.get("quantitative_display_contract")
    numeric_allowed = bool(
        isinstance(display_contract, dict)
        and display_contract.get("display_mode") == "external_verified"
        and display_contract.get("numeric_claims_allowed") is True
    )
    planning_estimate_allowed = _ai_product_business_impact_is_consistent(
        assessment.get("ai_product_business_impact"), research, assessment,
    )
    always_allowed = (
        "assessment.poc_measurement_design",
        "assessment.midterm_plan_analysis.management_targets",
        "assessment.executive_evidence.management_targets",
        # Candidate-ranking weights are an internal decision formula, not a
        # claimed customer or provider business effect.
        "assessment.poc_priority_decision.weights",
    )
    externally_verified_allowed = (
        "assessment.executive_evidence.benchmarks",
        "assessment.executive_evidence.value_scenarios",
    )
    issues: list[str] = []
    def walk(value: object, path: str) -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                walk(child, f"{path}.{key}")
            return
        if isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{path}[{index}]")
            return
        if not isinstance(value, str) or not _is_business_effect_metric(value):
            return
        if any(path.startswith(prefix) for prefix in always_allowed):
            return
        if numeric_allowed and any(
                path.startswith(prefix) for prefix in externally_verified_allowed):
            return
        if planning_estimate_allowed and path.startswith(
                "assessment.ai_product_business_impact"):
            return
        issues.append(
            f"{path}の割合・金額・生産性効果は、fetch_status=fetchedの"
            "構造化定量claimへ追跡できません。"
        )
    walk(assessment, "assessment")
    return issues
def validate_quantitative_evidence(payload: dict, *, allow_legacy_hypotheses: bool = False) -> list[str]:
    """数値を伴う主張は、原典と結び付く表示契約を満たす場合だけ許可する。"""
    errors: list[str] = []
    assessment = payload.get("assessment")
    research = payload.get("research")
    if not isinstance(assessment, dict) or not isinstance(research, dict):
        return errors
    display_contract = research.get("quantitative_display_contract")
    if display_contract is not None:
        errors.extend(quantitative_display_contract_issues(
            display_contract, assessment, research,
        ))
    if isinstance(assessment.get("quantitative_analysis"), dict) and assessment["quantitative_analysis"].get("metrics"):
        errors.append("assessment.quantitative_analysisは検証済みclaim契約を持たないため使用できません。")
    evidence = assessment.get("executive_evidence")
    if not isinstance(evidence, dict):
        return errors
    mode = str(evidence.get("evidence_mode") or "").strip()
    benchmarks = evidence.get("benchmarks") if isinstance(evidence.get("benchmarks"), list) else []
    scenarios = evidence.get("value_scenarios") if isinstance(evidence.get("value_scenarios"), list) else []
    management_targets = evidence.get("management_targets") if isinstance(evidence.get("management_targets"), list) else []
    sources = _research_sources_by_id(research)
    approved_ids = _research_approved_quantitative_source_ids(research)
    for index, target in enumerate(management_targets, 1):
        if not _management_target_is_verifiable(target, research):
            errors.append(f"中期経営計画の定量claim {index} はURL・原典抜粋・表示値との照合に失敗しました。")
    midterm_analysis = assessment.get("midterm_plan_analysis")
    if isinstance(midterm_analysis, dict):
        for index, target in enumerate(
                midterm_analysis.get("management_targets", [])
                if isinstance(midterm_analysis.get("management_targets"), list) else [], 1):
            if not _management_target_is_verifiable(target, research):
                errors.append(
                    f"中期経営計画ページの定量claim {index} はURL・原典取得状態・"
                    "数値抜粋との照合に失敗しました。"
                )
    if mode == "external_verified":
        if not benchmarks:
            errors.append("external_verifiedには少なくとも1件の検証済みbenchmarkが必要です。")
        claim_ids: set[str] = set()
        for index, benchmark in enumerate(benchmarks, 1):
            if not _quantitative_claim_is_verifiable(benchmark, sources, approved_ids):
                errors.append(f"定量claim {index} は承認済み原典・URL・数値抜粋との照合に失敗しました。")
                continue
            claim_id = str(benchmark.get("claim_id"))
            if claim_id in claim_ids:
                errors.append(f"定量claim_id {claim_id} が重複しています。")
            claim_ids.add(claim_id)
        benchmark_ids = {str(item.get("source_id")) for item in benchmarks if isinstance(item, dict)}
        benchmark_by_id = {
            str(item.get("source_id")): item for item in benchmarks if isinstance(item, dict)
        }
        for index, scenario in enumerate(scenarios, 1):
            scenario_ids = scenario.get("source_ids") if isinstance(scenario, dict) else None
            if (not isinstance(scenario_ids, list) or len(scenario_ids) != 1
                    or str(scenario_ids[0]) not in benchmark_ids):
                errors.append(f"value_scenarios {index} は表示済みbenchmark以外の数値根拠を参照しています。")
                continue
            benchmark = benchmark_by_id[str(scenario_ids[0])]
            metric_tokens = _quantitative_tokens(benchmark.get("headline_metric"))
            if (re.sub(r"\s+", "", str(scenario.get("use_case") or ""))
                    != re.sub(r"\s+", "", str(benchmark.get("use_case") or ""))
                    or re.sub(r"\s+", "", str(scenario.get("benchmark") or ""))
                    != re.sub(r"\s+", "", str(benchmark.get("headline_metric") or ""))
                    or not metric_tokens <= _quantitative_tokens(scenario.get("formula"))
                    or not metric_tokens <= _quantitative_tokens(scenario.get("poc_gate"))):
                errors.append(f"value_scenarios {index} はbenchmarkの業務・数値・算定式・判定基準と一致しません。")
    elif mode in {"", "pre_poc"}:
        if benchmarks or scenarios:
            errors.append("根拠未確認の定量主張はpre_pocでは表示できません。")
        if mode == "pre_poc":
            # 数字付きの開始判断文や疑似ソースを残すと、カードを消してもスライド本文で
            # 未根拠の効果率・期間が再表示される。定量は検証済みtarget以外許容しない。
            display_text = " ".join(str(evidence.get(field) or "") for field in (
                "plan_evidence_summary", "ai_necessity_analysis", "decision_message", "caveat", "evidence_gap_message",
            ))
            logic = evidence.get("strategic_logic") if isinstance(evidence.get("strategic_logic"), list) else []
            premises = evidence.get("strategic_premises") if isinstance(evidence.get("strategic_premises"), list) else []
            display_text += " " + json.dumps([*logic, *premises], ensure_ascii=False)
            if _quantitative_tokens(display_text):
                errors.append("pre_poc本文には未検証の数値・期間・金額を含められません。")
            for source in evidence.get("sources", []) if isinstance(evidence.get("sources"), list) else []:
                if (not isinstance(source, dict) or str(source.get("id") or "").startswith("H")
                        or not is_safe_public_https_url(source.get("url"))):
                    errors.append("pre_pocのsourcesには公開URLを持つ一次資料だけを残せます。")
    elif mode == "hypothesis":
        if (benchmarks or scenarios) and not allow_legacy_hypotheses:
            errors.append("数値付きhypothesisは承認対象にできません。PoCのconfirm項目へ移行してください。")
    else:
        errors.append("executive_evidence.evidence_modeが不正です。")
    return errors
