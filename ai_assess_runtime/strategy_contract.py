"""戦略ロジックの決定的な正規化。

LLM、Web、PPTXへ依存しない純粋関数だけを置き、調査結果の表現揺れが
後続の定量契約へ波及しないようにする。
"""


def normalize_strategic_logic(raw_logic: object) -> list[dict[str, str]]:
    """Return decision-grade strategic cards, or reject generic card-only output."""
    if not isinstance(raw_logic, list) or len(raw_logic) < 3:
        return []
    required_fields = (
        ("fact", "decision_implication"),
        ("current_constraint", "ai_decision_change"),
        ("delay_risk", "proof_conditions"),
    )
    detail_labels = (
        ("【事実】", "【経営判断】"),
        ("【現状制約】", "【AIで変える判断】"),
        ("【遅延リスク】", "【事業化条件】"),
    )
    normalized: list[dict[str, str]] = []
    for index, fields in enumerate(required_fields):
        item = raw_logic[index]
        if not isinstance(item, dict) or not all(
                str(item.get(field, "")).strip() for field in ("label", "headline", *fields)):
            return []
        first, second = (str(item[field]).strip() for field in fields)
        first_label, second_label = detail_labels[index]
        normalized.append({
            "label": str(item["label"]).strip(),
            "headline": str(item["headline"]).strip(),
            "detail": f"{first_label}{first} {second_label}{second}",
        })
    return normalized


def normalize_strategic_premises(raw_premises: object) -> list[dict[str, str]]:
    """Validate non-plan decision premises without presenting them as company targets."""
    if not isinstance(raw_premises, list):
        return []
    premises: list[dict[str, str]] = []
    premise_labels = ("事業機会", "優先テーマ", "事業化条件")
    for index, item in enumerate(raw_premises[:3]):
        if not isinstance(item, dict) or not all(
                str(item.get(field, "")).strip() for field in ("label", "headline", "detail", "basis")):
            continue
        premise = {field: str(item[field]).strip()
                   for field in ("label", "headline", "detail", "basis")}
        premise["label"] = premise_labels[index]
        premises.append(premise)
    return premises
