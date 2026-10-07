import unittest

from ai_assess_runtime.poc_contract import (
    _polish_technical_display_text,
    display_theme_name,
)
from ai_assess_runtime.presentation import (
    _business_value_measurement,
    _industry_reference_urls,
    assessment_intro_copy_for,
)
from ai_assess_runtime.quantitative_contract import (
    extract_verified_management_targets_from_midterm_plan,
    materialize_midterm_plan_targets,
)
from ai_assess_runtime.target_metric import _normalize_percentage_target


class AssessmentContentRevisionTest(unittest.TestCase):
    def test_percentage_range_accepts_percent_on_both_bounds(self) -> None:
        self.assertEqual("20〜35%", _normalize_percentage_target("20%〜35%"))

    def test_optimization_title_is_softened_without_solver(self) -> None:
        detail = {
            "theme": "生産優先順位最適化",
            "oracle_technologies": "Oracle Machine Learning（OML4SQL）／SQL制約ロジック",
        }
        poc = {"theme": detail["theme"], "technical_modality": "optimization"}
        self.assertEqual("生産優先順位提案", display_theme_name(detail, poc))

    def test_technical_copy_normalizes_version_ranges_and_link_role(self) -> None:
        text = _polish_technical_display_text(
            "Oracle Database 23aiで過去12～24か月を扱い、Database Link・APIで返却します。"
        )
        self.assertIn("Oracle AI Database 26ai", text)
        self.assertIn("対象履歴期間", text)
        self.assertIn("Database Linkで参照、またはAPIで返却", text)

    def test_business_formulas_match_value_area(self) -> None:
        delivery = _business_value_measurement("納期・売上機会", 0)
        productivity = _business_value_measurement("業務生産性・標準化", 1)
        inventory = _business_value_measurement("在庫・運転資金", 2)
        self.assertIn("粗利額", delivery[1])
        self.assertIn("工数削減", productivity[1])
        self.assertEqual("運転資金改善額＝対象在庫金額×在庫削減率", inventory[1])

    def test_intro_subtitle_uses_assessment_label(self) -> None:
        copy = assessment_intro_copy_for({
            "company_name": "株式会社エクス",
            "service_name": "Factory-ONE 電脳工場",
        })
        self.assertEqual("AIユースケースアセスメント", copy["subtitle"])

    def test_industry_references_use_official_and_audited_sources(self) -> None:
        sources = [
            {"id": "R1", "title": "別会社の中期経営計画", "url": "https://wrong.example/plan", "fetch_status": "fetched"},
            {"id": "R2", "title": "Factory-ONE 電脳工場 - 株式会社エクス", "url": "https://www.xeex.co.jp/", "fetch_status": "fetched"},
            {"id": "R3", "title": "製造業AI生産計画", "url": "https://industry.example/ai", "fetch_status": "fetched"},
        ]
        research = {
            "industry_sources": sources,
            "research_audit": [{"audit": {"usable_source_ids": ["R3"]}}],
        }
        self.assertEqual(
            ["https://www.xeex.co.jp/", "https://industry.example/ai"],
            _industry_reference_urls(
                research,
                {"company_name": "株式会社エクス", "service_name": "Factory-ONE 電脳工場"},
            ),
        )

    def test_management_target_extraction_rejects_mixed_pdf_columns_and_non_plan_targets(self) -> None:
        plan = {
            "status": "found",
            "url": "https://example.com/plan.pdf",
            "title": "中期経営計画(2024-2026)",
            "excerpt": (
                "重要経営指標 2027年3月期 中計目標 売上高 EPS 16%超 13%超 "
                "2033年3月期 売上高 1,200億円 EBITDA 20–25% 500億円"
            ),
        }
        self.assertEqual([], extract_verified_management_targets_from_midterm_plan(plan))

    def test_materialize_management_targets_omits_llm_claim_when_deterministic_extraction_is_empty(self) -> None:
        plan = {
            "status": "found",
            "url": "https://example.com/plan.pdf",
            "title": "中期経営計画(2024-2026)",
            "excerpt": "重要経営指標 2027年3月期 中計目標 売上高 EPS 16%超",
        }
        assessment = {
            "midterm_plan_analysis": {
                "ai_alignment": [{"plan_priority": "方針"}],
                "management_targets": [{"target": "1,200億円"}],
            },
        }
        research = {"midterm_plan": plan}
        self.assertEqual([], materialize_midterm_plan_targets(assessment, research))
        self.assertEqual([], assessment["midterm_plan_analysis"]["management_targets"])


if __name__ == "__main__":
    unittest.main()
