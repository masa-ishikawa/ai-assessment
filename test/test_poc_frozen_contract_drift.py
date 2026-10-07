"""PoC優先順位確定後の凍結JSON間のidentityずれを検知する回帰テスト。"""

from __future__ import annotations

import copy
from types import SimpleNamespace
import unittest

import generate_assessment as generator
from ai_assess_runtime import workflow
from test import test_generate_assessment as fixtures


class PocFrozenContractDriftTests(unittest.TestCase):
    """P1/P2/P3が連番でない実運用ケースも含め、全派生契約を拘束する。"""

    @staticmethod
    def _payload_with_non_contiguous_priorities() -> dict:
        assessment = fixtures.DynamicPptxPageTests.detailed_assessment()
        front = assessment["consulting_front_matter"]

        # 15件の中から UC01 / UC02 / UC08 が加重評価で選ばれる、
        # 実行時に起きる非連番の優先候補を汎用fixtureで再現する。
        for candidate in front["use_case_prioritization"]["candidates"]:
            candidate.update({
                "value": "low",
                "feasibility": "low",
                "data_readiness": "low",
                "scale": "low",
                "priority": "Watch",
            })
        for use_case_no, scale in ((1, "high"), (2, "medium"), (8, "medium")):
            front["use_case_prioritization"]["candidates"][use_case_no - 1].update({
                "value": "high",
                "feasibility": "high",
                "data_readiness": "high",
                "scale": scale,
            })
        front = generator.normalize_consulting_front_matter(front, assessment, {"I1"})
        assessment["consulting_front_matter"] = front
        assessment["poc_recommendations"] = [
            {
                "priority": priority,
                "use_case_id": f"UC{use_case_no:02d}",
                "use_case_no": str(use_case_no),
                "theme": f"ユースケース{use_case_no}",
                "reason": f"ユースケース{use_case_no}の価値を比較できる。",
                "first_step": f"ユースケース{use_case_no}の代表データを確認する。",
                "depends_on": "対象データ、比較条件、判断責任者を確認する。",
                "basis": "分析仮説",
                "architecture_implementation": "",
            }
            for priority, use_case_no in (("P1", 1), ("P2", 2), ("P3", 8))
        ]

        generator.materialize_poc_portfolio(assessment, front)
        # LLM応答後と同様、チャーターが先に存在する状態で全契約を凍結する。
        assessment["poc_charters"] = generator.fallback_poc_charters(assessment, front)
        assessment["multitenant_governance"] = generator.fallback_multitenant_governance()
        generator.materialize_assessment_decision_contract(assessment, front)
        return {
            "format": generator.ASSESSMENT_JSON_FORMAT,
            "assessment": assessment,
            "research": {"midterm_plan": {"status": "disabled"}, "industry_sources": []},
        }

    def test_non_contiguous_final_pocs_keep_every_frozen_contract_aligned(self) -> None:
        payload = self._payload_with_non_contiguous_priorities()
        assessment = payload["assessment"]
        expected = [("P1", "UC01"), ("P2", "UC02"), ("P3", "UC08")]

        self.assertEqual(expected, [
            (item["priority"], item["use_case_id"])
            for item in assessment["poc_portfolio"]["items"]
        ])
        self.assertEqual(expected, [
            (item["priority"], item["use_case_id"])
            for item in assessment["poc_charters"]["charters"]
        ])
        self.assertEqual(expected, [
            (item["priority"], item["use_case_id"])
            for item in assessment["poc_selection_scorecard"]["items"]
        ])
        self.assertEqual(expected, [
            (item["priority"], item["use_case_id"])
            for item in assessment["poc_priority_decision"]["items"]
            if item["priority"] in {"P1", "P2", "P3"}
        ])
        self.assertEqual(expected, [
            (item["priority"], item["use_case_id"])
            for item in assessment["poc_start_readiness"]["items"]
        ])
        self.assertEqual([], generator.validate_assessment_payload(payload))

    def test_each_frozen_poc_contract_rejects_identity_or_derived_value_drift(self) -> None:
        cases = (
            (
                "portfolio_missing_id",
                "poc_portfolio",
                lambda assessment: assessment["poc_portfolio"]["items"][2].pop("use_case_id"),
            ),
            (
                "portfolio_rank_drift",
                "poc_portfolio",
                lambda assessment: assessment["poc_portfolio"]["items"][2].update(priority="P2"),
            ),
            (
                "charter_identity_drift",
                "poc_charters",
                lambda assessment: assessment["poc_charters"]["charters"][2].update(
                    use_case_id="UC03",
                ),
            ),
            (
                "scorecard_identity_drift",
                "poc_selection_scorecard",
                lambda assessment: assessment["poc_selection_scorecard"]["items"][2].update(
                    use_case_id="UC03",
                ),
            ),
            (
                "priority_decision_score_drift",
                "poc_priority_decision",
                lambda assessment: assessment["poc_priority_decision"]["items"][0].update(
                    ranking_index=4.9,
                ),
            ),
            (
                "start_readiness_gate_count_drift",
                "poc_start_readiness",
                lambda assessment: assessment["poc_start_readiness"]["items"][0].update(
                    open_gate_count=0,
                ),
            ),
        )
        for label, error_fragment, mutate in cases:
            with self.subTest(case=label):
                payload = copy.deepcopy(self._payload_with_non_contiguous_priorities())
                mutate(payload["assessment"])
                errors = generator.validate_assessment_payload(payload)
                self.assertTrue(
                    any(error_fragment in error for error in errors),
                    f"{label}を検知できません: {errors}",
                )

    def test_customer_priority_binding_survives_repeated_front_matter_normalization(self) -> None:
        """Weighted P3=UC08より、顧客が明示したP3=UC03を凍結契約まで保つ。"""
        assessment = fixtures.DynamicPptxPageTests.detailed_assessment()
        weighted_front = assessment["consulting_front_matter"]
        for candidate in weighted_front["use_case_prioritization"]["candidates"]:
            candidate.update({
                "value": "low",
                "feasibility": "low",
                "data_readiness": "low",
                "scale": "low",
                "priority": "Watch",
            })
        for use_case_no, scale in ((1, "high"), (2, "medium"), (8, "medium")):
            weighted_front["use_case_prioritization"]["candidates"][use_case_no - 1].update({
                "value": "high",
                "feasibility": "high",
                "data_readiness": "high",
                "scale": scale,
            })
        weighted_front = generator.normalize_consulting_front_matter(
            weighted_front, assessment, {"I1"},
        )
        self.assertEqual([("1", "P1"), ("2", "P2"), ("8", "P3")], [
            (item["use_case_no"], item["priority"])
            for item in weighted_front["use_case_prioritization"]["candidates"]
            if item["priority"] != "Watch"
        ])

        services = SimpleNamespace(
            canonical_use_case_id=generator.canonical_use_case_id,
            normalized_use_case_label=generator.normalized_use_case_label,
            primary_use_case_catalog_for=generator.primary_use_case_catalog_for,
        )
        assessment, bound_front = workflow._bind_customer_priorities_to_front_matter(
            assessment,
            weighted_front,
            {
                "priority_use_cases": [
                    {"name": "ユースケース1", "priority": "高"},
                    {"name": "ユースケース2", "priority": "高"},
                    {"name": "ユースケース3", "priority": "高"},
                ],
            },
            services,
        )
        expected_candidates = [("1", "P1"), ("2", "P2"), ("3", "P3")]
        self.assertEqual(expected_candidates, [
            (item["use_case_no"], item["priority"])
            for item in bound_front["use_case_prioritization"]["candidates"]
            if item["priority"] != "Watch"
        ])

        normalized_once = generator.normalize_consulting_front_matter(
            bound_front, assessment, {"I1"},
        )
        normalized_twice = generator.normalize_consulting_front_matter(
            normalized_once, assessment, {"I1"},
        )
        for index, normalized in enumerate((normalized_once, normalized_twice), 1):
            self.assertEqual(expected_candidates, [
                (item["use_case_no"], item["priority"])
                for item in normalized["use_case_prioritization"]["candidates"]
                if item["priority"] != "Watch"
            ], f"normalize_consulting_front_matter {index}回目で顧客明示順位が変化しました")

        assessment["consulting_front_matter"] = normalized_twice
        generator.materialize_poc_portfolio(assessment, normalized_twice)
        generator.materialize_assessment_decision_contract(assessment, normalized_twice)
        frozen = generator.build_review_payload(
            assessment,
            {"midterm_plan": {"status": "disabled"}, "industry_sources": []},
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing={
                "priority_use_cases": [
                    {"name": "ユースケース1", "priority": "高"},
                    {"name": "ユースケース2", "priority": "高"},
                    {"name": "ユースケース3", "priority": "高"},
                ],
            },
            cost_estimate=copy.deepcopy(fixtures.TEST_COST_ESTIMATE),
        )
        expected_contract = [("P1", "UC01"), ("P2", "UC02"), ("P3", "UC03")]
        frozen_assessment = frozen["assessment"]
        self.assertEqual(expected_contract, [
            (item["priority"], item["use_case_id"])
            for item in frozen_assessment["poc_portfolio"]["items"]
        ])
        self.assertEqual(expected_contract, [
            (item["priority"], item["use_case_id"])
            for item in frozen_assessment["poc_charters"]["charters"]
        ])
        self.assertEqual(expected_contract, [
            (item["priority"], item["use_case_id"])
            for item in frozen_assessment["poc_selection_scorecard"]["items"]
        ])
        self.assertEqual(expected_contract, [
            (item["priority"], item["use_case_id"])
            for item in frozen_assessment["poc_priority_decision"]["items"]
            if item["priority"] in {"P1", "P2", "P3"}
        ])
        self.assertEqual(expected_contract, [
            (item["priority"], item["use_case_id"])
            for item in frozen_assessment["poc_start_readiness"]["items"]
        ])
        self.assertEqual([], generator.validate_assessment_payload(frozen, strict=True))

    def test_fixed_catalog_selection_links_list_priority_and_detail_contracts(self) -> None:
        """固定JSONのNo.1〜3が、再正規化後も一覧・詳細の同じIDへ収束する。"""
        assessment = fixtures.DynamicPptxPageTests.detailed_assessment()
        weighted_front = assessment["consulting_front_matter"]
        for candidate in weighted_front["use_case_prioritization"]["candidates"]:
            candidate.update({
                "value": "low",
                "feasibility": "low",
                "data_readiness": "low",
                "scale": "low",
                "priority": "Watch",
            })
        for use_case_no, scale in ((1, "high"), (2, "medium"), (8, "medium")):
            weighted_front["use_case_prioritization"]["candidates"][use_case_no - 1].update({
                "value": "high",
                "feasibility": "high",
                "data_readiness": "high",
                "scale": scale,
            })
        weighted_front = generator.normalize_consulting_front_matter(
            weighted_front, assessment, {"I1"},
        )
        self.assertEqual([("1", "P1"), ("2", "P2"), ("8", "P3")], [
            (item["use_case_no"], item["priority"])
            for item in weighted_front["use_case_prioritization"]["candidates"]
            if item["priority"] != "Watch"
        ])
        fixed_selection = {
            "format": "ai-assess/fixed-poc-selection-v1",
            "schema_version": "1",
            "selection_reason": "一覧の電球と後続の3枚の詳細を同じテーマへ連携する。",
            "items": [
                {
                    "priority": f"P{index}",
                    "use_case_id": f"UC{index:02d}",
                    "use_case_no": str(index),
                    "theme": f"ユースケース{index}",
                    "detail_slide_order": index,
                }
                for index in range(1, 4)
            ],
        }
        services = SimpleNamespace(
            canonical_use_case_id=generator.canonical_use_case_id,
            normalized_use_case_label=generator.normalized_use_case_label,
            primary_use_case_catalog_for=generator.primary_use_case_catalog_for,
        )
        assessment, bound_front = workflow._bind_customer_priorities_to_front_matter(
            assessment,
            weighted_front,
            {},
            services,
            fixed_poc_selection=fixed_selection,
        )
        expected = [("P1", "UC01"), ("P2", "UC02"), ("P3", "UC03")]
        normalized_once = generator.normalize_consulting_front_matter(
            bound_front, assessment, {"I1"},
        )
        normalized_twice = generator.normalize_consulting_front_matter(
            normalized_once, assessment, {"I1"},
        )
        for normalized in (normalized_once, normalized_twice):
            self.assertEqual(expected, [
                (item["priority"], item["use_case_id"])
                for item in normalized["use_case_prioritization"]["candidates"]
                if item["priority"] != "Watch"
            ])
        assessment["consulting_front_matter"] = normalized_twice
        generator.materialize_poc_portfolio(assessment, normalized_twice)
        generator.materialize_assessment_decision_contract(assessment, normalized_twice)
        self.assertEqual(expected, [
            (item["priority"], item["use_case_id"])
            for item in assessment["poc_portfolio"]["items"]
        ])
        self.assertEqual(["UC01", "UC02", "UC03"], [
            detail["use_case_id"] for detail in assessment["poc_logic_details"]
        ])


if __name__ == "__main__":
    unittest.main()
