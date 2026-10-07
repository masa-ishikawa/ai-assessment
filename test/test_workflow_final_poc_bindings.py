"""顧客指定と最終P1〜P3を定量分析まで同一identityで運ぶ回帰テスト。"""

import contextlib
import io
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import generate_assessment as generator
from ai_assess_runtime import workflow
from ai_assess_runtime.pipeline import PipelineStage, PipelineState


def _binding_services() -> SimpleNamespace:
    return SimpleNamespace(
        canonical_use_case_id=generator.canonical_use_case_id,
        normalized_use_case_label=generator.normalized_use_case_label,
        primary_use_case_catalog_for=generator.primary_use_case_catalog_for,
    )


def _generic_assessment() -> dict:
    return {
        "service_name": "Generic Operations Service",
        "use_cases": [
            {"no": index, "use_case_id": f"UC{index:02d}", "use_case": f"候補{index}"}
            for index in range(1, 16)
        ],
        # 初期LLM推薦は顧客指定の3件目と異なるfixtureにする。
        "poc_recommendations": [
            {
                "priority": f"P{rank}", "use_case_id": f"UC{case_no:02d}",
                "theme": f"候補{case_no}", "reason": "初期理由", "first_step": "初期確認",
            }
            for rank, case_no in enumerate((1, 2, 8), 1)
        ],
    }


def _generic_front_matter() -> dict:
    return {
        "use_case_prioritization": {
            "candidates": [
                {
                    "use_case_no": str(index),
                    "priority": "P1" if index == 1 else "P2" if index == 2 else "P3" if index == 8 else "Watch",
                    "rationale": f"候補{index}の選定理由",
                    "basis": "分析仮説",
                    "data_next_action": f"候補{index}の代表データを確認する",
                }
                for index in range(1, 16)
            ],
            "selection_logic": [],
        },
    }


def _fixed_catalog_selection() -> dict:
    return {
        "format": "ai-assess/fixed-poc-selection-v1",
        "schema_version": "1",
        "selection_reason": "一覧アイコンと詳細説明を同じ3テーマに固定する。",
        "items": [
            {
                "priority": f"P{index}",
                "use_case_id": f"UC{index:02d}",
                "use_case_no": str(index),
                "theme": f"候補{index}",
                "detail_slide_order": index,
            }
            for index in range(1, 4)
        ],
    }


class FinalPocBindingWorkflowTests(unittest.TestCase):
    def test_quantitative_failure_persists_and_prints_generation_audit(self) -> None:
        assessment = {
            "service_name": "Generic Operations Service",
            "business_model_role": "provider",
        }
        state = PipelineState(
            stage=PipelineStage.EVIDENCE_READY,
            analysis_context="会社名: Example株式会社\nサービス名: Generic Operations Service",
            assessment=assessment,
            research={"research_audit": []},
            midterm_plan={"status": "not_found"},
        )

        def fail_business_impact(*args, **kwargs):
            kwargs["analysis_log"].extend([
                {
                    "attempt": index,
                    "status": "repair_required",
                    "issue_count": 1,
                    "issues": [f"items[{index}]の会社KPIが不正です。"],
                    "candidate_response_sha256": f"hash-{index}",
                    "categories": [],
                    "candidate_field_summary": [],
                }
                for index in range(1, 4)
            ])
            return {}

        with tempfile.TemporaryDirectory() as temporary_directory:
            research_dir = Path(temporary_directory) / "research"
            services = SimpleNamespace(
                RESEARCH_OUTPUT_DIR=research_dir,
                QUANTITATIVE_EVIDENCE_SCHEMA_VERSION="1",
                safe_filename=generator.safe_filename,
                resolve_quantitative_display_contract=lambda *args, **kwargs: {},
                materialize_quantitative_display_contract=lambda *args, **kwargs: {
                    "status": "measurement_design", "approved_source_ids": [],
                },
                create_quantitative_analysis_client=lambda *args, **kwargs: object(),
                build_ai_product_business_impact=fail_business_impact,
            )
            args = SimpleNamespace(
                provider="oci_responses", skip_industry_research=False,
                model_id="test-model", oci_project_ocid="ocid1.test",
                oci_region="ap-tokyo-1", profile="DEFAULT", oci_config_file=None,
                oci_responses_auth_mode="api_key", openai_api_key=None,
                compartment_id="", endpoint="",
            )
            stdout = io.StringIO()
            with contextlib.redirect_stdout(stdout):
                with self.assertRaises(RuntimeError) as raised:
                    workflow._resolve_quantitative_stage(args, services, state)

            failure_path = research_dir / "Generic_Operations_Service_research.failed.json"
            self.assertTrue(failure_path.is_file())
            failure = json.loads(failure_path.read_text(encoding="utf-8"))
            self.assertEqual(
                "ai_product_business_impact",
                failure["generation_failure"]["stage"],
            )
            self.assertEqual(3, failure["generation_failure"]["attempt_count"])
            self.assertEqual(3, len(failure["ai_product_business_impact_generation_log"]))
            self.assertIn(str(failure_path), str(raised.exception))
            self.assertIn("items[3]の会社KPIが不正です。", str(raised.exception))
            self.assertIn("生成失敗監査ログ", stdout.getvalue())
            self.assertIn("生成試行3: repair_required", stdout.getvalue())

    def test_missing_customer_priority_is_restored_from_structured_input(self) -> None:
        reconciled = workflow._reconcile_customer_priorities_into_catalog(
            _generic_assessment(),
            {
                "priority_use_cases": [{
                    "name": "顧客指定の判断支援",
                    "priority": "高",
                    "business_domain": "業務支援",
                    "scenario": "担当者の判断を代表データで支援する。",
                    "ai_technology": "生成AI",
                }],
            },
            _binding_services(),
        )
        restored = reconciled["use_cases"][0]
        self.assertEqual("UC01", restored["use_case_id"])
        self.assertEqual("顧客指定の判断支援", restored["use_case"])
        self.assertEqual("担当者の判断を代表データで支援する。", restored["description"])
        self.assertEqual("UC01", reconciled["poc_recommendations"][0]["use_case_id"])
        self.assertEqual("顧客指定の判断支援", reconciled["poc_recommendations"][0]["theme"])

    def test_customer_high_priorities_replace_intermediate_weighted_selection(self) -> None:
        assessment, front = workflow._bind_customer_priorities_to_front_matter(
            _generic_assessment(),
            _generic_front_matter(),
            {
                "priority_use_cases": [
                    {"name": "候補1", "priority": "高"},
                    {"name": "候補2", "priority": "高"},
                    {"name": "候補3", "priority": "高"},
                ],
            },
            _binding_services(),
        )
        selected = [
            (item["priority"], item["use_case_no"])
            for item in front["use_case_prioritization"]["candidates"]
            if item["priority"] in {"P1", "P2", "P3"}
        ]
        self.assertEqual([("P1", "1"), ("P2", "2"), ("P3", "3")], selected)
        self.assertEqual(
            ["UC01", "UC02", "UC03"],
            [item["use_case_id"] for item in assessment["poc_recommendations"]],
        )

    def test_fixed_catalog_selection_replaces_weighted_selection_and_persists_binding(self) -> None:
        """一覧のNo.1〜3指定は、重み付きP3=UC08より優先して後段へ渡す。"""
        assessment, front = workflow._bind_customer_priorities_to_front_matter(
            _generic_assessment(),
            _generic_front_matter(),
            {
                "priority_use_cases": [
                    {"name": "今回入力の別表現の高優先テーマ", "priority": "高"},
                ],
            },
            _binding_services(),
            fixed_poc_selection=_fixed_catalog_selection(),
        )
        expected = [("P1", "1"), ("P2", "2"), ("P3", "3")]
        self.assertEqual(expected, [
            (item["priority"], item["use_case_no"])
            for item in front["use_case_prioritization"]["candidates"]
            if item["priority"] != "Watch"
        ])
        self.assertNotIn("customer_priority_binding", assessment)
        self.assertEqual(["UC01", "UC02", "UC03"], [
            item["use_case_id"]
            for item in assessment["fixed_poc_selection_binding"]["items"]
        ])
        portfolio = generator.materialize_poc_portfolio(assessment, front)
        self.assertEqual(["UC01", "UC02", "UC03"], [
            item["use_case_id"] for item in portfolio["items"]
        ])

    def test_final_bindings_are_passed_to_selector_and_executive_analysis(self) -> None:
        assessment = _generic_assessment()
        assessment["midterm_plan_analysis"] = {
            "management_targets": [
                {
                    "claim_id": "M1-01", "claim_status": "verified_external",
                    "label": "対象サービス売上高", "target": "120億円",
                    "source_id": "M1", "source_url": "https://example.com/plan",
                    "source_locator": "FY2028売上目標", "source_metric": "売上高120億円",
                    "evidence_excerpt": "対象サービス売上高120億円",
                },
                {
                    "claim_id": "M1-02", "claim_status": "verified_external",
                    "label": "営業利益", "target": "18億円",
                    "source_id": "M1", "source_url": "https://example.com/plan",
                    "source_locator": "FY2028利益目標", "source_metric": "営業利益18億円",
                    "evidence_excerpt": "営業利益18億円",
                },
            ],
        }
        state = PipelineState().advance(
            PipelineStage.INPUT_READY,
            input_text="raw input",
            analysis_context="raw input\npublic context",
            input_preprocessing={
                "priority_use_cases": [
                    {"name": f"候補{index}", "priority": "高"}
                    for index in (1, 2, 3)
                ],
            },
        )
        state = state.advance(PipelineStage.STRATEGY_READY, midterm_plan={"status": "not_found"})
        state = state.advance(PipelineStage.ASSESSMENT_READY, assessment=assessment)
        state = state.advance(
            PipelineStage.RESEARCH_READY,
            research={"research_audit": []},
            front_matter_client=object(),
            research_client=object(),
            research_sources=({"id": "R1", "url": "https://example.com/case"},),
        )
        captured: dict[str, list[dict]] = {}

        def select_sources(*args, **kwargs):
            captured["selector"] = kwargs["poc_bindings"]
            return {"approved_source_ids": ["R1"], "reason": "test"}

        def build_evidence(*args, **kwargs):
            captured["analysis"] = kwargs["poc_bindings"]
            return {
                "management_targets": [
                    assessment["midterm_plan_analysis"]["management_targets"][0],
                ],
            }

        services = SimpleNamespace(
            build_consulting_front_matter=lambda *args, **kwargs: _generic_front_matter(),
            build_final_poc_logic_details=lambda *args, **kwargs: [],
            normalize_adb_terminology=lambda value: value,
            materialize_poc_portfolio=generator.materialize_poc_portfolio,
            select_verified_quantitative_sources=select_sources,
            build_executive_evidence_analysis=build_evidence,
            create_quantitative_analysis_client=lambda *args, **kwargs: object(),
            build_ai_product_business_impact=lambda *args, **kwargs: (
                captured.setdefault("management_targets", kwargs["management_targets"])
                and {"items": [1, 2, 3]}
            ),
            materialize_ai_product_business_impact_contract=lambda *args, **kwargs: {},
            QUANTITATIVE_EVIDENCE_SCHEMA_VERSION="1",
            resolve_quantitative_display_contract=lambda *args, **kwargs: object(),
            materialize_quantitative_display_contract=lambda *args, **kwargs: {"status": "measurement_design"},
            canonical_use_case_id=generator.canonical_use_case_id,
            normalized_use_case_label=generator.normalized_use_case_label,
            primary_use_case_catalog_for=generator.primary_use_case_catalog_for,
        )
        args = SimpleNamespace(
            provider="oci_responses", skip_industry_research=False,
            model_id="test-model", oci_project_ocid="ocid1.test", oci_region="ap-tokyo-1",
            profile="DEFAULT", oci_config_file=None, oci_responses_auth_mode="api_key",
            openai_api_key=None, compartment_id="", endpoint="",
        )

        state = workflow._materialize_portfolio_stage(args, services, state)
        expected = [
            {"priority": f"P{index}", "use_case_id": f"UC{index:02d}", "theme": f"候補{index}"}
            for index in range(1, 4)
        ]
        self.assertEqual(expected, list(state.final_poc_bindings))
        state = workflow._select_quantitative_evidence_stage(args, services, state)
        state = workflow._resolve_quantitative_stage(args, services, state)
        self.assertEqual(expected, captured["selector"])
        self.assertEqual(expected, captured["analysis"])
        self.assertEqual(
            assessment["midterm_plan_analysis"]["management_targets"],
            captured["management_targets"],
        )

    def test_materialize_stage_passes_fixed_catalog_selection_to_final_bindings(self) -> None:
        assessment = _generic_assessment()
        state = PipelineState(use_case_catalog={"poc_selection": _fixed_catalog_selection()}).advance(
            PipelineStage.INPUT_READY,
            input_text="raw input",
            analysis_context="raw input\npublic context",
            input_preprocessing={},
        )
        state = state.advance(PipelineStage.STRATEGY_READY, midterm_plan={"status": "not_found"})
        state = state.advance(PipelineStage.ASSESSMENT_READY, assessment=assessment)
        state = state.advance(
            PipelineStage.RESEARCH_READY,
            research={"research_audit": []},
            front_matter_client=object(),
            research_client=object(),
            research_sources=(),
        )
        services = SimpleNamespace(
            build_consulting_front_matter=lambda *args, **kwargs: _generic_front_matter(),
            build_final_poc_logic_details=lambda *args, **kwargs: [],
            normalize_adb_terminology=lambda value: value,
            materialize_poc_portfolio=generator.materialize_poc_portfolio,
            canonical_use_case_id=generator.canonical_use_case_id,
            normalized_use_case_label=generator.normalized_use_case_label,
            primary_use_case_catalog_for=generator.primary_use_case_catalog_for,
        )
        args = SimpleNamespace(
            provider="oci_responses", model_id="test-model", oci_project_ocid="ocid1.test",
            oci_region="ap-tokyo-1", profile="DEFAULT", oci_config_file=None,
            oci_responses_auth_mode="api_key",
        )
        result = workflow._materialize_portfolio_stage(args, services, state)
        self.assertEqual(["UC01", "UC02", "UC03"], [
            item["use_case_id"] for item in result.final_poc_bindings
        ])

    def test_prompts_use_explicit_final_bindings_not_stale_recommendations(self) -> None:
        stale = {
            "poc_recommendations": [
                {"priority": f"P{index}", "use_case_id": f"UC{index + 7:02d}", "theme": f"旧候補{index}"}
                for index in range(1, 4)
            ],
        }
        final = [
            {"priority": f"P{index}", "use_case_id": f"UC{index:02d}", "theme": f"最終候補{index}"}
            for index in range(1, 4)
        ]
        source = {
            "id": "R1", "title": "公開事例", "url": "https://example.com/case",
            "excerpt": "対象業務の処理時間を20%削減した。",
        }
        with patch.object(generator, "_response_json", return_value={
            "approved_source_ids": ["R1"], "reason": "test",
        }) as selector_response:
            generator.select_verified_quantitative_sources(
                "顧客情報", stale, [source], client=object(), model_id="test-model",
                poc_bindings=final,
            )
        selector_prompt = selector_response.call_args.kwargs["prompt"]
        self.assertIn("最終候補3", selector_prompt)
        self.assertNotIn("旧候補3", selector_prompt)

        with patch.object(generator, "_response_json", return_value={}) as analysis_response:
            generator.build_executive_evidence_analysis(
                "顧客情報", stale, [source], {"status": "not_found"},
                client=object(), model_id="test-model", poc_bindings=final,
            )
        first_analysis_prompt = analysis_response.call_args_list[0].kwargs["prompt"]
        self.assertIn("最終候補3", first_analysis_prompt)
        self.assertNotIn("旧候補3", first_analysis_prompt)


if __name__ == "__main__":
    unittest.main()
