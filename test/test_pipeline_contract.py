"""通常生成の段階順と依存注入を固定する単体テスト。"""

from pathlib import Path
from types import SimpleNamespace
import unittest

from ai_assess_runtime import workflow
from ai_assess_runtime.pipeline import GenerationServices, PipelineStage, PipelineState


class PipelineContractTests(unittest.TestCase):
    def test_pipeline_stage_is_strictly_forward_only(self) -> None:
        state = PipelineState()
        state = state.advance(PipelineStage.INPUT_READY, input_file=Path("input.xlsx"))
        self.assertEqual(PipelineStage.INPUT_READY, state.stage)
        with self.assertRaisesRegex(RuntimeError, "生成段階の遷移が不正"):
            state.advance(PipelineStage.ASSESSMENT_READY)

    def test_pipeline_state_keeps_previous_envelope_unchanged(self) -> None:
        initial = PipelineState()
        advanced = initial.advance(PipelineStage.INPUT_READY, source_text="input")
        self.assertEqual(PipelineStage.INITIAL, initial.stage)
        self.assertEqual("", initial.source_text)
        self.assertEqual(PipelineStage.INPUT_READY, advanced.stage)
        self.assertEqual("input", advanced.source_text)
        self.assertEqual("input", advanced.input_text)
        self.assertEqual("input", advanced.analysis_context)

    def test_input_text_is_immutable_after_input_stage(self) -> None:
        state = PipelineState().advance(
            PipelineStage.INPUT_READY,
            input_text="raw input",
            analysis_context="raw input",
        )
        with self.assertRaisesRegex(RuntimeError, "input_textは読込完了後に変更できません"):
            state.advance(
                PipelineStage.STRATEGY_READY,
                input_text="mutated input",
                analysis_context="raw input\npublic context",
            )

    def test_quantitative_source_selection_uses_final_poc_portfolio(self) -> None:
        initial_items = [
            {"priority": f"P{index}", "use_case_id": f"UC{index:02d}", "theme": f"初期候補{index}"}
            for index in range(1, 4)
        ]
        final_items = [
            {"priority": f"P{index}", "use_case_id": f"UC{index + 3:02d}", "theme": f"最終候補{index}"}
            for index in range(1, 4)
        ]
        state = PipelineState().advance(
            PipelineStage.INPUT_READY,
            input_text="raw input",
            analysis_context="raw input\n公開中計",
        )
        state = state.advance(PipelineStage.STRATEGY_READY, midterm_plan={"status": "found"})
        state = state.advance(
            PipelineStage.ASSESSMENT_READY,
            assessment={"poc_recommendations": initial_items},
        )
        captured: list[dict] = []

        def select_sources(context, assessment, sources, **kwargs):
            captured.append({
                "context": context,
                "items": assessment["poc_portfolio"]["items"],
            })
            return {"approved_source_ids": ["R1"], "reason": "test"}

        services = SimpleNamespace(
            create_oci_responses_client=lambda **kwargs: object(),
            collect_industry_research=lambda *args, **kwargs: [
                {"id": "R1", "url": "https://example.com/evidence"},
            ],
            build_midterm_plan_analysis=lambda *args, **kwargs: {},
            build_consulting_front_matter=lambda *args, **kwargs: {},
            build_final_poc_logic_details=lambda *args, **kwargs: [],
            normalize_adb_terminology=lambda value: value,
            materialize_poc_portfolio=lambda assessment, front: {
                "schema_version": "1", "items": final_items,
            },
            select_verified_quantitative_sources=select_sources,
        )
        args = SimpleNamespace(
            provider="oci_responses",
            skip_industry_research=False,
            research_max_rounds=1,
            model_id="test-model",
            oci_project_ocid="ocid1.generativeaiproject.oc1..test",
            oci_region="ap-tokyo-1",
            profile="DEFAULT",
            oci_config_file=None,
            oci_responses_auth_mode="api_key",
        )

        state = workflow._collect_research_stage(args, services, state)
        self.assertEqual(PipelineStage.RESEARCH_READY, state.stage)
        self.assertEqual([], captured, "収集段階では採用判定してはいけません")
        state = workflow._materialize_portfolio_stage(args, services, state)
        self.assertEqual(PipelineStage.PORTFOLIO_READY, state.stage)
        self.assertEqual([], captured, "ポートフォリオ確定前には採用判定してはいけません")
        state = workflow._select_quantitative_evidence_stage(args, services, state)

        self.assertEqual(PipelineStage.EVIDENCE_READY, state.stage)
        self.assertEqual(1, len(captured))
        self.assertEqual(final_items, captured[0]["items"])
        self.assertNotEqual(initial_items, captured[0]["items"])
        self.assertEqual("raw input\n公開中計", captured[0]["context"])
        self.assertEqual(({"id": "R1", "url": "https://example.com/evidence"},), state.executive_sources)

    def test_generation_services_fails_fast_when_dependency_is_missing(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "通常生成に必要な依存"):
            GenerationServices.from_api(SimpleNamespace())


if __name__ == "__main__":
    unittest.main()
