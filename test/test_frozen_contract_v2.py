"""凍結JSON v2の入力・根拠・描画再現性を固定する回帰テスト。"""

import copy
import hashlib
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import Mock

import generate_assessment as generator
from ai_assess_runtime.from_json import run_from_json
from ai_assess_runtime import quantitative_contract
from ai_assess_runtime.review_snapshot import cost_estimate_from_snapshot
from test import test_generate_assessment as fixtures


class FrozenContractV2Tests(unittest.TestCase):
    @staticmethod
    def _cost_snapshot_source() -> dict:
        return copy.deepcopy(fixtures.TEST_COST_ESTIMATE)

    @classmethod
    def _valid_payload(cls, *, preprocessing: object = None) -> dict:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        generator.materialize_assessment_decision_contract(assessment)
        return generator.build_review_payload(
            assessment,
            {"midterm_plan": {"status": "disabled"}, "industry_sources": []},
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=preprocessing,
            cost_estimate=cls._cost_snapshot_source(),
        )

    @staticmethod
    def _rehash_v2(payload: dict) -> None:
        research = payload["research"]
        research.pop("provenance", None)
        provenance = payload["provenance"]
        provenance["input_record_sha256"] = generator.canonical_sha256(payload["input"])
        provenance["assessment_sha256"] = generator.canonical_sha256(payload["assessment"])
        provenance["research_content_sha256"] = generator.canonical_sha256(research)
        provenance["rendering_sha256"] = generator.canonical_sha256(payload["rendering"])
        provenance["assessment_run_id"] = hashlib.sha256(
            ":".join((
                provenance["input_record_sha256"],
                provenance["assessment_sha256"],
                provenance["research_content_sha256"],
                provenance["generator_sha256"],
                provenance["rendering_sha256"],
            )).encode("utf-8")
        ).hexdigest()[:24]
        research["provenance"] = {
            "format": provenance["format"],
            "assessment_run_id": provenance["assessment_run_id"],
            "source_text_sha256": provenance["source_text_sha256"],
            "input_record_sha256": provenance["input_record_sha256"],
            "assessment_sha256": provenance["assessment_sha256"],
            "research_content_sha256": provenance["research_content_sha256"],
            "rendering_sha256": provenance["rendering_sha256"],
        }

    def test_strict_rejects_hash_consistent_minimal_fake_assessment(self) -> None:
        payload = generator.build_review_payload(
            {
                "company_name": "Example株式会社",
                "service_name": "Example Service",
                "business_model_role": "provider",
                "use_cases": [{"no": 1, "use_case": "問い合わせ支援"}],
                "poc_recommendations": [{"use_case_id": "UC01", "theme": "問い合わせ支援"}],
            },
            {"midterm_plan": {"status": "disabled"}, "industry_sources": []},
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )

        errors = generator.validate_assessment_payload(payload, strict=True)

        self.assertTrue(any("現行標準デッキの凍結契約" in error for error in errors))
        self.assertTrue(any("use_casesを15件" in error for error in errors))
        self.assertTrue(any("poc_recommendationsを3件" in error for error in errors))

    def test_new_contract_binds_the_complete_input_record(self) -> None:
        payload = self._valid_payload(preprocessing={"mode": "xlsx", "answer_count": 10})
        self.assertEqual("ai-assess/reproducibility-v2", payload["provenance"]["format"])
        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))

        tampered = copy.deepcopy(payload)
        tampered["input"]["preprocessing"]["answer_count"] = 11
        errors = generator.validate_assessment_payload(tampered, strict=True)

        self.assertTrue(any("input_record_sha256" in error for error in errors))
        self.assertTrue(any("assessment_run_id" in error for error in errors))

    def test_generator_fingerprint_must_be_a_sha256(self) -> None:
        payload = self._valid_payload()
        payload["provenance"]["generator_sha256"] = "self-declared"

        errors = generator.validate_assessment_payload(payload, strict=True)

        self.assertTrue(any("64桁のSHA-256" in error for error in errors))

    def test_existing_reproducibility_v1_payload_remains_readable(self) -> None:
        payload = self._valid_payload()
        provenance = payload["provenance"]
        provenance["format"] = "ai-assess/reproducibility-v1"
        provenance.pop("input_record_sha256")
        provenance.pop("rendering_sha256")
        provenance["assessment_run_id"] = hashlib.sha256(
            (
                f"{provenance['source_text_sha256']}:"
                f"{provenance['assessment_sha256']}:"
                f"{provenance['research_content_sha256']}:"
                f"{provenance['generator_sha256']}"
            ).encode("utf-8")
        ).hexdigest()[:24]
        research_provenance = payload["research"]["provenance"]
        research_provenance["format"] = "ai-assess/reproducibility-v1"
        research_provenance["assessment_run_id"] = provenance["assessment_run_id"]
        research_provenance.pop("input_record_sha256")
        research_provenance.pop("rendering_sha256")

        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))

    def test_reproducibility_v1_json_only_upgrade_preserves_input_and_cost(self) -> None:
        payload = self._valid_payload(preprocessing={"mode": "xlsx", "answer_count": 12})
        payload["input"].update({
            "source_file": "/approved/archive/customer-input.xlsx",
            "source_file_sha256": "a" * 64,
            "import_contract": {"schema_version": "1", "sheet": "Assessment"},
        })
        original_input = copy.deepcopy(payload["input"])
        original_cost = copy.deepcopy(payload["rendering"]["cost_estimate"])
        original_profile = copy.deepcopy(payload["rendering"]["render_profile"])

        provenance = payload["provenance"]
        provenance["format"] = "ai-assess/reproducibility-v1"
        provenance.pop("input_record_sha256")
        provenance.pop("rendering_sha256")
        provenance["assessment_run_id"] = hashlib.sha256(
            (
                f"{provenance['source_text_sha256']}:"
                f"{provenance['assessment_sha256']}:"
                f"{provenance['research_content_sha256']}:"
                f"{provenance['generator_sha256']}"
            ).encode("utf-8")
        ).hexdigest()[:24]
        research_provenance = payload["research"]["provenance"]
        research_provenance["format"] = "ai-assess/reproducibility-v1"
        research_provenance["assessment_run_id"] = provenance["assessment_run_id"]
        research_provenance.pop("input_record_sha256")
        research_provenance.pop("rendering_sha256")
        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "review-v1.json"
            output = root / "review-v2.json"
            source.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            create_pptx = Mock(side_effect=AssertionError("JSON移行でPPTXを生成してはいけません"))
            api = self._api(root, payload, create_pptx)
            api.build_review_payload = generator.build_review_payload
            api.validate_assessment_payload = generator.validate_assessment_payload
            args = SimpleNamespace(
                from_json=source,
                strict_json=True,
                include_source_text=False,
                json_only=True,
                output_json=output,
                architecture_image=None,
                output=None,
            )

            self.assertEqual(0, run_from_json(
                args,
                SimpleNamespace(error=lambda message: self.fail(message)),
                api,
            ))
            upgraded = json.loads(output.read_text(encoding="utf-8"))
            architecture_bytes = api.DEFAULT_ARCHITECTURE_IMAGE.read_bytes()

        self.assertEqual("ai-assess/reproducibility-v2", upgraded["provenance"]["format"])
        self.assertEqual(original_input, upgraded["input"])
        self.assertEqual(original_cost, upgraded["rendering"]["cost_estimate"])
        self.assertEqual(original_profile, upgraded["rendering"]["render_profile"])
        self.assertEqual(
            hashlib.sha256(architecture_bytes).hexdigest(),
            upgraded["rendering"]["architecture_asset"]["sha256"],
        )
        self.assertEqual([], generator.validate_assessment_payload(upgraded, strict=True))
        api.build_poc_cost_estimate.assert_not_called()
        create_pptx.assert_not_called()

    def test_search_snippet_alone_never_verifies_a_numeric_claim(self) -> None:
        claim = {
            "claim_id": "Q01",
            "claim_status": "verified_external",
            "source_id": "R1",
            "source_url": "https://example.com/case",
            "source_locator": "検索結果",
            "source_metric": "処理時間を50%削減",
            "evidence_excerpt": "処理時間を50%削減",
            "headline_metric": "50%",
        }
        snippet_source = {
            "R1": {
                "id": "R1",
                "url": "https://example.com/case",
                "snippet": "処理時間を50%削減",
                "excerpt": "処理時間を50%削減",
                "fetch_status": "search_snippet",
            },
        }
        fetched_source = copy.deepcopy(snippet_source)
        fetched_source["R1"]["fetch_status"] = "fetched"

        self.assertFalse(quantitative_contract._quantitative_claim_is_verifiable(
            claim, snippet_source, {"R1"},
        ))
        self.assertTrue(quantitative_contract._quantitative_claim_is_verifiable(
            claim, fetched_source, {"R1"},
        ))

        # repro-v2ではfetch_statusだけを自己申告しても不十分で、数値を
        # 実際に取得した原文本文(excerpt)へ照合できなければ採用しない。
        fetched_without_body = copy.deepcopy(fetched_source)
        fetched_without_body["R1"]["excerpt"] = ""
        self.assertFalse(quantitative_contract._quantitative_claim_is_verifiable(
            claim, fetched_without_body, {"R1"}, require_fetched=True,
        ))

    def test_strict_v2_rejects_unstructured_business_effect_numbers(self) -> None:
        payload = self._valid_payload()
        payload["assessment"]["business_value"]["impact_areas"][0][
            "expected_impact"
        ] = "AI実装により業務生産性を20%向上"
        payload["assessment"]["midterm_plan_analysis"] = {
            "plan_summary": "AIで売上を30%向上",
            "ai_alignment": [],
            "management_targets": [],
        }
        payload["assessment"]["use_cases"][0]["description"] = "工数を25%削減"
        payload["assessment"]["poc_logic_details"][0][
            "validation_plan"
        ] = "正答率95%、工数20%削減"
        self._rehash_v2(payload)

        errors = generator.validate_assessment_payload(payload, strict=True)

        for path in (
            "assessment.business_value.impact_areas[0].expected_impact",
            "assessment.midterm_plan_analysis.plan_summary",
            "assessment.use_cases[0].description",
            "assessment.poc_logic_details[0].validation_plan",
        ):
            self.assertTrue(
                any(path in error and "fetch_status=fetched" in error for error in errors),
                msg=f"未根拠の定量表示が拒否されていません: {path}\n{errors}",
            )

    def test_business_effect_detector_covers_common_non_percent_notation(self) -> None:
        for text in (
            "工数を3割削減",
            "工数を30パーセント削減",
            "処理能力を1.5x向上",
            "処理性能を100TPSへ向上",
            "処理能力を120req/sへ向上",
            "生産性を50件/時へ改善",
        ):
            self.assertTrue(
                quantitative_contract._is_business_effect_metric(text),
                msg=f"未検出の定量効果表記: {text}",
            )

    def test_new_v2_freeze_omits_legacy_llm_quantitative_estimates(self) -> None:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        research = {"midterm_plan": {"status": "disabled"}, "industry_sources": []}
        outcome = generator.apply_quantitative_research_outcome(
            assessment,
            research,
            {},
            llm_estimate=fixtures.DynamicPptxPageTests.llm_quantitative_estimate_fixture(),
            llm_model_id="legacy-test-model",
            llm_generation_log=[{"attempt": 1, "status": "accepted"}],
            missing_evidence=["原典取得済みの公開根拠が不足"],
        )
        self.assertEqual("llm_estimate", outcome["status"])
        self.assertEqual("decision_thresholds", assessment["poc_measurement_design"]["status"])
        generator.materialize_assessment_decision_contract(assessment)

        payload = generator.build_review_payload(
            assessment,
            research,
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )

        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))
        self.assertEqual("pre_poc", payload["assessment"]["poc_measurement_design"]["status"])
        self.assertEqual(
            "measurement_design_only",
            payload["research"]["quantitative_display_contract"]["display_mode"],
        )
        self.assertFalse(
            payload["research"]["quantitative_display_contract"]["numeric_claims_allowed"]
        )
        self.assertNotIn("llm_quantitative_estimate", payload["research"])
        serialized_assessment = json.dumps(payload["assessment"], ensure_ascii=False)
        for metric in ("7.3〜11.8%", "18.6〜23.4%", "12.2〜16.9%"):
            self.assertNotIn(metric, serialized_assessment)

    def test_new_v2_preserves_audited_company_planning_estimates(self) -> None:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        generator.materialize_assessment_decision_contract(assessment)
        model = fixtures.DynamicPptxPageTests.ai_product_business_impact_fixture(
            assessment,
        )
        research = {"midterm_plan": {"status": "disabled"}, "industry_sources": []}
        quantitative_contract.materialize_ai_product_business_impact_contract(
            assessment,
            research,
            model,
            model_id="planning-model-v1",
            analysis_log=[{"attempt": 1, "status": "accepted"}],
            generation_reason=(
                "外部事例とは区別し、対象企業・製品全体のAI導入効果を"
                "売上・利益・継続収益の計画試算として整理するため。"
            ),
        )

        payload = generator.build_review_payload(
            assessment,
            research,
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )

        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))
        self.assertEqual(model, payload["assessment"]["ai_product_business_impact"])
        audit = payload["research"]["ai_product_business_impact_audit"]
        self.assertEqual("llm_estimate", audit["evidence_mode"])
        self.assertEqual("planning_estimate", audit["estimate_classification"])
        self.assertEqual("AI導入効果目標", audit["display_label"])
        self.assertEqual("planning-model-v1", audit["model_id"])
        self.assertEqual(3, len(audit["calculation_assumptions"]))
        self.assertEqual(0, audit["external_evidence_count"])
        self.assertEqual([], audit["external_source_ids"])
        self.assertFalse(audit["actual_result_claimed"])

    def test_strict_v2_rejects_tampered_company_planning_estimate_audit(self) -> None:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        generator.materialize_assessment_decision_contract(assessment)
        model = fixtures.DynamicPptxPageTests.ai_product_business_impact_fixture(
            assessment,
        )
        research = {"midterm_plan": {"status": "disabled"}, "industry_sources": []}
        quantitative_contract.materialize_ai_product_business_impact_contract(
            assessment,
            research,
            model,
            model_id="planning-model-v1",
            generation_reason="対象企業・製品全体の計画試算を作成するため。",
        )
        payload = generator.build_review_payload(
            assessment,
            research,
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )
        payload["research"]["ai_product_business_impact_audit"][
            "calculation_assumptions"
        ][0]["formula"] = "監査外の算定式"
        self._rehash_v2(payload)

        errors = generator.validate_assessment_payload(payload, strict=True)

        self.assertTrue(any(
            "evidence_mode=llm_estimate" in error for error in errors
        ), errors)

    def test_v1_company_planning_estimate_audit_is_upgraded_without_changing_ranges(self) -> None:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        generator.materialize_assessment_decision_contract(assessment)
        model = fixtures.DynamicPptxPageTests.ai_product_business_impact_fixture(
            assessment,
        )
        items = model["items"]
        assessment["ai_product_business_impact"] = copy.deepcopy(model)
        research = {
            "midterm_plan": {"status": "disabled"},
            "industry_sources": [],
            "ai_product_business_impact": copy.deepcopy(model),
            "ai_product_business_impact_audit": {
                "schema_version": generator.AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION,
                "status": "decision_thresholds",
                "generation_source": "llm_decision_threshold",
                "display_contract": "assessed_company_ai_business_impact",
                "model_id": "legacy-planning-model",
                "prompt_contract_version": generator.AI_PRODUCT_BUSINESS_IMPACT_SCHEMA_VERSION,
                "estimate_count": 3,
                "estimate_ids": [item["estimate_id"] for item in items],
                "business_kpi_ids": [item["business_kpi_id"] for item in items],
                "categories": [item["category"] for item in items],
                "displayed_targets": [item["headline_metric"] for item in items],
                "formulas": [item["formula"] for item in items],
                "assessed_company_actuals_assumed": False,
                "customer_poc_metrics_reused": False,
                "reason": "会社KPIの事業化判断目標として記録する。",
            },
        }
        payload = generator.build_review_payload(
            assessment,
            research,
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )

        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))
        self.assertEqual(
            [item["headline_metric"] for item in items],
            [
                item["headline_metric"]
                for item in payload["assessment"]["ai_product_business_impact"]["items"]
            ],
        )
        audit = payload["research"]["ai_product_business_impact_audit"]
        self.assertEqual("llm_estimate", audit["evidence_mode"])
        self.assertEqual("planning_estimate", audit["estimate_classification"])

    def test_strict_v2_rejects_hash_consistent_llm_numeric_contract(self) -> None:
        payload = self._valid_payload()
        legacy_assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        legacy_research = {"midterm_plan": {"status": "disabled"}, "industry_sources": []}
        generator.apply_quantitative_research_outcome(
            legacy_assessment,
            legacy_research,
            {},
            llm_estimate=fixtures.DynamicPptxPageTests.llm_quantitative_estimate_fixture(),
            llm_model_id="legacy-test-model",
            llm_generation_log=[{"attempt": 1, "status": "accepted"}],
            missing_evidence=["原典取得済みの公開根拠が不足"],
        )
        payload["assessment"]["poc_measurement_design"] = copy.deepcopy(
            legacy_assessment["poc_measurement_design"]
        )
        for field in (
            "poc_measurement_design", "poc_measurement_design_audit",
            "llm_quantitative_estimate", "quantitative_analysis",
        ):
            payload["research"][field] = copy.deepcopy(legacy_research[field])
        payload["research"].pop("quantitative_display_contract", None)
        self._rehash_v2(payload)

        errors = generator.validate_assessment_payload(payload, strict=True)

        self.assertTrue(any("decision_thresholds" in error for error in errors))
        self.assertTrue(any("quantitative_display_contract" in error for error in errors))
        self.assertTrue(any("LLM独自定量試算" in error for error in errors))

    def test_legacy_migration_drops_unfetched_midterm_numbers(self) -> None:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        plan = {
            "status": "found",
            "title": "中期経営計画(2026-2027)",
            "url": "https://example.com/plan",
            "excerpt": "重要経営指標 2028年 売上高 100億円を目標とする。",
        }
        targets = generator.extract_verified_management_targets_from_midterm_plan(plan)
        self.assertTrue(targets)
        assessment["midterm_plan_analysis"] = {
            "plan_summary": "成長方針",
            "ai_alignment": [{
                "plan_priority": "成長",
                "ai_role": "判断支援",
                "why_now": "重点施策",
                "related_use_case": "ユースケース1",
            }],
            "management_targets": targets,
        }
        generator.materialize_assessment_decision_contract(assessment)
        payload = generator.build_review_payload(
            assessment,
            {"midterm_plan": plan, "industry_sources": []},
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )

        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))
        self.assertEqual(
            [], payload["assessment"]["midterm_plan_analysis"]["management_targets"],
        )
        self.assertIn(
            "assessment.midterm_plan_analysis.management_targets",
            payload["research"]["legacy_quantitative_migration_audit"]["omitted_fields"],
        )

    def test_fetched_midterm_numbers_remain_traceable_in_new_v2(self) -> None:
        assessment = copy.deepcopy(fixtures.DynamicPptxPageTests.detailed_assessment())
        plan = {
            "status": "found",
            "title": "中期経営計画(2026-2027)",
            "url": "https://example.com/plan",
            "excerpt": "重要経営指標 2028年 売上高 100億円を目標とする。",
            "fetch_status": "fetched",
        }
        targets = generator.extract_verified_management_targets_from_midterm_plan(plan)
        assessment["midterm_plan_analysis"] = {
            "plan_summary": "2028年の売上高100億円目標へ向けた成長方針",
            "ai_alignment": [{
                "plan_priority": "成長",
                "ai_role": "判断支援",
                "why_now": "重点施策",
                "related_use_case": "ユースケース1",
            }],
            "management_targets": targets,
        }
        generator.materialize_assessment_decision_contract(assessment)
        payload = generator.build_review_payload(
            assessment,
            {"midterm_plan": plan, "industry_sources": []},
            source_file=None,
            source_text="会社名: Example株式会社\nサービス名: Example Service",
            preprocessing=None,
            cost_estimate=self._cost_snapshot_source(),
        )

        self.assertEqual([], generator.validate_assessment_payload(payload, strict=True))
        self.assertEqual(
            targets, payload["assessment"]["midterm_plan_analysis"]["management_targets"],
        )

    def test_fetched_plan_text_without_structured_target_is_not_enough(self) -> None:
        payload = self._valid_payload()
        payload["research"]["midterm_plan"] = {
            "status": "found",
            "title": "中期経営計画",
            "url": "https://example.com/plan",
            "excerpt": "2028年 売上高 100億円を目標とする。",
            "fetch_status": "fetched",
        }
        payload["assessment"]["midterm_plan_analysis"] = {
            "plan_summary": "2028年に売上高100億円を目指す",
            "ai_alignment": [],
            "management_targets": [],
        }
        self._rehash_v2(payload)

        errors = generator.validate_assessment_payload(payload, strict=True)

        self.assertTrue(any(
            "assessment.midterm_plan_analysis.plan_summary" in error
            and "fetch_status=fetched" in error
            for error in errors
        ))

    @staticmethod
    def _api(root: Path, payload: dict, create_pptx) -> SimpleNamespace:
        architecture = root / "architecture.png"
        if not architecture.exists():
            architecture.write_bytes(b"approved-architecture")
        return SimpleNamespace(
            ASSESSMENT_JSON_FROZEN_FORMAT=generator.ASSESSMENT_JSON_FROZEN_FORMAT,
            DEFAULT_ARCHITECTURE_IMAGE=architecture,
            JSON_OUTPUT_DIR=root / "json",
            PPTX_OUTPUT_DIR=root / "pptx",
            RESEARCH_OUTPUT_DIR=root / "research",
            build_poc_cost_estimate=Mock(side_effect=AssertionError("再見積り不可")),
            build_review_payload=Mock(),
            cost_estimate_from_snapshot=cost_estimate_from_snapshot,
            create_pptx=create_pptx,
            load_assessment_payload=Mock(return_value=payload),
            migrate_legacy_payload_for_review=Mock(),
            safe_filename=generator.safe_filename,
            validate_assessment_payload=Mock(return_value=[]),
            write_render_manifest=Mock(return_value=root / "manifest.json"),
        )

    @staticmethod
    def _args(source: Path, output: Path, architecture: Path | None = None) -> SimpleNamespace:
        return SimpleNamespace(
            from_json=source,
            strict_json=True,
            include_source_text=False,
            json_only=False,
            output_json=None,
            architecture_image=architecture,
            output=output,
        )

    def test_from_json_renderer_mutation_does_not_commit_artifact(self) -> None:
        payload = self._valid_payload()
        # v1にはasset SHA契約がなかったため、既存JSON再生の互換経路として扱う。
        payload["provenance"]["format"] = "ai-assess/reproducibility-v1"
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "review.json"
            output = root / "assessment.pptx"
            source.write_text(json.dumps({"format": generator.ASSESSMENT_JSON_FROZEN_FORMAT}), encoding="utf-8")
            output.write_bytes(b"previous-approved-artifact")

            def mutating_renderer(assessment, _architecture, temporary, _cost, _research=None) -> None:
                temporary.write_bytes(b"mutated-artifact")
                assessment["company_name"] = "描画中に改ざん"

            api = self._api(root, payload, mutating_renderer)
            with self.assertRaisesRegex(RuntimeError, "承認済みassessmentを変更"):
                run_from_json(
                    self._args(source, output),
                    SimpleNamespace(error=lambda message: self.fail(message)),
                    api,
                )

            self.assertEqual(b"previous-approved-artifact", output.read_bytes())
            api.write_render_manifest.assert_not_called()
            api.build_review_payload.assert_not_called()

    def test_from_json_rejects_architecture_asset_hash_mismatch(self) -> None:
        payload = self._valid_payload()
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "review.json"
            output = root / "assessment.pptx"
            architecture = root / "different.png"
            source.write_text(json.dumps({"format": generator.ASSESSMENT_JSON_FROZEN_FORMAT}), encoding="utf-8")
            architecture.write_bytes(b"different-architecture")
            create_pptx = Mock()
            api = self._api(root, payload, create_pptx)

            with self.assertRaisesRegex(ValueError, "構成図画像.*SHA-256"):
                run_from_json(
                    self._args(source, output, architecture),
                    SimpleNamespace(error=lambda message: self.fail(message)),
                    api,
                )

            create_pptx.assert_not_called()
            self.assertFalse(output.exists())

    def test_manifest_records_current_renderer_fingerprint(self) -> None:
        payload = self._valid_payload()
        with tempfile.TemporaryDirectory() as temporary_directory:
            artifact = Path(temporary_directory) / "assessment.pptx"
            artifact.write_bytes(b"pptx")
            manifest_path = generator.write_render_manifest(
                artifact, payload, artifact_format="pptx", design="default",
            )
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

        current = manifest["renderer"]["current_generator_sha256"]
        self.assertRegex(current, r"^[0-9a-f]{64}$")
        self.assertEqual(payload["rendering"]["render_profile"], manifest["renderer"]["render_profile"])
        self.assertEqual(payload["rendering"]["architecture_asset"], manifest["renderer"]["architecture_asset"])


if __name__ == "__main__":
    unittest.main()
