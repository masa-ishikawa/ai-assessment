"""固定ユースケースカタログ再利用の回帰テスト。"""

import contextlib
import copy
import io
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

import generate_assessment as assessment_generator
from ai_assess_runtime.assessment_contract import _reused_catalog_provenance_issues
from ai_assess_runtime.pipeline import PipelineStage, PipelineState
from ai_assess_runtime.workflow import _analyze_assessment_stage
from ai_assess_runtime.use_case_catalog import _service_names_match


def _fixed_groups(service_name: str = "固定サービス") -> list[dict]:
    return [{
        "service_name": service_name,
        "service_type": "業務管理",
        "use_cases": [
            {
                "no": index,
                "use_case_id": f"UC{index:02d}",
                "coverage_area": f"固定領域{index}",
                "use_case": f"固定ユースケース{index}",
                "ai_technology": "生成AI・RAG",
                "description": f"固定した業務データを使い、判断支援{index}を行う。",
            }
            for index in range(1, 16)
        ],
    }]


def _analysis_result(service_name: str = "固定サービス") -> dict:
    return {
        "company_name": "テスト株式会社",
        "service_name": service_name,
        "business_model_role": "provider",
        "use_cases": [
            {
                "no": index,
                "use_case_id": f"UC{index:02d}",
                "coverage_area": f"AI生成領域{index}",
                "use_case": f"AI生成ユースケース{index}",
                "ai_technology": "機械学習",
                "description": f"AI生成の説明{index}",
            }
            for index in range(1, 16)
        ],
        "service_use_case_groups": _fixed_groups("AI生成サービス"),
        "poc_recommendations": [{"theme": "過去PoC", "reason": "旧理由", "first_step": "旧手順"}],
        "poc_logic_details": [{"theme": "過去PoC"}],
        "technical_proposal": {"state": "old"},
        "executive_summary": "今回の対象サービスを分析する。",
    }


def _fixed_poc_selection() -> dict:
    return {
        "format": "ai-assess/fixed-poc-selection-v1",
        "schema_version": "1",
        "selection_reason": "合意済みの3テーマを一覧アイコンと後続詳細に連携する。",
        "items": [
            {
                "priority": f"P{index}",
                "use_case_id": f"UC{index:02d}",
                "use_case_no": str(index),
                "theme": f"固定ユースケース{index}",
                "detail_slide_order": index,
            }
            for index in range(1, 4)
        ],
    }


class UseCaseCatalogTests(unittest.TestCase):
    def test_service_name_match_ignores_trademark_marks_and_uses_catalog_formal_name(self) -> None:
        formal_name = "Logistics Station iWMS® G5"
        self.assertTrue(_service_names_match("iWMS G5", formal_name))
        self.assertFalse(_service_names_match("別製品 G5", formal_name))
        catalog = {
            "service_use_case_groups": _fixed_groups(formal_name),
        }
        result = assessment_generator.apply_use_case_catalog(
            _analysis_result("iWMS G5"), catalog,
        )
        self.assertEqual(formal_name, result["service_name"])
        self.assertEqual(formal_name, result["service_use_case_groups"][0]["service_name"])

    def _write_frozen_catalog(self, directory: Path, *, service_name: str = "固定サービス") -> Path:
        groups = _fixed_groups(service_name)
        payload = {
            "format": assessment_generator.ASSESSMENT_JSON_FROZEN_FORMAT,
            "assessment": {
                "company_name": "テスト株式会社",
                "service_name": service_name,
                "service_use_case_groups": groups,
                "use_cases": copy.deepcopy(groups[0]["use_cases"]),
                "poc_recommendations": [{"theme": "過去PoC"}],
            },
        }
        source = directory / "previous_assessment.json"
        source.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return source

    def test_loads_frozen_assessment_and_uses_only_its_catalog(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = self._write_frozen_catalog(Path(temporary_directory))
            catalog = assessment_generator.load_use_case_catalog(source)

        self.assertEqual("ai-assess/use-case-catalog-v1", catalog["format"])
        self.assertEqual("固定ユースケース2", catalog["use_cases"][1]["use_case"])
        self.assertEqual("ai-assess/assessment-v2", catalog["provenance"]["source_format"])
        self.assertEqual(1, catalog["provenance"]["service_count"])
        self.assertRegex(catalog["provenance"]["catalog_sha256"], r"^[0-9a-f]{64}$")
        self.assertIsNone(catalog["poc_selection"])

    def test_loads_explicit_lightweight_poc_selection_and_records_its_provenance(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "catalog.json"
            selection = _fixed_poc_selection()
            source.write_text(json.dumps({
                "format": "ai-assess/use-case-catalog-v1",
                "schema_version": "1",
                "service_use_case_groups": _fixed_groups(),
                "poc_selection": selection,
            }, ensure_ascii=False), encoding="utf-8")
            catalog = assessment_generator.load_use_case_catalog(source)

        self.assertEqual(selection, catalog["poc_selection"])
        self.assertEqual(selection, catalog["provenance"]["poc_selection"])
        self.assertEqual(3, catalog["provenance"]["poc_selection_count"])
        self.assertRegex(catalog["provenance"]["poc_selection_sha256"], r"^[0-9a-f]{64}$")

    def test_rejects_invalid_explicit_lightweight_poc_selection(self) -> None:
        cases = (
            ("priority", lambda value: value["items"][1].update(priority="P3")),
            ("use_case_id", lambda value: value["items"][2].update(use_case_id="UC02")),
            ("use_case_no", lambda value: value["items"][0].update(use_case_no="2")),
            ("theme", lambda value: value["items"][1].update(theme="別ユースケース")),
            ("detail_slide_order", lambda value: value["items"][2].update(detail_slide_order=2)),
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "catalog.json"
            for label, mutate in cases:
                with self.subTest(case=label):
                    selection = _fixed_poc_selection()
                    mutate(selection)
                    source.write_text(json.dumps({
                        "format": "ai-assess/use-case-catalog-v1",
                        "service_use_case_groups": _fixed_groups(),
                        "poc_selection": selection,
                    }, ensure_ascii=False), encoding="utf-8")
                    with self.assertRaises(ValueError):
                        assessment_generator.load_use_case_catalog(source)

    def test_loads_lightweight_catalog_and_rejects_inconsistent_ids(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "catalog.json"
            source.write_text(json.dumps({
                "format": "ai-assess/use-case-catalog-v1",
                "service_use_case_groups": _fixed_groups(),
            }, ensure_ascii=False), encoding="utf-8")
            catalog = assessment_generator.load_use_case_catalog(source)
            self.assertEqual("固定サービス", catalog["service_use_case_groups"][0]["service_name"])

            invalid = _fixed_groups()
            invalid[0]["use_cases"][0]["use_case_id"] = "UC15"
            source.write_text(json.dumps({
                "format": "ai-assess/use-case-catalog-v1",
                "service_use_case_groups": invalid,
            }, ensure_ascii=False), encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "use_case_id"):
                assessment_generator.load_use_case_catalog(source)

    def test_analysis_stage_replaces_llm_catalog_and_discards_old_poc_values(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            catalog = assessment_generator.load_use_case_catalog(
                self._write_frozen_catalog(Path(temporary_directory))
            )
        services = SimpleNamespace(
            analyze_assessment=lambda *_args, **_kwargs: _analysis_result(),
            normalize_adb_terminology=lambda value: value,
            derive_display_company_name=lambda _text, fallback: fallback,
            derive_display_service_name=lambda _text, fallback: fallback,
            extract_service_genre=lambda _text, fallback: fallback,
            apply_use_case_catalog=assessment_generator.apply_use_case_catalog,
            primary_use_case_catalog_for=assessment_generator.primary_use_case_catalog_for,
            normalized_use_case_label=assessment_generator.normalized_use_case_label,
        )
        args = SimpleNamespace(
            provider="test", model_id="test-model", openai_api_key=None, compartment_id="",
            profile="", endpoint="", oci_config_file="", oci_project_ocid="", oci_region="",
            oci_responses_auth_mode="auto", company_name=None,
        )
        state = PipelineState(
            stage=PipelineStage.STRATEGY_READY,
            analysis_context="会社名: テスト株式会社\nサービス名: 固定サービス",
            use_case_catalog=catalog,
        )
        result = _analyze_assessment_stage(args, services, state)

        self.assertEqual(PipelineStage.ASSESSMENT_READY, result.stage)
        self.assertEqual("固定ユースケース1", result.assessment["use_cases"][0]["use_case"])
        self.assertEqual("固定ユースケース1", result.assessment["service_use_case_groups"][0]["use_cases"][0]["use_case"])
        self.assertEqual([], result.assessment["poc_recommendations"])
        self.assertNotIn("poc_logic_details", result.assessment)
        self.assertNotIn("technical_proposal", result.assessment)

    def test_fixed_catalog_does_not_replace_missing_customer_priority(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            catalog = assessment_generator.load_use_case_catalog(
                self._write_frozen_catalog(Path(temporary_directory))
            )
        services = SimpleNamespace(
            analyze_assessment=lambda *_args, **_kwargs: _analysis_result(),
            normalize_adb_terminology=lambda value: value,
            derive_display_company_name=lambda _text, fallback: fallback,
            derive_display_service_name=lambda _text, fallback: fallback,
            extract_service_genre=lambda _text, fallback: fallback,
            apply_use_case_catalog=assessment_generator.apply_use_case_catalog,
            primary_use_case_catalog_for=assessment_generator.primary_use_case_catalog_for,
            normalized_use_case_label=assessment_generator.normalized_use_case_label,
        )
        args = SimpleNamespace(
            provider="test", model_id="test-model", openai_api_key=None, compartment_id="",
            profile="", endpoint="", oci_config_file="", oci_project_ocid="", oci_region="",
            oci_responses_auth_mode="auto", company_name=None,
        )
        state = PipelineState(
            stage=PipelineStage.STRATEGY_READY,
            analysis_context="サービス名: 固定サービス",
            input_preprocessing={
                "priority_use_cases": [{"name": "カタログ外の高優先テーマ", "priority": "高"}],
            },
            use_case_catalog=catalog,
        )
        with self.assertRaisesRegex(RuntimeError, "カタログ外の高優先テーマ"):
            _analyze_assessment_stage(args, services, state)

    def test_catalog_provenance_requires_the_final_groups_to_stay_fixed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            catalog = assessment_generator.load_use_case_catalog(
                self._write_frozen_catalog(Path(temporary_directory))
            )
        assessment = assessment_generator.apply_use_case_catalog(_analysis_result(), catalog)
        input_record = {"preprocessing": {"use_case_catalog_reuse": catalog["provenance"]}}
        self.assertEqual([], _reused_catalog_provenance_issues(input_record, assessment))
        assessment["service_use_case_groups"][0]["use_cases"][0]["description"] = "途中で変更された説明"
        self.assertTrue(_reused_catalog_provenance_issues(input_record, assessment))

    def test_catalog_selection_provenance_requires_the_final_poc_portfolio_to_stay_fixed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            source = Path(temporary_directory) / "catalog.json"
            selection = _fixed_poc_selection()
            source.write_text(json.dumps({
                "format": "ai-assess/use-case-catalog-v1",
                "service_use_case_groups": _fixed_groups(),
                "poc_selection": selection,
            }, ensure_ascii=False), encoding="utf-8")
            catalog = assessment_generator.load_use_case_catalog(source)
        assessment = assessment_generator.apply_use_case_catalog(_analysis_result(), catalog)
        assessment["fixed_poc_selection_binding"] = {
            "schema_version": "1",
            "source": "fixed_use_case_catalog",
            "selection_reason": selection["selection_reason"],
            "items": copy.deepcopy(selection["items"]),
        }
        assessment["poc_portfolio"] = {
            "schema_version": "1",
            "items": [
                {
                    **copy.deepcopy(item),
                    "reason": "代表データで業務価値を検証する。",
                    "first_step": "対象データを確認する。",
                    "depends_on": "対象データと責任者を確認する。",
                    "basis": "分析仮説",
                    "architecture_implementation": "",
                }
                for item in selection["items"]
            ],
        }
        input_record = {"preprocessing": {"use_case_catalog_reuse": catalog["provenance"]}}
        self.assertEqual([], _reused_catalog_provenance_issues(input_record, assessment))
        assessment["poc_portfolio"]["items"][2]["use_case_id"] = "UC04"
        self.assertTrue(_reused_catalog_provenance_issues(input_record, assessment))

    def test_cli_rejects_catalog_reuse_with_from_json(self) -> None:
        with patch.object(sys, "argv", [
            "generate_assessment.py", "--from-json", "previous.json",
            "--use-case-catalog-json", "catalog.json",
        ]):
            with contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as raised:
                    assessment_generator.main()
        self.assertEqual(2, raised.exception.code)


if __name__ == "__main__":
    unittest.main()
