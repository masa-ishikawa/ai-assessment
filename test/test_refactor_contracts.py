"""モジュール分割後も入口・出口・再現性を固定する回帰テスト。"""

import ast
import copy
import contextlib
import io
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import Mock

import generate_assessment as generator
from ai_assess_runtime import assessment_contract
from ai_assess_runtime.cli import CliDefaults, build_argument_parser
from ai_assess_runtime.deck_plan import build_default_slide_plan
from ai_assess_runtime.fingerprint import generator_source_sha256
from ai_assess_runtime.from_json import run_from_json
from ai_assess_runtime import poc_contract, presentation, quantitative_contract
from ai_assess_runtime.pptx_canvas import PptCanvas as RuntimePptCanvas
from ai_assess_runtime.review_snapshot import canonical_sha256, cost_estimate_from_snapshot
from ai_assess_runtime.source_input import load_source_text as runtime_load_source_text


class RefactorContractsTest(unittest.TestCase):
    def test_legacy_from_json_is_refrozen_and_strictly_revalidated(self) -> None:
        raw_payload = {
            "format": generator.ASSESSMENT_JSON_FORMAT,
            "input": {
                "source_file": "assessment_inputs/example.txt",
                "source_text": "会社名: Example株式会社\nサービス名: Example Service",
                "preprocessing": {"mode": "text"},
            },
            "assessment": {"company_name": "Example株式会社", "service_name": "Example Service"},
            "research": {"midterm_plan": {"status": "disabled"}},
        }
        migrated = copy.deepcopy(raw_payload)
        frozen = {
            "format": generator.ASSESSMENT_JSON_FROZEN_FORMAT,
            "assessment": copy.deepcopy(raw_payload["assessment"]),
            "research": copy.deepcopy(raw_payload["research"]),
            "rendering": {"cost_estimate": {}},
            "provenance": {},
        }
        migrate = Mock(return_value=migrated)
        rebuild = Mock(return_value=frozen)
        validate = Mock(return_value=[])
        load = Mock(side_effect=AssertionError("旧JSONを凍結v2として直接読み込んではいけません"))

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source_path = root / "legacy.json"
            output_path = root / "frozen.json"
            source_path.write_text(json.dumps(raw_payload, ensure_ascii=False), encoding="utf-8")
            api = SimpleNamespace(
                ASSESSMENT_JSON_FROZEN_FORMAT=generator.ASSESSMENT_JSON_FROZEN_FORMAT,
                DEFAULT_ARCHITECTURE_IMAGE=root / "unused.png",
                JSON_OUTPUT_DIR=root / "json",
                PPTX_OUTPUT_DIR=root / "pptx",
                RESEARCH_OUTPUT_DIR=root / "research",
                build_poc_cost_estimate=Mock(return_value={}),
                build_review_payload=rebuild,
                cost_estimate_from_snapshot=cost_estimate_from_snapshot,
                create_pptx=Mock(),
                load_assessment_payload=load,
                migrate_legacy_payload_for_review=migrate,
                safe_filename=generator.safe_filename,
                validate_assessment_payload=validate,
                write_render_manifest=Mock(),
            )
            args = SimpleNamespace(
                from_json=source_path, strict_json=False, include_source_text=False,
                json_only=True, output_json=output_path, architecture_image=None,
                output=None,
            )
            parser = SimpleNamespace(error=lambda message: self.fail(message))

            self.assertEqual(0, run_from_json(args, parser, api))
            self.assertEqual(frozen, json.loads(output_path.read_text(encoding="utf-8")))

        migrate.assert_called_once_with(raw_payload)
        rebuild.assert_called_once()
        validate.assert_called_once_with(frozen, strict=True)
        load.assert_not_called()

    def test_frozen_from_json_renders_exact_payload_and_manifest_hash(self) -> None:
        assessment = {
            "company_name": "Example株式会社",
            "service_name": "Example Service",
            "midterm_plan_analysis": {"management_targets": [{"target": "承認済み"}]},
        }
        research = {"midterm_plan": {"status": "found"}}
        assessment_hash = canonical_sha256(assessment)
        payload = {
            "format": generator.ASSESSMENT_JSON_FROZEN_FORMAT,
            "assessment": assessment,
            "research": research,
            "rendering": {
                "cost_estimate": {
                    "schema_version": "1", "currency": "JPY",
                    "lines": [{
                        "name": "test", "monthly_jpy": 1,
                        "assumption": "test", "source": "test",
                    }],
                    "total_monthly_jpy": 1, "notes": [], "api_priced_count": 0,
                },
            },
            "provenance": {
                "assessment_run_id": "run",
                "assessment_sha256": assessment_hash,
                "research_content_sha256": "research",
                "cost_estimate_sha256": "cost",
            },
        }
        loaded_payload = copy.deepcopy(payload)
        loaded_before = copy.deepcopy(loaded_payload)
        load_payload = Mock(return_value=loaded_payload)
        materialize = Mock(side_effect=AssertionError("凍結v2を再補完してはいけません"))
        rendered: dict[str, object] = {}

        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source_path = root / "frozen.json"
            architecture_path = root / "architecture.png"
            output_path = root / "assessment.pptx"
            source_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
            architecture_path.write_bytes(b"png")
            source_file_hash = generator.file_sha256(source_path)

            def create_pptx(exact_assessment, _architecture, output, _cost, _research=None) -> None:
                rendered["assessment"] = copy.deepcopy(exact_assessment)
                output.write_bytes(b"pptx")

            api = SimpleNamespace(
                ASSESSMENT_JSON_FROZEN_FORMAT=generator.ASSESSMENT_JSON_FROZEN_FORMAT,
                DEFAULT_ARCHITECTURE_IMAGE=architecture_path,
                JSON_OUTPUT_DIR=root / "json",
                PPTX_OUTPUT_DIR=root / "pptx",
                RESEARCH_OUTPUT_DIR=root / "research",
                build_poc_cost_estimate=Mock(side_effect=AssertionError("再見積り不可")),
                build_review_payload=Mock(),
                cost_estimate_from_snapshot=cost_estimate_from_snapshot,
                create_pptx=create_pptx,
                load_assessment_payload=load_payload,
                materialize_midterm_plan_targets=materialize,
                migrate_legacy_payload_for_review=Mock(),
                safe_filename=generator.safe_filename,
                validate_assessment_payload=Mock(return_value=[]),
                write_render_manifest=generator.write_render_manifest,
            )
            args = SimpleNamespace(
                from_json=source_path, strict_json=True, include_source_text=False,
                json_only=False, output_json=None, architecture_image=architecture_path,
                output=output_path,
            )
            parser = SimpleNamespace(error=lambda message: self.fail(message))
            self.assertEqual(0, run_from_json(args, parser, api))
            manifest = json.loads(
                output_path.with_suffix(".pptx.manifest.json").read_text(encoding="utf-8")
            )

            self.assertEqual(source_file_hash, generator.file_sha256(source_path))
            self.assertEqual(loaded_before, loaded_payload)
            self.assertEqual(payload["assessment"], rendered["assessment"])
            self.assertEqual(
                canonical_sha256(rendered["assessment"]),
                manifest["review"]["assessment_sha256"],
            )
            materialize.assert_not_called()
            load_payload.assert_called_once_with(source_path, strict=True)

    def test_poc_logic_accessor_does_not_mutate_nested_input(self) -> None:
        themes = ("需要予測", "問い合わせ回答支援", "業務サマリー")
        assessment = {
            "service_name": "Example Service",
            "poc_recommendations": [
                {"use_case_id": f"UC{index:02d}", "theme": theme}
                for index, theme in enumerate(themes, 1)
            ],
            "poc_logic_details": [
                {
                    "use_case_id": f"UC{index:02d}", "theme": theme,
                    "business_challenge": "challenge", "target_data": "data",
                    "processing_steps": [
                        {"label": "step 1", "description": "description 1"},
                        {"label": "step 2", "description": "description 2"},
                        {"label": "step 3", "description": "description 3"},
                    ],
                    "oci_roles": "roles", "business_integration": "integration",
                    "validation_plan": "validation", "design_notes": "notes",
                }
                for index, theme in enumerate(themes, 1)
            ],
        }
        before = copy.deepcopy(assessment)

        details = poc_contract.poc_logic_details_for(assessment)

        self.assertEqual(before, assessment)
        self.assertEqual("OMLによる予測・判定", details[0]["processing_steps"][1]["label"])
        details[0]["processing_steps"][0]["label"] = "changed"
        self.assertEqual(before, assessment)

    def test_generate_assessment_remains_the_compatible_public_facade(self) -> None:
        expected = (
            "main", "create_pptx", "load_source_text", "validate_assessment_payload",
            "build_review_payload", "PptCanvas", "PPTX_WIDESCREEN_WIDTH",
            "PPTX_WIDESCREEN_HEIGHT", "MIN_PPTX_FONT_SIZE",
        )
        for name in expected:
            with self.subTest(name=name):
                self.assertTrue(hasattr(generator, name))
        self.assertIs(RuntimePptCanvas, generator.PptCanvas)
        self.assertIs(runtime_load_source_text, generator.load_source_text)
        self.assertIs(assessment_contract.validate_assessment_payload, generator.validate_assessment_payload)
        self.assertIs(poc_contract.normalize_poc_portfolio, generator.normalize_poc_portfolio)
        self.assertIs(
            quantitative_contract.normalize_poc_measurement_design,
            generator.normalize_poc_measurement_design,
        )
        self.assertIs(presentation.create_pptx, generator.create_pptx)

    def test_cli_contract_is_pptx_only_and_keeps_current_defaults(self) -> None:
        defaults = CliDefaults(
            input_file=Path("assessment_inputs/input.xlsx"),
            provider="oci_responses",
            settings=SimpleNamespace(
                COMPARTMENT_ID="compartment", GENAI_ENDPOINT="endpoint", OCI_PROFILE="DEFAULT",
                OCI_CONFIG_FILE="config", OCI_GENAI_PROJECT_OCID="project", OCI_REGION="region",
            ),
        )
        parser = build_argument_parser(defaults)
        args = parser.parse_args([])
        self.assertEqual(defaults.input_file, args.input_file)
        self.assertEqual("oci_responses", args.provider)
        self.assertEqual(4, args.research_max_rounds)
        self.assertTrue(args.include_midterm_plan)
        for legacy in (("--format", "pdf"), ("--design", "graphical"), ("--detailed",)):
            with self.subTest(legacy=legacy), contextlib.redirect_stderr(io.StringIO()):
                with self.assertRaises(SystemExit) as error:
                    parser.parse_args(list(legacy))
                self.assertEqual(2, error.exception.code)

    def test_slide_plan_is_the_single_source_for_count_and_order(self) -> None:
        pure_plan = build_default_slide_plan(
            has_midterm=True,
            use_case_page_keys=["use_case_list:1:1", "use_case_list:1:2"],
            include_cost_estimate=True,
        )
        self.assertEqual(
            [
                "cover", "intro", "midterm", "industry_value", "business_impact",
                "use_case_list:1:1", "use_case_list:1:2", "poc_selection",
                "poc_detail:1", "poc_detail:2", "poc_detail:3", "architecture", "cost",
                "support", "closing",
            ],
            [spec.key for spec in pure_plan],
        )
        assessment = {
            "service_name": "Example Service",
            "service_genre": "Example",
            "use_cases": [{"no": index} for index in range(1, 17)],
            "midterm_plan_analysis": {"ai_alignment": [{"priority": "growth"}]},
        }
        generator_plan = generator.default_document_slide_plan(
            assessment, include_cost_estimate=True,
        )
        self.assertEqual([spec.key for spec in pure_plan], [spec.key for spec in generator_plan])
        self.assertEqual(len(generator_plan), generator.standard_document_page_count(assessment))

    def test_generator_fingerprint_includes_extracted_runtime_modules(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            project = Path(temporary_directory)
            (project / "generate_assessment.py").write_text("VALUE = 1\n", encoding="utf-8")
            runtime = project / "ai_assess_runtime"
            runtime.mkdir()
            module = runtime / "source_input.py"
            module.write_text("VALUE = 1\n", encoding="utf-8")
            before = generator_source_sha256(project)
            module.write_text("VALUE = 2\n", encoding="utf-8")
            after = generator_source_sha256(project)
        self.assertNotEqual(before, after)

    def test_generator_fingerprint_includes_rendering_assets(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            project = Path(temporary_directory)
            (project / "generate_assessment.py").write_text("VALUE = 1\n", encoding="utf-8")
            asset = project / "assets" / "icons" / "priority_insight.png"
            asset.parent.mkdir(parents=True)
            asset.write_bytes(b"icon-v1")
            before = generator_source_sha256(project)
            asset.write_bytes(b"icon-v2")
            after = generator_source_sha256(project)
        self.assertNotEqual(before, after)

    def test_runtime_modules_never_import_the_compatibility_facade(self) -> None:
        runtime_dir = Path(__file__).resolve().parents[1] / "ai_assess_runtime"
        for path in runtime_dir.glob("*.py"):
            tree = ast.parse(path.read_text(encoding="utf-8"))
            imported = {
                alias.name
                for node in ast.walk(tree)
                if isinstance(node, ast.Import)
                for alias in node.names
            }
            imported.update(
                str(node.module or "")
                for node in ast.walk(tree)
                if isinstance(node, ast.ImportFrom)
            )
            with self.subTest(path=path.name):
                self.assertNotIn("generate_assessment", imported)

    def test_runtime_dependency_direction_is_explicit_and_acyclic(self) -> None:
        allowed = {
            "assessment_contract": {
                "poc_contract", "quantitative_contract", "rendering_profile",
                "review_snapshot", "source_input",
            },
            "poc_contract": {"pptx_canvas"},
            "pptx_canvas": {"rendering_profile"},
            "presentation": {
                "deck_plan", "display_text", "paths", "poc_contract", "pptx_canvas",
                "presentation_common", "quantitative_contract",
            },
            "presentation_common": {"paths", "pptx_canvas"},
            "quantitative_contract": {"http_safety", "poc_contract", "target_metric"},
            "from_json": {"safe_io"},
            "use_case_catalog": {"poc_contract", "review_snapshot"},
            "workflow": {"pipeline", "safe_io"},
        }
        runtime_dir = Path(__file__).resolve().parents[1] / "ai_assess_runtime"
        actual: dict[str, set[str]] = {}
        for path in runtime_dir.glob("*.py"):
            tree = ast.parse(path.read_text(encoding="utf-8"))
            dependencies = {
                str(node.module).split(".")[-1]
                for node in ast.walk(tree)
                if isinstance(node, ast.ImportFrom)
                and str(node.module or "").startswith("ai_assess_runtime.")
            }
            if dependencies:
                actual[path.stem] = dependencies
        self.assertEqual(allowed, actual)

    def test_main_is_only_a_small_cli_dispatcher(self) -> None:
        path = Path(generator.__file__).resolve()
        tree = ast.parse(path.read_text(encoding="utf-8"))
        main = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "main")
        self.assertLessEqual(main.end_lineno - main.lineno + 1, 40)
        self.assertLess(path.read_text(encoding="utf-8").count("\n") + 1, 3_500)

    def test_runtime_modules_do_not_regrow_into_a_second_monolith(self) -> None:
        runtime_dir = Path(__file__).resolve().parents[1] / "ai_assess_runtime"
        for path in runtime_dir.glob("*.py"):
            with self.subTest(path=path.name):
                self.assertLess(
                    path.read_text(encoding="utf-8").count("\n") + 1,
                    3_000,
                )

    def test_manifest_round_trip_detects_artifact_tampering(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            artifact = Path(temporary_directory) / "assessment.pptx"
            artifact.write_bytes(b"pptx-v1")
            payload = {
                "provenance": {
                    "assessment_run_id": "run",
                    "assessment_sha256": "assessment",
                    "research_content_sha256": "research",
                    "cost_estimate_sha256": "cost",
                },
            }
            manifest = generator.write_render_manifest(
                artifact, payload, artifact_format="pptx", design="default",
            )
            self.assertEqual([], generator.validate_render_manifest(manifest))
            artifact.write_bytes(b"pptx-v2")
            errors = generator.validate_render_manifest(manifest)
        self.assertTrue(any("SHA-256" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
