import copy
import json
import tempfile
import unittest
from pathlib import Path

import generate_assessment as generator
from ai_assess_runtime.cli import CliDefaults, build_argument_parser


class RenderManifestIntegrityTests(unittest.TestCase):
    def _payload(self) -> dict:
        hashes = {
            "source_text_sha256": "1" * 64,
            "input_record_sha256": "2" * 64,
            "assessment_sha256": "3" * 64,
            "research_content_sha256": "4" * 64,
            "generator_sha256": "5" * 64,
            "cost_estimate_sha256": "6" * 64,
            "rendering_sha256": "7" * 64,
        }
        return {
            "format": "ai-assess/assessment-v2",
            "generated_at": "2026-08-28T00:00:00+09:00",
            "input": {"source_text_sha256": hashes["source_text_sha256"]},
            "assessment": {"company_name": "Example株式会社", "service_name": "Example"},
            "research": {"industry_sources": []},
            "rendering": {
                "cost_estimate": {"schema_version": "1"},
                "render_profile": {
                    "schema_version": "1",
                    "artifact_format": "pptx",
                    "page_layout": "widescreen-16:9",
                    "design": "default",
                    "min_editable_font_pt": 12,
                    "common_header_footer_font_pt": 10,
                },
                "architecture_asset": {
                    "schema_version": "1",
                    "source": "default",
                    "path": "assets/architecture/example.png",
                    "sha256": "8" * 64,
                    "bytes": 123,
                },
            },
            "provenance": {
                "format": "ai-assess/reproducibility-v2",
                "assessment_run_id": "9" * 24,
                **hashes,
            },
        }

    @staticmethod
    def _read(path: Path) -> dict:
        return json.loads(path.read_text(encoding="utf-8"))

    @staticmethod
    def _write(path: Path, value: dict) -> None:
        path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")

    def _manifest(self, root: Path, payload: dict | None = None) -> tuple[Path, Path, dict]:
        artifact = root / "assessment.pptx"
        artifact.write_bytes(b"pptx-artifact")
        frozen = copy.deepcopy(payload or self._payload())
        manifest = generator.write_render_manifest(
            artifact, frozen, artifact_format="pptx", design="default",
        )
        return artifact, manifest, frozen

    def test_new_manifest_binds_internal_records_and_matching_review_json(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            _, manifest, payload = self._manifest(root)
            review_json = root / "review.json"
            self._write(review_json, payload)

            data = self._read(manifest)
            self.assertEqual("1", data["binding"]["schema_version"])
            self.assertEqual(data["review"]["assessment_sha256"],
                             data["provenance_reference"]["assessment_sha256"])
            self.assertEqual([], generator.validate_render_manifest(manifest, review_json))

    def test_review_or_provenance_reference_edit_is_detected(self) -> None:
        for field in ("review", "provenance_reference"):
            with self.subTest(field=field), tempfile.TemporaryDirectory() as temporary_directory:
                root = Path(temporary_directory)
                _, manifest, _ = self._manifest(root)
                data = self._read(manifest)
                data[field]["assessment_sha256"] = "0" * 64
                self._write(manifest, data)

                errors = generator.validate_render_manifest(manifest)

                self.assertTrue(any("binding" in error for error in errors), errors)
                self.assertTrue(any("assessment_sha256" in error for error in errors), errors)

    def test_current_renderer_edit_and_generator_drift_are_detected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            _, manifest, _ = self._manifest(root)
            data = self._read(manifest)
            data["renderer"]["current_generator_sha256"] = "0" * 64
            self._write(manifest, data)

            errors = generator.validate_render_manifest(manifest)

        self.assertTrue(any("current_generator_sha256" in error for error in errors), errors)

    def test_render_profile_and_architecture_contract_edits_are_detected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            _, manifest, _ = self._manifest(root)
            data = self._read(manifest)
            data["renderer"]["render_profile"]["page_layout"] = "standard-4:3"
            data["renderer"]["architecture_asset"]["sha256"] = "not-a-sha"
            self._write(manifest, data)

            errors = generator.validate_render_manifest(manifest)

        self.assertTrue(any("page_layout" in error for error in errors), errors)
        self.assertTrue(any("architecture_asset.sha256" in error for error in errors), errors)

    def test_required_artifact_types_are_checked(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            _, manifest, _ = self._manifest(root)
            data = self._read(manifest)
            data["artifact"]["sha256"] = 123
            data["artifact"]["bytes"] = True
            self._write(manifest, data)

            errors = generator.validate_render_manifest(manifest)

        self.assertTrue(any("artifact.sha256" in error for error in errors), errors)
        self.assertTrue(any("artifact.bytes" in error for error in errors), errors)

    def test_swapped_review_json_is_detected(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            _, manifest, payload = self._manifest(root)
            wrong_payload = copy.deepcopy(payload)
            wrong_payload["assessment"]["service_name"] = "Other Service"
            wrong_review = root / "wrong-review.json"
            self._write(wrong_review, wrong_payload)

            errors = generator.validate_render_manifest(manifest, wrong_review)

        self.assertTrue(any("レビューJSON" in error and "一致しません" in error for error in errors), errors)

    def test_binding_free_legacy_manifest_keeps_previous_validation_behavior(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            _, manifest, _ = self._manifest(root)
            data = self._read(manifest)
            data.pop("binding")
            data.pop("provenance_reference")
            # 既存sidecarではrenderer driftを記録値として保持し、過去成果物の
            # artifact SHA/bytes検証を引き続き利用できるようにする。
            data["renderer"]["current_generator_sha256"] = "0" * 64
            self._write(manifest, data)

            self.assertEqual([], generator.validate_render_manifest(manifest))

    def test_cli_accepts_optional_review_json_for_manifest_verification(self) -> None:
        defaults = CliDefaults(
            input_file=Path("input.xlsx"), provider="oci_responses",
            settings=generator.settings,
        )
        args = build_argument_parser(defaults).parse_args([
            "--verify-manifest", "assessment.pptx.manifest.json",
            "--verify-review-json", "assessment.json",
        ])
        self.assertEqual(Path("assessment.pptx.manifest.json"), args.verify_manifest)
        self.assertEqual(Path("assessment.json"), args.verify_review_json)


if __name__ == "__main__":
    unittest.main()
