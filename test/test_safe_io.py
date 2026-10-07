import json
from pathlib import Path
import tempfile
import unittest

from ai_assess_runtime.safe_io import (
    atomic_output_path,
    atomic_render_immutable,
    atomic_write_json,
    atomic_write_text,
)


class SafeIoTests(unittest.TestCase):
    def test_atomic_write_json_replaces_complete_document(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "review.json"
            path.write_text("old", encoding="utf-8")
            atomic_write_json(path, {"value": "新"})
            self.assertEqual(json.loads(path.read_text(encoding="utf-8")), {"value": "新"})
            self.assertFalse(list(path.parent.glob(f".{path.name}.*.tmp")))

    def test_failed_atomic_output_preserves_existing_artifact(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "artifact.pptx"
            path.write_bytes(b"approved")
            with self.assertRaisesRegex(RuntimeError, "render failed"):
                with atomic_output_path(path) as temporary:
                    temporary.write_bytes(b"partial")
                    raise RuntimeError("render failed")
            self.assertEqual(path.read_bytes(), b"approved")
            self.assertFalse(list(path.parent.glob(f".{path.name}.*.tmp")))

    def test_atomic_write_text_creates_parent_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "nested" / "value.txt"
            atomic_write_text(path, "complete")
            self.assertEqual(path.read_text(encoding="utf-8"), "complete")

    def test_atomic_render_rejects_mutation_before_replacing_artifact(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "assessment.pptx"
            path.write_bytes(b"approved")
            frozen = {"company_name": "Example株式会社"}

            def mutating_renderer(temporary: Path) -> None:
                temporary.write_bytes(b"mutated-output")
                frozen["company_name"] = "changed"

            with self.assertRaisesRegex(RuntimeError, "approved value changed"):
                atomic_render_immutable(
                    path,
                    frozen,
                    mutating_renderer,
                    mutation_message="approved value changed",
                )

            self.assertEqual(b"approved", path.read_bytes())
            self.assertFalse(list(path.parent.glob(f".{path.name}.*.tmp")))

    def test_atomic_render_commits_when_approved_value_is_unchanged(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "assessment.pptx"
            frozen = {"company_name": "Example株式会社"}

            atomic_render_immutable(
                path,
                frozen,
                lambda temporary: temporary.write_bytes(b"complete"),
                mutation_message="must not mutate",
            )

            self.assertEqual(b"complete", path.read_bytes())
            self.assertEqual({"company_name": "Example株式会社"}, frozen)


if __name__ == "__main__":
    unittest.main()
