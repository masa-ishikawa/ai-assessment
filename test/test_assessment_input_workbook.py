from pathlib import Path
import tempfile
import unittest
import zipfile

from assessment_input_workbook import read_form_schema, write_answers


PROJECT_DIR = Path(__file__).resolve().parents[1]
TEMPLATE = PROJECT_DIR / "assessment_inputs" / "ISV_AI_Use_Case_Assessment_Input_Template.xlsx"


class AssessmentInputWorkbookTests(unittest.TestCase):
    def test_discovers_complete_form_schema_and_validation_choices(self) -> None:
        schema = read_form_schema(TEMPLATE)
        questions = {question.id: question for question in schema.questions}

        self.assertEqual("入力フォーム", schema.sheet_name)
        self.assertEqual(107, len(schema.questions))
        self.assertEqual(24, sum(question.required for question in schema.questions))
        self.assertEqual(10, questions["F101"].row)
        self.assertEqual("会社名", questions["F101"].question)
        self.assertEqual("例：株式会社〇〇", questions["F101"].guide)
        self.assertEqual("Text（短い回答）", questions["F101"].input_type)
        self.assertEqual((), questions["F101"].choices)
        self.assertEqual(("高", "中", "低", "未定"), questions["F213"].choices)
        self.assertIn("製造業", questions["F105"].choices)
        self.assertEqual("Choice（複数選択）", questions["F105"].input_type)
        self.assertEqual(";", questions["F105"].multi_select_delimiter)
        self.assertIsNone(questions["F213"].multi_select_delimiter)

    def test_write_changes_only_the_input_worksheet_package_part(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "populated.xlsx"
            summary = write_answers(TEMPLATE, output, {"F101": "テスト株式会社"})

            self.assertEqual(1, summary.written_count)
            self.assertEqual(("F101",), summary.written)
            with zipfile.ZipFile(TEMPLATE) as source, zipfile.ZipFile(output) as populated:
                self.assertEqual(source.namelist(), populated.namelist())
                for member in source.namelist():
                    if member == summary.worksheet_path:
                        self.assertNotEqual(source.read(member), populated.read(member))
                    else:
                        self.assertEqual(source.read(member), populated.read(member), member)

    def test_existing_answers_are_preserved_then_overwritten_only_when_requested(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            directory = Path(temporary_directory)
            initial = directory / "initial.xlsx"
            preserved = directory / "preserved.xlsx"
            overwritten = directory / "overwritten.xlsx"
            write_answers(TEMPLATE, initial, {"F101": "既存会社名"})

            preserve_summary = write_answers(
                initial,
                preserved,
                {"F101": "新しい会社名", "F103": "新しい事業概要"},
            )
            preserved_questions = {
                question.id: question for question in read_form_schema(preserved).questions
            }
            self.assertEqual(("F101",), preserve_summary.preserved)
            self.assertEqual(("F103",), preserve_summary.written)
            self.assertEqual("既存会社名", preserved_questions["F101"].current_answer)
            self.assertEqual("新しい事業概要", preserved_questions["F103"].current_answer)

            overwrite_summary = write_answers(
                preserved,
                overwritten,
                {"F101": "新しい会社名"},
                overwrite=True,
            )
            overwritten_questions = {
                question.id: question for question in read_form_schema(overwritten).questions
            }
            self.assertEqual(("F101",), overwrite_summary.written)
            self.assertEqual((), overwrite_summary.preserved)
            self.assertEqual("新しい会社名", overwritten_questions["F101"].current_answer)

    def test_rejects_same_template_and_output_path_without_modifying_it(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook = Path(temporary_directory) / "template.xlsx"
            workbook.write_bytes(TEMPLATE.read_bytes())
            original = workbook.read_bytes()

            with self.assertRaisesRegex(ValueError, "different from the template"):
                write_answers(workbook, workbook, {"F101": "変更不可"})

            self.assertEqual(original, workbook.read_bytes())


if __name__ == "__main__":
    unittest.main()
