import io
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch

from assessment_input_workbook import FormQuestion, list_questions
from assessment_input_workbook import write_answers as write_workbook_answers
import fill_assessment_input as filler


def question(
    question_id: str,
    text: str,
    *,
    row: int,
    required: bool = False,
    choices: tuple[str, ...] = (),
    multi: bool = False,
    current_answer: str = "",
) -> FormQuestion:
    return FormQuestion(
        id=question_id,
        row=row,
        required=required,
        question=text,
        guide=f"{text}の入力ガイド",
        input_type="Choice（複数選択）" if multi else ("Choice（単一選択）" if choices else "Text（長い回答）"),
        choices=choices,
        current_answer=current_answer,
        multi_select_delimiter=";" if multi else None,
    )


QUESTIONS = (
    question("F001", "情報利用目的を確認しましたか？", row=7, required=True, choices=("確認しました",)),
    question("F101", "会社名", row=10, required=True),
    question("F103", "事業概要", row=12, required=True),
    question("F105", "対象業種・業態", row=14, choices=("製造業", "物流/SCM", "その他"), multi=True),
    question("F107", "入力者氏名", row=16, required=True),
    question("F108", "入力者メール", row=17, required=True),
    question("F109", "入力者所属/役職", row=18),
    question("F202", "対象プロジェクト数", row=21, choices=("1件", "2件", "3件以上", "未定")),
    question("F1010", "入力内容の共有可否", row=123, required=True, choices=("共有可", "共有不可")),
)


def response(*answers: dict[str, str]) -> str:
    return json.dumps({"answers": list(answers)}, ensure_ascii=False)


def answer_record(
    question_id: str,
    answer: str,
    *,
    basis: str = "explicit",
    confidence: str = "high",
    reason: str = "入力文に明記されている",
) -> dict[str, str]:
    return {
        "question_id": question_id,
        "answer": answer,
        "basis": basis,
        "confidence": confidence,
        "reason": reason,
    }


class FakeResponses:
    def __init__(self, output_text: str) -> None:
        self.output_text = output_text
        self.calls: list[dict[str, object]] = []

    def create(self, **arguments: object) -> object:
        self.calls.append(arguments)
        return SimpleNamespace(output_text=self.output_text)


class FakeClient:
    def __init__(self, output_text: str) -> None:
        self.responses = FakeResponses(output_text)


class FillAssessmentInputTests(unittest.TestCase):
    def test_local_extraction_accepts_id_and_label_and_redacts_remaining_pii(self) -> None:
        source = (
            "入力者氏名：山田 太郎\n"
            "F108: taro.yamada@example.com\n"
            "F109：AI推進部 部長\n"
            "問い合わせ先は support@example.net / 090-1234-5678\n"
            "会社はExample株式会社です。\n"
        )
        redacted, local = filler.extract_local_answers(source, QUESTIONS)

        self.assertEqual(
            {"F107": "山田 太郎", "F108": "taro.yamada@example.com", "F109": "AI推進部 部長"},
            local,
        )
        self.assertNotIn("F108", redacted)
        self.assertNotIn("入力者氏名", redacted)
        self.assertNotIn("山田", redacted)
        self.assertNotIn("support@example.net", redacted)
        self.assertNotIn("090-1234-5678", redacted)
        self.assertIn("[REDACTED_EMAIL]", redacted)
        self.assertIn("[REDACTED_PHONE]", redacted)
        self.assertIn("Example株式会社", redacted)

    def test_prompt_and_review_never_contain_local_values_or_contact_details(self) -> None:
        source = (
            "F001: 確認しました\n"
            "入力者氏名: 山田 太郎\n"
            "入力者メール: taro.yamada@example.com\n"
            "F109: AI推進部 部長\n"
            "入力内容の共有可否: 共有可\n"
            "Example株式会社は物流向けサービスを提供。連絡先 03-1234-5678\n"
        )
        fake = FakeClient(response(answer_record("F101", "Example株式会社")))

        review, _answers, prompt_text = filler.generate_review_and_answers(
            source_text=source,
            source_bytes=source.encode(),
            template_path=filler.DEFAULT_TEMPLATE,
            questions=QUESTIONS,
            client=fake,
            provider="test",
            model_id="test-model",
            overwrite=False,
        )

        self.assertEqual(1, len(fake.responses.calls))
        self.assertEqual(prompt_text, fake.responses.calls[0]["input"])
        for secret in (
            "F001", "F107", "F108", "F109", "F1010", "山田 太郎",
            "taro.yamada@example.com", "AI推進部 部長", "03-1234-5678",
        ):
            self.assertNotIn(secret, prompt_text)
        serialized = json.dumps(review, ensure_ascii=False)
        self.assertNotIn(source, serialized)
        self.assertNotIn("山田 太郎", serialized)
        self.assertNotIn("taro.yamada@example.com", serialized)
        self.assertNotIn("AI推進部 部長", serialized)
        self.assertEqual(
            ["F001", "F107", "F108", "F109", "F1010"],
            review["local_answer_ids"],
        )
        self.assertEqual(filler._sha256_bytes(source.encode()), review["source"]["sha256"])

    def test_response_validation_rejects_unknown_duplicate_choice_number_and_local_id(self) -> None:
        valid = answer_record("F101", "Example株式会社")
        invalid_cases = {
            "unknown": response(answer_record("F999", "value")),
            "duplicate": response(valid, valid),
            "single_choice": response(answer_record("F202", "2プロジェクト")),
            "multi_choice_value": response(answer_record("F105", "製造業;金融")),
            "multi_choice_delimiter": response(answer_record("F105", "製造業、物流/SCM")),
            "unsupported_number": response(answer_record("F103", "売上を30%改善する")),
            "local_only": response(answer_record("F107", "山田 太郎")),
        }
        for label, model_output in invalid_cases.items():
            with self.subTest(label=label), self.assertRaises(ValueError):
                filler.parse_model_response(
                    model_output, QUESTIONS, allowed_quantitative_tokens=set(),
                )

    def test_response_accepts_exact_choices_semicolon_and_source_numbers(self) -> None:
        records = filler.parse_model_response(
            response(
                answer_record("F105", "製造業;物流/SCM", basis="inferred", confidence="medium"),
                answer_record("F202", "2件", reason="入力文に2件と明記されている"),
                answer_record("F103", "現行比30%の改善を確認したい", reason="入力文の30%を転記"),
            ),
            QUESTIONS,
            allowed_quantitative_tokens=filler.quantitative_tokens("対象は2件。目標は現行比30%。"),
        )
        self.assertEqual(["F105", "F202", "F103"], [item["question_id"] for item in records])

    def test_explicit_only_fields_reject_inference_and_reason_numbers_are_not_claims(self) -> None:
        with self.assertRaises(ValueError):
            filler.parse_model_response(
                response(answer_record("F101", "Example株式会社", basis="inferred")),
                QUESTIONS,
                allowed_quantitative_tokens=set(),
            )
        records = filler.parse_model_response(
            response(answer_record(
                "F101", "Example株式会社", basis="explicit",
                reason="設問1の入力記述を根拠に転記",
            )),
            QUESTIONS,
            allowed_quantitative_tokens=set(),
        )
        self.assertEqual("F101", records[0]["question_id"])

    def test_extracted_local_choice_must_match_exactly(self) -> None:
        fake = FakeClient(response())
        with self.assertRaises(ValueError):
            filler.generate_review_and_answers(
                source_text="F001: はい、確認しました",
                source_bytes="F001: はい、確認しました".encode(),
                template_path=filler.DEFAULT_TEMPLATE,
                questions=QUESTIONS,
                client=fake,
                provider="test",
                model_id="test-model",
                overwrite=False,
            )
        self.assertEqual([], fake.responses.calls)

    def test_review_roundtrip_validates_without_raw_text_or_local_values(self) -> None:
        source = "F001: 確認しました\nF107: 佐藤 花子\nF108: hanako@example.com\nF1010: 共有可\n会社名はExample株式会社。"
        fake = FakeClient(response(answer_record("F101", "Example株式会社")))
        review, _answers, _prompt = filler.generate_review_and_answers(
            source_text=source,
            source_bytes=source.encode(),
            template_path=filler.DEFAULT_TEMPLATE,
            questions=QUESTIONS,
            client=fake,
            provider="test",
            model_id="test-model",
            overwrite=False,
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            review_path = Path(temporary_directory) / "review.json"
            filler._write_json_atomic(review_path, review)
            loaded = filler.load_review(review_path, QUESTIONS)

        self.assertEqual(review, loaded)
        serialized = json.dumps(loaded, ensure_ascii=False)
        self.assertNotIn("佐藤", serialized)
        self.assertNotIn("hanako@example.com", serialized)
        self.assertNotIn("会社名はExample株式会社", serialized)
        self.assertEqual(["F001", "F107", "F108", "F1010"], loaded["local_answer_ids"])

    def test_main_writes_template_copy_and_sidecar_with_mocked_client(self) -> None:
        source = (
            "F001: 確認しました\n"
            "F002: 確認しました\n"
            "F107: 山田 太郎\n"
            "F108: taro.yamada@example.com\n"
            "F109: AI推進部\n"
            "F1010: 共有可\n"
            "会社名はExample株式会社。事業は物流ソフトウェア。\n"
        )
        fake = FakeClient(response(
            answer_record("F101", "Example株式会社"),
            answer_record("F103", "物流ソフトウェアの開発・提供"),
        ))
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "filled.xlsx"
            sidecar = output.with_suffix(".review.json")
            with patch.object(filler, "create_model_client", return_value=fake), \
                 patch("sys.stdout", new_callable=io.StringIO):
                result = filler.main([
                    "--text", source,
                    "--template", str(filler.DEFAULT_TEMPLATE),
                    "--output", str(output),
                    "--model-id", "test-model",
                ])

            self.assertEqual(0, result)
            self.assertTrue(output.is_file())
            self.assertTrue(sidecar.is_file())
            values = {item.id: item.current_answer for item in list_questions(output)}
            self.assertEqual("Example株式会社", values["F101"])
            self.assertEqual("物流ソフトウェアの開発・提供", values["F103"])
            self.assertEqual("山田 太郎", values["F107"])
            self.assertEqual("taro.yamada@example.com", values["F108"])
            payload = json.loads(sidecar.read_text(encoding="utf-8"))
            serialized = json.dumps(payload, ensure_ascii=False)
            self.assertNotIn("山田 太郎", serialized)
            self.assertNotIn("taro.yamada@example.com", serialized)
            self.assertEqual(1, payload["model"]["calls"])
            self.assertEqual(1, len(fake.responses.calls))

    def test_from_json_uses_review_overwrite_and_reapplies_local_answer_without_model(self) -> None:
        actual_questions = list_questions(filler.DEFAULT_TEMPLATE)
        source = "会社名はNew株式会社。"
        fake = FakeClient(response(answer_record("F101", "New株式会社")))
        review, _answers, _prompt = filler.generate_review_and_answers(
            source_text=source,
            source_bytes=source.encode(),
            template_path=filler.DEFAULT_TEMPLATE,
            questions=actual_questions,
            client=fake,
            provider="test",
            model_id="test-model",
            overwrite=True,
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary_path = Path(temporary_directory)
            populated_template = temporary_path / "existing.xlsx"
            output = temporary_path / "rendered.xlsx"
            review_path = temporary_path / "review.json"
            write_workbook_answers(
                filler.DEFAULT_TEMPLATE, populated_template,
                {"F101": "Old株式会社"}, overwrite=False,
            )
            filler._write_json_atomic(review_path, review)
            with patch.object(
                filler, "create_model_client", side_effect=AssertionError("model must not be called")
            ), patch("sys.stdout", new_callable=io.StringIO):
                result = filler.main([
                    "--from-json", str(review_path),
                    "--template", str(populated_template),
                    "--output", str(output),
                    "--answer", "F107=鈴木 一郎",
                ])
            values = {item.id: item.current_answer for item in list_questions(output)}

        self.assertEqual(0, result)
        self.assertEqual("New株式会社", values["F101"])
        self.assertEqual("鈴木 一郎", values["F107"])

    def test_malformed_response_does_not_create_workbook_or_review(self) -> None:
        fake = FakeClient("not json")
        with tempfile.TemporaryDirectory() as temporary_directory:
            output = Path(temporary_directory) / "must-not-exist.xlsx"
            sidecar = output.with_suffix(".review.json")
            with patch.object(filler, "create_model_client", return_value=fake), \
                 patch("sys.stderr", new_callable=io.StringIO), \
                 self.assertRaises(SystemExit):
                filler.main([
                    "--text", "Example株式会社のサービスです。",
                    "--template", str(filler.DEFAULT_TEMPLATE),
                    "--output", str(output),
                    "--model-id", "test-model",
                ])
            self.assertFalse(output.exists())
            self.assertFalse(sidecar.exists())
            self.assertEqual(1, len(fake.responses.calls))

    def test_validate_json_rejects_answer_override(self) -> None:
        with patch("sys.stderr", new_callable=io.StringIO), self.assertRaises(SystemExit):
            filler.main([
                "--validate-json", "does-not-matter.json",
                "--answer", "F101=Example株式会社",
            ])


if __name__ == "__main__":
    unittest.main()
