"""PPTX専用描画層の構造・再現性・版面契約を固定するテスト。"""

import ast
import copy
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from reportlab.lib.units import mm
from pptx.util import Pt

from ai_assess_runtime import presentation, presentation_common
from ai_assess_runtime.paths import DEFAULT_ARCHITECTURE_IMAGE
from ai_assess_runtime.pptx_canvas import (
    PPTX_FOOTER_SHAPE_NAME,
    PPTX_HEADER_SHAPE_NAME,
    PPTX_SOURCE_NOTE_SHAPE_NAME,
)
from test import test_generate_assessment as fixtures


class PptxOnlyPresentationTest(unittest.TestCase):
    @staticmethod
    def _is_ten_point_chrome(shape, _slide_height: int) -> bool:
        """文字列や座標ではなく、共通描画層の明示タグだけを例外扱いする。"""
        return shape.name in {PPTX_HEADER_SHAPE_NAME, PPTX_FOOTER_SHAPE_NAME}

    @staticmethod
    def _is_nine_point_source(shape) -> bool:
        return shape.name == PPTX_SOURCE_NOTE_SHAPE_NAME

    @staticmethod
    def _assessment() -> dict:
        return fixtures.DynamicPptxPageTests.detailed_assessment()

    def test_presentation_module_has_no_pdf_renderer_or_runtime_type_branch(self) -> None:
        source_path = Path(presentation.__file__)
        source = source_path.read_text(encoding="utf-8")
        tree = ast.parse(source)
        imported_modules = {
            node.module
            for node in ast.walk(tree)
            if isinstance(node, ast.ImportFrom) and node.module
        }

        self.assertFalse({
            "reportlab.pdfgen", "reportlab.platypus",
            "reportlab.pdfbase", "reportlab.pdfbase.ttfonts",
        } & imported_modules)
        self.assertNotIn("canvas.Canvas", source)
        self.assertNotIn("isinstance(pdf, PptCanvas)", source)
        self.assertNotIn("Paragraph(", source)
        self.assertNotIn("Table(", source)

    def test_common_primitives_are_split_and_reexported_without_wrappers(self) -> None:
        """共通描画の実体は一箇所に置き、旧公開名は同一オブジェクトを返す。"""
        common_exports = (
            "assessment_subsection_label",
            "poc_design_subsection_label",
            "register_japanese_font",
            "draw_footer",
            "draw_paragraph",
            "draw_number_badge",
            "wide_content_width",
            "estimate_wide_table_lines",
            "wide_table_row_height",
            "basis_label",
            "draw_story_source_note",
            "draw_pptx_base_header",
            "_executive_evidence_header",
        )
        presentation_tree = ast.parse(Path(presentation.__file__).read_text(encoding="utf-8"))
        local_functions = {
            node.name for node in presentation_tree.body if isinstance(node, ast.FunctionDef)
        }
        for name in common_exports:
            self.assertNotIn(name, local_functions)
            self.assertIs(getattr(presentation, name), getattr(presentation_common, name))

        self.assertEqual(12.0, presentation.MIN_PPTX_FONT_SIZE)
        self.assertEqual(10.0, presentation_common.PPTX_HEADER_FONT_SIZE)
        self.assertEqual(10.0, presentation.PPTX_FOOTER_FONT_SIZE)

    def test_footer_stays_above_dynamic_label_safe_zone_at_ten_point(self) -> None:
        """テナントの動的機密ラベルとの重なりを、文字を縮小せず位置で防ぐ。"""
        width = presentation.PPTX_WIDESCREEN_WIDTH
        height = presentation.PPTX_WIDESCREEN_HEIGHT
        canvas = presentation.PptCanvas(width, height)
        presentation.draw_footer(canvas, width, 2)
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "footer-safe-zone.pptx"
            canvas.write(path)
            deck = Presentation(path)

        slide = deck.slides[0]
        footer_shapes = [
            shape for shape in slide.shapes if shape.has_text_frame
            and (shape.text.startswith("Copyright ©") or shape.text.strip() == "2")
        ]
        self.assertEqual(2, len(footer_shapes))
        safe_bottom = round(5.5 * mm * presentation.PptCanvas.POINT_TO_EMU)
        for shape in footer_shapes:
            self.assertLessEqual(shape.top + shape.height, deck.slide_height - safe_bottom)
            runs = [
                run for paragraph in shape.text_frame.paragraphs
                for run in paragraph.runs if run.text.strip()
            ]
            self.assertTrue(runs)
            self.assertTrue(all(
                run.font.size.pt == presentation.PPTX_FOOTER_FONT_SIZE == 10.0
                for run in runs
            ))
            self.assertEqual(PPTX_FOOTER_SHAPE_NAME, shape.name)

    def test_common_header_is_tagged_and_kept_at_ten_point(self) -> None:
        width = presentation.PPTX_WIDESCREEN_WIDTH
        height = presentation.PPTX_WIDESCREEN_HEIGHT
        canvas = presentation.PptCanvas(width, height)
        presentation.draw_pptx_base_header(canvas, width, height, "本文タイトル", "section")
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "header-font-contract.pptx"
            canvas.write(path)
            deck = Presentation(path)

        header = next(
            shape for shape in deck.slides[0].shapes
            if shape.name == PPTX_HEADER_SHAPE_NAME
        )
        header_sizes = [
            run.font.size.pt
            for paragraph in header.text_frame.paragraphs
            for run in paragraph.runs if run.text.strip()
        ]
        self.assertEqual([10.0], header_sizes)
        body_sizes = [
            run.font.size.pt
            for shape in deck.slides[0].shapes
            if shape.has_text_frame and shape.shape_id != header.shape_id
            for paragraph in shape.text_frame.paragraphs
            for run in paragraph.runs if run.text.strip()
        ]
        self.assertTrue(body_sizes)
        self.assertTrue(all(size >= 12.0 for size in body_sizes))

    def test_normal_font_call_cancels_an_unused_chrome_exception(self) -> None:
        canvas = presentation.PptCanvas(
            presentation.PPTX_WIDESCREEN_WIDTH,
            presentation.PPTX_WIDESCREEN_HEIGHT,
        )
        canvas.setFooterFont("AssessmentJapanese", 10)
        canvas.setFont("AssessmentJapanese", 8)
        canvas.drawString(20 * mm, 40 * mm, "本文")
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "no-chrome-role-leak.pptx"
            canvas.write(path)
            deck = Presentation(path)

        shape = next(
            shape for shape in deck.slides[0].shapes
            if shape.has_text_frame and shape.text == "本文"
        )
        self.assertNotIn(shape.name, {PPTX_HEADER_SHAPE_NAME, PPTX_FOOTER_SHAPE_NAME})
        self.assertEqual(12.0, shape.text_frame.paragraphs[0].runs[0].font.size.pt)

    def test_write_repairs_direct_and_inherited_text_below_twelve_point(self) -> None:
        """個別レイアウトが共通ヘルパーを迂回しても最終成果物では12pt以上にする。"""
        width = presentation.PPTX_WIDESCREEN_WIDTH
        height = presentation.PPTX_WIDESCREEN_HEIGHT
        canvas = presentation.PptCanvas(width, height)

        text_shape = canvas.slide.shapes.add_textbox(
            canvas._x(20 * mm), canvas._y(height - 40 * mm),
            canvas._width(80 * mm), canvas._height(12 * mm),
        )
        direct_run = text_shape.text_frame.paragraphs[0].add_run()
        direct_run.text = "direct small text"
        direct_run.font.size = Pt(8)

        inherited_shape = canvas.slide.shapes.add_textbox(
            canvas._x(20 * mm), canvas._y(height - 60 * mm),
            canvas._width(80 * mm), canvas._height(12 * mm),
        )
        inherited_run = inherited_shape.text_frame.paragraphs[0].add_run()
        inherited_run.text = "theme inherited text"
        self.assertIsNone(inherited_run.font.size)

        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "minimum-font-guard.pptx"
            canvas.write(path)
            deck = Presentation(path)

        sizes = [
            run.font.size.pt
            for shape in deck.slides[0].shapes if shape.has_text_frame
            for paragraph in shape.text_frame.paragraphs
            for run in paragraph.runs if run.text.strip()
        ]
        self.assertEqual([12.0, 12.0], sizes)

    def test_industry_sources_are_tagged_nine_point_and_clear_the_footer(self) -> None:
        assessment = self._assessment()
        research = {
            "industry_trend_digest": {
                "title": "最新動向とAI活用",
                "lead": "公開情報から対象業務の変化を整理します。",
                "rows": [
                    {
                        "topic": f"注目トピック{index}",
                        "development": "新しい使われ方と業務上の変化を整理します。",
                        "implication": "対象サービスへの示唆を整理します。",
                        "url": f"https://example.com/research/topic-{index}/long-source-path",
                    }
                    for index in range(1, 4)
                ],
                "proposal": "対象サービスのデータを活かすAI活用が有力です。",
                "policy_source_url": "https://example.com/policy/current-guidance",
            }
        }
        canvas = presentation.PptCanvas(
            presentation.PPTX_WIDESCREEN_WIDTH,
            presentation.PPTX_WIDESCREEN_HEIGHT,
        )
        presentation.draw_industry_trends_page(
            canvas,
            presentation.PPTX_WIDESCREEN_WIDTH,
            presentation.PPTX_WIDESCREEN_HEIGHT,
            assessment,
            research,
            3,
        )
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "industry-source-font.pptx"
            canvas.write(path)
            deck = Presentation(path)

        slide = deck.slides[0]
        source = next(shape for shape in slide.shapes if shape.name == PPTX_SOURCE_NOTE_SHAPE_NAME)
        footer = next(shape for shape in slide.shapes if shape.name == PPTX_FOOTER_SHAPE_NAME)
        source_sizes = [
            run.font.size.pt
            for paragraph in source.text_frame.paragraphs
            for run in paragraph.runs if run.text.strip()
        ]
        self.assertEqual([9.0], source_sizes)
        self.assertLessEqual(source.top + source.height, footer.top)

    def test_width_estimation_uses_same_twelve_point_floor_as_rendering(self) -> None:
        canvas = presentation.PptCanvas(
            presentation.PPTX_WIDESCREEN_WIDTH,
            presentation.PPTX_WIDESCREEN_HEIGHT,
        )
        self.assertEqual(
            canvas.stringWidth("ABC日本", "AssessmentJapanese", 12),
            canvas.stringWidth("ABC日本", "AssessmentJapanese", 8),
        )

    def test_frozen_pptx_renderer_is_deterministic_widescreen_and_safe(self) -> None:
        assessment = self._assessment()
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            first_path = root / "first.pptx"
            second_path = root / "second.pptx"
            presentation.create_pptx(
                assessment, DEFAULT_ARCHITECTURE_IMAGE, first_path,
            )
            presentation.create_pptx(
                assessment, DEFAULT_ARCHITECTURE_IMAGE, second_path,
            )

            with ZipFile(first_path) as first_zip, ZipFile(second_path) as second_zip:
                self.assertEqual(first_zip.namelist(), second_zip.namelist())
                self.assertEqual(
                    {name: first_zip.read(name) for name in first_zip.namelist()},
                    {name: second_zip.read(name) for name in second_zip.namelist()},
                )

            deck = Presentation(first_path)

        self.assertEqual(round(presentation.PPTX_WIDESCREEN_WIDTH * 12700), deck.slide_width)
        self.assertEqual(round(presentation.PPTX_WIDESCREEN_HEIGHT * 12700), deck.slide_height)
        self.assertAlmostEqual(16 / 9, deck.slide_width / deck.slide_height, places=6)
        self.assertEqual(
            "AIユースケースアセスメント｜"
            f"{assessment['service_name']}におけるAI活用・PoCのご提案",
            deck.core_properties.title,
        )
        self.assertEqual(
            "AIユースケースアセスメント", deck.core_properties.subject,
        )

        text_runs = []
        for slide_index, slide in enumerate(deck.slides):
            for shape in slide.shapes:
                # 表紙写真だけは全面トリミング用の意図的なbleedを許可する。
                if slide_index > 0:
                    self.assertGreaterEqual(shape.left, 0)
                    self.assertGreaterEqual(shape.top, 0)
                    self.assertLessEqual(shape.left + shape.width, deck.slide_width)
                    self.assertLessEqual(shape.top + shape.height, deck.slide_height)

                frames = [shape.text_frame] if shape.has_text_frame else []
                if shape.has_table:
                    frames.extend(
                        cell.text_frame for row in shape.table.rows for cell in row.cells
                    )
                for frame in frames:
                    for paragraph in frame.paragraphs:
                        for run in paragraph.runs:
                            if not run.text.strip():
                                continue
                            text_runs.append(run)
                            self.assertIsNotNone(run.font.size)
                            if self._is_ten_point_chrome(shape, deck.slide_height):
                                self.assertEqual(10.0, run.font.size.pt)
                            elif self._is_nine_point_source(shape):
                                self.assertEqual(9.0, run.font.size.pt)
                            else:
                                self.assertGreaterEqual(run.font.size.pt, 12.0)

        self.assertTrue(text_runs)
        self.assertEqual({"Meiryo UI"}, {run.font.name for run in text_runs})

    def test_standard_deck_keeps_required_content_and_twelve_point_floor(self) -> None:
        """中計・費用を含む標準構成を、12pt・16:9・版面内で固定する。"""
        assessment = self._assessment()
        assessment["midterm_plan_analysis"] = {
            "plan_summary": (
                "公開計画の重点方針とAI実装の接点を整理し、"
                "優先テーマと実行条件を確認する。"
            ),
            "source": {
                "title": "公開中期経営計画2028",
                "url": "https://example.com/investors/management-plan/2028",
            },
            "management_targets": [{
                "label": "継続収益", "target": "120億円",
                "period": "FY2028", "comparison": "FY2025比+20%",
            }],
            "ai_alignment": [
                {
                    "plan_priority": f"重点方針{index}",
                    "ai_role": f"AI支援{index}",
                    "why_now": f"優先テーマ{index}の実行条件を確認する。",
                    "related_use_case": f"関連AIテーマ{index}",
                }
                for index in range(1, 4)
            ],
            "caveat": "事業効果は同一条件のPoCで測定し、投資判断へ接続する。",
        }
        assessment["ai_product_business_impact"] = (
            fixtures.DynamicPptxPageTests.ai_product_business_impact_fixture(assessment)
        )
        cost_estimate = fixtures.assessment_generator.cost_estimate_from_snapshot({
            "schema_version": "1", **copy.deepcopy(fixtures.TEST_COST_ESTIMATE),
        })
        self.assertIsNotNone(cost_estimate)
        with tempfile.TemporaryDirectory() as temporary_directory:
            path = Path(temporary_directory) / "standard-16-slides.pptx"
            presentation.create_pptx(
                assessment,
                DEFAULT_ARCHITECTURE_IMAGE,
                path,
                cost_estimate,
            )
            with ZipFile(path) as package:
                self.assertIsNone(package.testzip(), "PPTX package is corrupt")
            deck = Presentation(path)

        self.assertEqual(14, len(deck.slides))
        self.assertEqual(9 * deck.slide_width, 16 * deck.slide_height)
        self.assertEqual(12.0, presentation.MIN_PPTX_FONT_SIZE)
        self.assertEqual(10.0, presentation_common.PPTX_HEADER_FONT_SIZE)
        self.assertEqual(10.0, presentation.PPTX_FOOTER_FONT_SIZE)

        required_slide_text = {
            1: "OCI AI Use Case Assessment",
            2: "AIのビジネス価値を共同で構築",
            3: "中期経営計画とAI実装の整合性",
            4: "最新動向とAI活用",
            5: "ビジネスインパクト",
            6: "AIユースケース一覧",
            7: "15のAIユースケースから整理した AI技術アプローチの代表3テーマ",
            8: "AI技術テーマ 1／3",
            9: "AI技術テーマ 2／3",
            10: "AI技術テーマ 3／3",
            11: "代表3テーマを実現するOCI構成",
            12: "PoCにおけるOCI概算費用",
            13: "PoC実施・支援概要",
        }
        for slide_number, slide in enumerate(deck.slides, 1):
            slide_text = "\n".join(
                shape.text for shape in slide.shapes
                if shape.has_text_frame and shape.text.strip()
            )
            if slide_number in required_slide_text:
                self.assertIn(
                    required_slide_text[slide_number], slide_text,
                    f"slide {slide_number} lost its required narrative marker",
                )

            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    # Full-bleed photos may extend beyond the canvas, but must
                    # intersect it. Every editable/non-picture shape stays in-bounds.
                    self.assertLess(shape.left, deck.slide_width)
                    self.assertLess(shape.top, deck.slide_height)
                    self.assertGreater(shape.left + shape.width, 0)
                    self.assertGreater(shape.top + shape.height, 0)
                else:
                    self.assertGreaterEqual(
                        shape.left, 0,
                        f"slide {slide_number}: shape begins left of the canvas",
                    )
                    self.assertGreaterEqual(
                        shape.top, 0,
                        f"slide {slide_number}: shape begins above the canvas",
                    )
                    self.assertLessEqual(
                        shape.left + shape.width, deck.slide_width,
                        f"slide {slide_number}: shape exceeds the right edge",
                    )
                    self.assertLessEqual(
                        shape.top + shape.height, deck.slide_height,
                        f"slide {slide_number}: shape exceeds the bottom edge",
                    )

                frames = [shape.text_frame] if shape.has_text_frame else []
                if shape.has_table:
                    frames.extend(
                        cell.text_frame for row in shape.table.rows for cell in row.cells
                    )
                for frame in frames:
                    for paragraph in frame.paragraphs:
                        for run in paragraph.runs:
                            if not run.text.strip():
                                continue
                            self.assertIsNotNone(
                                run.font.size,
                                f"slide {slide_number}: editable text has no explicit size: {run.text!r}",
                            )
                            if self._is_ten_point_chrome(shape, deck.slide_height):
                                self.assertEqual(
                                    run.font.size.pt, 10.0,
                                    f"slide {slide_number}: common header/footer must be 10pt: {run.text!r}",
                                )
                            elif self._is_nine_point_source(shape):
                                self.assertEqual(
                                    run.font.size.pt, 9.0,
                                    f"slide {slide_number}: industry source note must be 9pt: {run.text!r}",
                                )
                            else:
                                self.assertGreaterEqual(
                                    run.font.size.pt, 12.0,
                                    f"slide {slide_number}: editable content below 12pt: {run.text!r}",
                                )

        closing = deck.slides[-1]
        self.assertTrue(
            any(shape.shape_type == MSO_SHAPE_TYPE.PICTURE for shape in closing.shapes),
            "the final Oracle closing slide must retain its logo artwork",
        )


if __name__ == "__main__":
    unittest.main()
