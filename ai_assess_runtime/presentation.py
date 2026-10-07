"""AIアセスメントの編集可能な16:9 PowerPoint描画層。"""

from __future__ import annotations

import copy
from pathlib import Path
import re

from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
try:
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.util import Pt
except ImportError:
    MSO_ANCHOR = MSO_AUTO_SIZE = PP_ALIGN = Pt = None  # type: ignore[assignment]

from ai_assess_runtime.deck_plan import SlideSpec, build_default_slide_plan
from ai_assess_runtime.display_text import normalize_public_source_title
from ai_assess_runtime.paths import (
    ASSESSMENT_INTRO_IMAGE,
    COVER_BACKGROUND_IMAGE,
    ORACLE_LOGO_IMAGE,
)
from ai_assess_runtime.poc_contract import (
    _front_text,
    _matching_source_poc,
    assessment_story_for,
    consultative_page_lead,
    consulting_front_matter_for,
    display_theme_name,
    normalize_poc_priority_decision,
    poc_logic_details_for,
    poc_technical_design_for,
    poc_selection_scorecard_for,
    priority_decision_for,
    priority_pocs_for,
    technical_modality_for,
)
from ai_assess_runtime.pptx_canvas import (
    MIN_PPTX_FONT_SIZE,
    PPTX_SOURCE_NOTE_SHAPE_NAME,
    PptCanvas,
    _plain_text,
)
from ai_assess_runtime.presentation_common import (
    ASSESSMENT_SUBSECTION_LABELS,
    BUSINESS_VALUE_CARD_ACCENT_HEX,
    FONT_BOLD_PATH_CANDIDATES,
    FONT_PATH_CANDIDATES,
    PPTX_FOOTER_FONT_SIZE,
    PPTX_HEADER_FONT_SIZE,
    PPTX_WIDESCREEN_HEIGHT,
    PPTX_WIDESCREEN_WIDTH,
    STANDARD_HEADER_HEIGHT,
    STANDARD_TITLE_SIZE,
    STANDARD_TITLE_X,
    STANDARD_TITLE_Y_OFFSET,
    USE_CASE_LIST_BODY_COLOR,
    USE_CASE_LIST_CELL_MARGIN_MM,
    USE_CASE_LIST_COLUMN_WIDTHS,
    USE_CASE_LIST_HEADER_HEIGHT,
    USE_CASE_LIST_PRIORITY_ICON_GAP,
    USE_CASE_LIST_PRIORITY_ICON_IMAGE,
    USE_CASE_LIST_PRIORITY_ICON_SIZE,
    USE_CASE_LIST_ROW_HEIGHT,
    USE_CASE_LIST_TABLE_BOTTOM,
    USE_CASE_LIST_TABLE_TOP,
    USE_CASE_LIST_TEXT_SAFETY_FACTOR,
    USE_CASE_LIST_WRAP_TOLERANCE,
    USE_CASES_PER_LIST_SLIDE,
    WIDE_CONTENT_LEFT,
    WIDE_CONTENT_RIGHT,
    WIDE_EXPLANATORY_LEAD_TOP,
    WIDE_TABLE_MIN_ROW_HEIGHT,
    _executive_evidence_header,
    assessment_subsection_label,
    basis_label,
    draw_footer,
    draw_number_badge,
    draw_paragraph,
    draw_pptx_base_header,
    draw_story_source_note,
    estimate_pptx_paragraph_height,
    estimate_wide_table_lines,
    poc_design_subsection_label,
    register_japanese_font,
    wide_content_width,
    wide_table_row_height,
)
from ai_assess_runtime.quantitative_contract import (
    ai_product_business_impact_for,
    normalize_ai_product_business_impact_candidate,
    normalize_poc_measurement_design,
    poc_measurement_design_for,
    poc_measurement_design_items,
)

# 優先PoCを扱う各ページで共有する、テーマ対応用のアクセント色。順位や評価の
# 優劣を表す色ではなく、一覧・構成・詳細の同じテーマを追えるようにする。
POC_THEME_ACCENT_HEX = ("#467653", "#C74634", "#367A9B")


def _draw_pptx_cost_estimate_page(deck: PptCanvas, page_width: float, page_height: float,
                                  cost_estimate: dict, page_number: int) -> None:
    """16:9・14ptで費用表と前提を安全に収めるPPTX専用レイアウト。"""
    dark = colors.HexColor("#1D252C")
    muted = colors.HexColor("#4C5961")
    green = colors.HexColor("#467653")
    light = colors.HexColor("#EEF2F5")
    # 顧客向けにはOCI利用料の項目・金額・利用前提だけを示す。価格取得元は
    # 監査用の凍結JSONへ保持し、スライドでは前提欄へ横幅を優先配分する。
    table_widths = [55 * mm, 34 * mm, wide_content_width(page_width) - 89 * mm]
    def exact_yen(value: int | float) -> str:
        return f"{round(value):,}円"

    rows = [["項目", "月額（円）", "前提・備考"]]
    for line in cost_estimate["lines"]:
        assumption = _clean_pptx_display_text(line.assumption)
        if "入力100 token" in assumption:
            assumption = (
                "月間10,000リクエストの仮置き。RAGの検索文脈を含む入出力tokenを実測し再見積り"
            )
        rows.append([
            _pptx_summary_text(line.name, 28),
            exact_yen(line.monthly_jpy),
            assumption,
        ])
    rows.append([
        "合計",
        exact_yen(cost_estimate['total_monthly_jpy']),
        f"10週間では{exact_yen(cost_estimate['total_monthly_jpy'] * 2.5)}。OCI利用料のみの単純換算",
    ])
    # 前提・備考は切り詰めず、14ptで必要な行数に応じて行高を確保する。
    row_heights = [9.5 * mm] + [
        wide_table_row_height(row, table_widths, font_size=14.0, minimum=11.5 * mm)
        for row in rows[1:]
    ]

    draw_story_header(deck, page_width, page_height, "PoCにおけるOCI概算費用（OCI月額利用料）", section=assessment_subsection_label("technical"))
    subtitle = ParagraphStyle(
        "pptx_cost_subtitle", fontName="AssessmentJapanese", fontSize=14.0,
        leading=16.6, textColor=muted, wordWrap="CJK",
    )
    line_count = len(cost_estimate.get("lines", []))
    api_count = int(cost_estimate.get("api_priced_count") or 0)
    if line_count and api_count == line_count:
        pricing_basis = "OCI List Pricing APIベース"
    elif api_count:
        pricing_basis = "OCI List Pricing API価格＋前提値"
    else:
        pricing_basis = "現行前提値ベース（SKU確定後に再計算）"
    draw_paragraph(
        deck,
        f"{pricing_basis}の月額概算です。対象はOCI利用料のみで、税、データ転送量、契約割引は含みません。"
        "採用モデル、リージョン、SKU、利用量を確定した時点で同じ計算表を更新します。",
        subtitle, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP + 2 * mm, wide_content_width(page_width),
    )

    # 経営判断で最初に見る合計と対象範囲を、表から独立した強調帯にする。
    deck.setFillColor(colors.HexColor("#EEF4F0")); deck.roundRect(
        WIDE_CONTENT_LEFT, 127 * mm, wide_content_width(page_width), 18 * mm, 2.5 * mm, stroke=0, fill=1,
    )
    def aligned_text(text, x, y, width, height, size, color, bold=False, name=""):
        box = deck.slide.shapes.add_textbox(deck._x(x), deck._y(y + height),
                                           deck._width(width), deck._height(height))
        box.name = name or "AI_ASSESS_COST_TEXT"
        frame = box.text_frame
        frame.clear(); frame.word_wrap = True; frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = frame.paragraphs[0]
        paragraph.space_before = paragraph.space_after = Pt(0)
        run = paragraph.add_run(); run.text = text
        run.font.name = "Meiryo UI"; run.font.size = Pt(size)
        run.font.bold = bold; run.font.color.rgb = deck._rgb(color)
        return box

    for text, offset, width, size, color, name in (
        ("OCI月額利用料 合計", 6, 49, 14, green, "AI_ASSESS_COST_TOTAL_LABEL"),
        (exact_yen(cost_estimate['total_monthly_jpy']) + "／月", 57, 58, 18, dark, "AI_ASSESS_COST_TOTAL_AMOUNT"),
        (f"{pricing_basis} ｜ {cost_estimate['currency']}・税抜", 119,
         wide_content_width(page_width) / mm - 124, 14, muted, "AI_ASSESS_COST_TOTAL_BASIS"),
    ):
        aligned_text(text, WIDE_CONTENT_LEFT + offset * mm, 127 * mm,
                     width * mm, 18 * mm, size, color, name=name, bold=name != "AI_ASSESS_COST_TOTAL_BASIS")

    # 上端を固定して、表と下段の注記を独立した編集可能な領域へ分ける。
    table_top = 125 * mm
    table_y = table_top - sum(row_heights)
    deck.add_table(rows, WIDE_CONTENT_LEFT, table_y, table_widths, row_heights, green,
                  [light] * max(1, len(rows) - 2) + [colors.HexColor("#E6EBF3")], 14.0,
                  cell_margin_mm=1.5)

    # Anchor all notes below the actual table bottom, with a visible gap.
    notes = (
        ("価格前提", "税抜概算。地域・SKU・価格時点・割引・無料枠・転送費を確認します。"),
        ("稼働前提", "Compute／ADBは月744時間稼働。停止時は従量課金を再計算します。"),
        ("再計算条件", "10週間＝月額×2.5か月。初回Embedding・実装支援・転送費は別途確認。"),
    )
    note_h, gap = 9 * mm, 1 * mm
    note_top = table_y - 3 * mm
    if note_top - len(notes) * (note_h + gap) < 13 * mm:
        raise ValueError("OCI cost rows exceed the footer-safe area; shorten assumptions or reduce line items before rendering")
    note_colors = ["#EEF4F0", "#EDF4F8", "#FFF4E9"]
    note_accents = [green, colors.HexColor("#367A9B"), colors.HexColor("#C74634")]
    for index, (heading, text) in enumerate(notes):
        x = WIDE_CONTENT_LEFT; y = note_top - note_h - index * (note_h + gap)
        deck.setFillColor(colors.HexColor(note_colors[index]))
        deck.roundRect(x, y, wide_content_width(page_width), note_h, 1.2 * mm, stroke=0, fill=1)
        deck.slide.shapes[-1].name = f"AI_ASSESS_COST_NOTE_BAND_{index}"
        deck.setFillColor(note_accents[index]); deck.rect(x, y, 33 * mm, note_h, stroke=0, fill=1)
        aligned_text(heading, x + 4 * mm, y, 27 * mm, note_h, 14, colors.white, True,
                     f"AI_ASSESS_COST_NOTE_LABEL_{index}")
        aligned_text(text, x + 38 * mm, y, wide_content_width(page_width) - 43 * mm,
                     note_h, 14, dark, name=f"AI_ASSESS_COST_NOTE_TEXT_{index}")
    draw_footer(deck, page_width, page_number)


def draw_cost_estimate_page(deck: PptCanvas, page_width: float, page_height: float,
                            cost_estimate: dict, page_number: int) -> None:
    """構成図の次に表示する、PoC向け月額概算費用PPTXページ。"""
    _draw_pptx_cost_estimate_page(deck, page_width, page_height, cost_estimate, page_number)


def _clean_pptx_display_text(value: object) -> str:
    """PPTXへ渡す文章の空白だけを正規化し、原文は削らず保持する。"""
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return re.sub(r"\s+", " ", text).strip()


def _pptx_summary_text(value: object, limit: int) -> str:
    """表示領域の文字数上限で原文を削らない、互換用の正規化関数。

    ``limit`` は既存の呼び出し互換として受け取る。枠に収める処理は描画層で
    折返し・行高・ページ分割に委ね、ここでは文・節・文字のいずれも落とさない。
    """
    _ = limit
    return _clean_pptx_display_text(value)


def _pptx_reason_summary(value: object, *, limit: int = 30) -> str:
    """選定理由・整合理由を、表示用に削らずそのまま返す。"""
    _ = limit
    return _clean_pptx_display_text(value)


def _pptx_formula_text(value: object, target: object, limit: int) -> str:
    """算定式を表示領域の上限で途中省略せず、原文のまま返す。"""
    _ = target, limit
    return _clean_pptx_display_text(value)


def _pptx_kpi_label(value: object, limit: int = 12) -> str:
    """KPIラベルを表示領域の上限で短縮せず、原文のまま返す。"""
    _ = limit
    return _clean_pptx_display_text(value)


def _pptx_period_label(value: object, limit: int = 12) -> str:
    """測定期間を表示領域の上限で短縮せず、原文のまま返す。"""
    _ = limit
    return _clean_pptx_display_text(value)


def _pptx_action_summary(value: object, *, limit: int = 18) -> str:
    """開始条件・改善アクションを表示領域の上限で短縮せず返す。"""
    _ = limit
    return _clean_pptx_display_text(value)


def _pptx_data_readiness_label(score: object) -> str:
    """PoCデータ準備度の点数を、顧客が判断できる短い状態説明へ変換する。"""
    if isinstance(score, (int, float)):
        numeric = int(score)
        if numeric >= 5:
            return "範囲・品質・権限・正解を確認済"
        if numeric >= 3:
            return "中核データ特定済・加工確認要"
        return "中核データの追加確認が必要"
    return "評価条件を確認"


def _pptx_noun_list_summary(value: object, *, limit: int = 36) -> str:
    """対象データ等を表示領域の上限で圧縮せず、原文のまま返す。"""
    _ = limit
    return _clean_pptx_display_text(value)


def _pptx_use_case_name_summary(value: object) -> str:
    """一覧用のユースケース名を、語の置換や短文化なしで返す。"""
    return _clean_pptx_display_text(value)


def _pptx_business_challenge_summary(value: object, *, limit: int = 68) -> str:
    """PoCの課題・価値を表示領域の上限で短文化せず返す。"""
    _ = limit
    return _clean_pptx_display_text(value)


def draw_fixed_architecture_reference_page(deck: PptCanvas, page_width: float, page_height: float,
                                          assessment: dict, architecture_image: Path, page_number: int) -> None:
    """固定OCI構成図と、優先3ユースケースの実現方法を一枚に結ぶ。"""
    if not architecture_image.is_file():
        raise FileNotFoundError(f"固定構成図が見つかりません: {architecture_image}")
    draw_pptx_base_header(
        deck, page_width, page_height, "代表3テーマを実現するOCI構成",
        assessment_subsection_label("technical"),
    )
    lead = ParagraphStyle(
        "fixed_architecture_lead", fontName="AssessmentJapanese", fontSize=14.0,
        leading=16.2, textColor=colors.HexColor("#4C5961"), wordWrap="CJK",
    )
    draw_paragraph(
        deck,
        consultative_page_lead(
            "業務連携・アプリケーション層、Autonomous AI Database、OCI Generative AIを組み合わせ、"
            "\n代表技術テーマ別の判断支援を、共通のデータ・統制基盤上で実現します。",
            title="代表3テーマを実現するOCI構成",
        ),
        lead, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP, wide_content_width(page_width),
    )
    reader = ImageReader(str(architecture_image))
    image_width, image_height = reader.getSize()
    # 左側は固定構成図を十分な大きさで表示し、右側には案件ごとの優先3テーマを
    # 同じ構成へ接続する。固定図は基盤、右のカードは案件ごとの使い分けを示す。
    left_x, left_width, available_height = WIDE_CONTENT_LEFT, 165 * mm, 103 * mm
    scale = min(left_width / image_width, available_height / image_height)
    draw_width, draw_height = image_width * scale, image_height * scale
    image_x = left_x + (left_width - draw_width) / 2
    image_y = 36 * mm + (available_height - draw_height) / 2
    deck.drawImage(reader, image_x, image_y, draw_width, draw_height, preserveAspectRatio=True, mask="auto")

    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    accents = [colors.HexColor(value) for value in POC_THEME_ACCENT_HEX]
    fills = [colors.HexColor("#EEF4F0"), colors.HexColor("#FFF4E9"), colors.HexColor("#EDF4F8")]
    right_x = left_x + left_width + 7 * mm
    right_w = page_width - right_x - WIDE_CONTENT_RIGHT
    deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 14.0)
    deck.drawString(right_x, 140 * mm, "代表技術テーマごとの実現方法")
    detail_rows = poc_logic_details_for(assessment)
    priority_pocs = priority_pocs_for(assessment)[:3]
    card_style = ParagraphStyle(
        "fixed_architecture_card", fontName="AssessmentJapanese", fontSize=14.0,
        leading=16.2, textColor=muted, wordWrap="CJK",
    )
    for index, poc in enumerate(priority_pocs):
        detail = next((item for item in detail_rows if _matching_source_poc(item, poc)), {})
        modality = technical_modality_for(poc, detail)
        # 14ptの日本語本文が3行になっても次のカード背景へはみ出さない高さを確保する。
        card_h, card_y = 33 * mm, 101 * mm - index * 33.5 * mm
        accent, fill = accents[index], fills[index]
        deck.setFillColor(fill); deck.roundRect(right_x, card_y, right_w, card_h, 2.0 * mm, stroke=0, fill=1)
        draw_number_badge(deck, right_x + 8 * mm, card_y + card_h - 9 * mm, 4.3 * mm, index + 1, accent)
        theme = _pptx_summary_text(display_theme_name(detail, poc), 40) or "AI技術テーマ"
        deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 14.0)
        title_bottom = draw_paragraph(
            deck, theme,
            ParagraphStyle(
                f"fixed_architecture_theme_{index}", fontName="AssessmentJapaneseBold",
                fontSize=14.0, leading=16.2, textColor=dark, wordWrap="CJK",
            ),
            right_x + 16 * mm, card_y + card_h - 7.4 * mm, right_w - 21 * mm,
        )
        # 右カードは「方式＋実現価値」の一文に限定する。原文を末尾記号で
        # 切るのではなく、JSONに保持した課題文から完結する節を選ぶ。
        detail_text = _pptx_business_challenge_summary(
            detail.get("business_challenge") or poc.get("reason"), limit=30,
        )
        draw_paragraph(
            deck, f"{modality['label']}：{detail_text}", card_style,
            right_x + 5 * mm, max(card_y + 4 * mm, title_bottom - 1.0 * mm), right_w - 10 * mm,
        )

    foundation_y, foundation_h = 12 * mm, 18 * mm
    deck.setFillColor(colors.HexColor("#E5F0EA")); deck.roundRect(
        WIDE_CONTENT_LEFT, foundation_y, wide_content_width(page_width), foundation_h, 2 * mm, stroke=0, fill=1,
    )
    deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 14.0)
    deck.drawString(WIDE_CONTENT_LEFT + 5 * mm, foundation_y + 10.5 * mm, "共通基盤の考え方")
    foundation = (
        "業務画面・APIで候補と根拠を返し、Autonomous AI Databaseでデータ・検索・分析・監査記録を管理します。"
        "図中のモデル名は構成例であり、採用モデル、リージョン提供可否、性能、単価はPoC開始時に再確認します。"
    )
    draw_paragraph(
        deck, _pptx_summary_text(foundation, 80), card_style,
        WIDE_CONTENT_LEFT + 52 * mm, foundation_y + 11.2 * mm,
        wide_content_width(page_width) - 57 * mm,
    )
    draw_footer(deck, page_width, page_number)


def draw_deployment_options_page(deck: PptCanvas, page_width: float, page_height: float,
                                 assessment: dict, page_number: int) -> None:
    """既存DB内実行とADBオフロードを、選択条件付きで並列に示す。"""
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    accents = [colors.HexColor("#467653"), colors.HexColor("#367A9B"), colors.HexColor("#C74634")]
    body = ParagraphStyle("deployment_body", fontName="AssessmentJapanese", fontSize=14.0,
                          leading=16.2, textColor=muted, wordWrap="CJK")
    bold = ParagraphStyle("deployment_bold", fontName="AssessmentJapaneseBold", fontSize=14.0,
                          leading=16.2, textColor=dark, wordWrap="CJK")
    draw_pptx_base_header(deck, page_width, page_height, "既存システムとAI・OCIの役割分担",
                          assessment_subsection_label("technical"))
    draw_paragraph(
        deck,
        "既存システムの確定処理を維持し、AIは候補・スコア・根拠を返す構成を基本案とします。"
        "実装位置は、既存Oracle Databaseの版数・負荷・権限とPoCデータの扱いを確認して選択できます。",
        body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP, wide_content_width(page_width),
    )
    cards = [
        ("既存DB内で実行", "利用可能なAI・ML機能が既存DBの版数と構成に合う場合、データ移動を抑えてDB内で処理します。", "確認：DB版数・オプション・処理負荷"),
        ("検証データをADBへ複製", "本番更新から分離した検証データをAutonomous AI Databaseへ持ち込み、予測・検索・監査を試します。", "確認：複製範囲・更新頻度・匿名化"),
        ("参照・API連携でオフロード", "Database Linkは参照経路、APIは入出力経路として使い分け、AI処理を外側へ分離します。", "確認：権限・遅延・障害時の戻し方"),
    ]
    gap = 6 * mm
    card_w = (wide_content_width(page_width) - gap * 2) / 3
    card_y, card_h = 54 * mm, 72 * mm
    for index, (title, description, condition) in enumerate(cards):
        x = WIDE_CONTENT_LEFT + index * (card_w + gap)
        accent = accents[index]
        deck.setFillColor(colors.white); deck.setStrokeColor(colors.HexColor("#D8DEE2")); deck.setLineWidth(0.8)
        deck.roundRect(x, card_y, card_w, card_h, 3 * mm, stroke=1, fill=1)
        deck.setFillColor(accent); deck.rect(x, card_y + card_h - 3 * mm, card_w, 3 * mm, stroke=0, fill=1)
        draw_number_badge(deck, x + 11 * mm, card_y + card_h - 14 * mm, 6 * mm, index + 1, accent, 14.0)
        draw_paragraph(deck, title, bold, x + 21 * mm, card_y + card_h - 9 * mm, card_w - 27 * mm)
        draw_paragraph(deck, description, body, x + 7 * mm, card_y + card_h - 30 * mm, card_w - 14 * mm)
        deck.setFillColor(colors.HexColor("#F6F7F8")); deck.roundRect(
            x + 6 * mm, card_y + 7 * mm, card_w - 12 * mm, 17 * mm, 2 * mm, stroke=0, fill=1,
        )
        draw_paragraph(deck, condition, bold, x + 9 * mm, card_y + 19 * mm, card_w - 18 * mm)
    deck.setFillColor(colors.HexColor("#EEF4F0")); deck.roundRect(
        WIDE_CONTENT_LEFT, 22 * mm, wide_content_width(page_width), 22 * mm, 2.5 * mm, stroke=0, fill=1,
    )
    deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 14.0)
    deck.drawString(WIDE_CONTENT_LEFT + 6 * mm, 36 * mm, "PoCでの基本案")
    draw_paragraph(
        deck,
        "最初は読取中心の検証データまたは参照連携から始め、結果は既存画面やAPIへ返します。"
        "本番への書込み範囲は、精度・運用・統制を確認した後に具体化します。",
        body, WIDE_CONTENT_LEFT + 40 * mm, 37 * mm, wide_content_width(page_width) - 46 * mm,
    )
    draw_footer(deck, page_width, page_number)


def draw_poc_support_icon(deck: PptCanvas, kind: str, x: float, y: float, color: colors.Color) -> None:
    """PoC支援スライド用の簡易ベクトルアイコン。x, yは左下。"""
    deck.setStrokeColor(color); deck.setFillColor(color); deck.setLineWidth(1.4)
    if kind == "checklist":
        deck.roundRect(x + 3 * mm, y, 13 * mm, 18 * mm, 1.2 * mm, stroke=1, fill=0)
        deck.roundRect(x + 6.5 * mm, y + 16 * mm, 6 * mm, 3 * mm, 1 * mm, stroke=1, fill=1)
        for offset in (12, 8, 4):
            deck.line(x + 6 * mm, y + offset * mm / 1.5, x + 8 * mm, y + (offset - 1.5) * mm / 1.5)
            deck.line(x + 8 * mm, y + (offset - 1.5) * mm / 1.5, x + 10.5 * mm, y + (offset + 2) * mm / 1.5)
            deck.line(x + 12 * mm, y + offset * mm / 1.5, x + 15 * mm, y + offset * mm / 1.5)
    elif kind == "architecture":
        deck.roundRect(x + 7 * mm, y + 13 * mm, 6 * mm, 5 * mm, 0.8 * mm, stroke=1, fill=1)
        for box_x in (2, 10, 18):
            deck.roundRect(x + box_x * mm, y + 2 * mm, 5 * mm, 4 * mm, 0.8 * mm, stroke=1, fill=0)
        deck.line(x + 10 * mm, y + 13 * mm, x + 10 * mm, y + 8 * mm)
        deck.line(x + 4.5 * mm, y + 8 * mm, x + 20.5 * mm, y + 8 * mm)
        for endpoint in (4.5, 12.5, 20.5):
            deck.line(x + endpoint * mm, y + 8 * mm, x + endpoint * mm, y + 6 * mm)
    elif kind == "team":
        deck.circle(x + 11 * mm, y + 13 * mm, 3.4 * mm, stroke=1, fill=1)
        deck.circle(x + 4.5 * mm, y + 11 * mm, 2.6 * mm, stroke=1, fill=1)
        deck.circle(x + 17.5 * mm, y + 11 * mm, 2.6 * mm, stroke=1, fill=1)
        deck.roundRect(x + 6 * mm, y + 2 * mm, 10 * mm, 6 * mm, 3 * mm, stroke=1, fill=1)
        deck.roundRect(x, y + 2 * mm, 7 * mm, 4 * mm, 2 * mm, stroke=1, fill=1)
        deck.roundRect(x + 15 * mm, y + 2 * mm, 7 * mm, 4 * mm, 2 * mm, stroke=1, fill=1)
    elif kind == "cloud":
        deck.circle(x + 7 * mm, y + 9 * mm, 4 * mm, stroke=1, fill=0)
        deck.circle(x + 12 * mm, y + 11 * mm, 5 * mm, stroke=1, fill=0)
        deck.circle(x + 17 * mm, y + 8.5 * mm, 3.5 * mm, stroke=1, fill=0)
        deck.line(x + 4 * mm, y + 5 * mm, x + 20 * mm, y + 5 * mm)
    elif kind == "database":
        deck.ellipse(x + 4 * mm, y + 13 * mm, x + 18 * mm, y + 18 * mm, stroke=1, fill=0)
        deck.line(x + 4 * mm, y + 15.5 * mm, x + 4 * mm, y + 4 * mm)
        deck.line(x + 18 * mm, y + 15.5 * mm, x + 18 * mm, y + 4 * mm)
        deck.ellipse(x + 4 * mm, y + 1.5 * mm, x + 18 * mm, y + 6.5 * mm, stroke=1, fill=0)
    else:  # chart
        for offset, height in ((2, 6), (8, 11), (14, 16)):
            deck.rect(x + offset * mm, y + 2 * mm, 4 * mm, height * mm / 2, stroke=0, fill=1)


def _draw_pptx_poc_support_page(deck: PptCanvas, page_width: float, page_height: float,
                                 assessment: dict, page_number: int) -> None:
    """PoC支援概要を16:9・本文12ptの安全領域へ再配置する。"""
    dark, navy, teal = colors.HexColor("#1D252C"), colors.HexColor("#124883"), colors.HexColor("#087A78")
    orange, purple, green = colors.HexColor("#D46A00"), colors.HexColor("#6941A5"), colors.HexColor("#467653")
    light, muted, border = colors.HexColor("#F6F9FC"), colors.HexColor("#4C5961"), colors.HexColor("#9ABBE0")
    body = ParagraphStyle("support_wide_body", fontName="AssessmentJapanese", fontSize=14.0,
                          leading=16.4, textColor=dark, wordWrap="CJK")
    title_style = ParagraphStyle("support_wide_title", fontName="AssessmentJapaneseBold", fontSize=14.0,
                                 leading=16.4, textColor=dark, wordWrap="CJK")
    draw_story_header(deck, page_width, page_height, f"{assessment['service_name']}向け PoC実施・支援概要", section=assessment_subsection_label("technical"))

    def centered_text(text: str, x: float, y: float, width: float, height: float, *,
                      color: colors.Color, bold: bool = False, align=PP_ALIGN.LEFT) -> None:
        """帯・表セルの文字を、PowerPoint上でも垂直中央に揃える。"""
        box = deck.slide.shapes.add_textbox(deck._x(x), deck._y(y + height), deck._width(width), deck._height(height))
        frame = box.text_frame; frame.clear(); frame.word_wrap = True
        frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        frame.margin_left = frame.margin_right = 0
        frame.margin_top = frame.margin_bottom = 0
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = frame.paragraphs[0]; paragraph.alignment = align
        run = paragraph.add_run(); run.text = text
        run.font.name = "Meiryo UI"; run.font.size = Pt(14.0); run.font.bold = bold
        run.font.color.rgb = deck._rgb(color)

    def section_label(number: str, label: str, y: float) -> None:
        deck.setFillColor(navy); deck.roundRect(WIDE_CONTENT_LEFT, y, 54 * mm, 8 * mm, 1.5 * mm, stroke=0, fill=1)
        centered_text(f"{number}  {label}", WIDE_CONTENT_LEFT + 4 * mm, y, 46 * mm, 8 * mm,
                      color=colors.white, bold=True)

    # 1. PoC立上げ
    # 見出し直下の余白を抑え、カード内の12pt本文にも十分な高さを確保する。
    section_label("1", "PoC立上げ", 146 * mm)
    card_gap, card_y, card_h = 5 * mm, 108 * mm, 34 * mm
    card_w = (wide_content_width(page_width) - card_gap * 2) / 3
    top_cards = [
        ("checklist", "最初に扱うテーマの具体化", "代表3テーマから1テーマを選び、対象業務・データ・評価基準を確認します。"),
        ("architecture", "PoCアーキテクチャの確定", "業務連携・アプリ層、Autonomous AI Database、生成AIの連携方式を決定します。"),
        ("team", "体制・進め方の合意", "お客様チームがPoCを主体的に実施し、Oracleが設計レビューとQAを支援します。"),
    ]
    for index, (icon, title, description) in enumerate(top_cards):
        x = WIDE_CONTENT_LEFT + index * (card_w + card_gap)
        deck.setFillColor(light); deck.roundRect(x, card_y, card_w, card_h, 2 * mm, stroke=0, fill=1)
        deck.setStrokeColor(border); deck.setLineWidth(0.8); deck.roundRect(x, card_y, card_w, card_h, 2 * mm, stroke=1, fill=0)
        # アイコンの描画原点を、スケール後の高さからカード中央へ合わせる。
        icon_scale, icon_height = 0.64, 18 * mm * 0.64
        icon_y = card_y + (card_h - icon_height) / 2
        deck.saveState(); deck.translate(x + 5 * mm, icon_y); deck.scale(icon_scale, icon_scale)
        draw_poc_support_icon(deck, icon, 0, 0, navy)
        deck.restoreState()
        draw_paragraph(deck, title, title_style, x + 22 * mm, card_y + card_h - 8 * mm, card_w - 27 * mm)
        draw_paragraph(
            deck, _pptx_summary_text(description, 54), body,
            x + 22 * mm, card_y + card_h - 14.5 * mm, card_w - 27 * mm,
        )

    # 2. 体制・役割
    section_label("2", "体制・役割", 99 * mm)
    # 連携ラベルを2枚のカード上へ重ねないよう、中央に明示的な接続帯を取る。
    role_gap, role_y, role_h = 24 * mm, 69 * mm, 24 * mm
    role_w = (wide_content_width(page_width) - role_gap) / 2
    role_cards = [
        (WIDE_CONTENT_LEFT, navy, "お客様プロジェクトチーム（PoC主体）", "製品責任者、業務担当、DB・アプリ担当、セキュリティ、評価責任者で対象・データ・結果を確認", "team"),
        (WIDE_CONTENT_LEFT + role_w + role_gap, teal, "Oracle（PoC支援）", "構成レビュー、技術QA、評価設計、定例会を支援し、お客様の判断材料を整理", "oracle_logo"),
    ]
    logo_reader = ImageReader(str(ORACLE_LOGO_IMAGE))
    logo_width, logo_height = logo_reader.getSize()
    for x, color, title, description, icon in role_cards:
        deck.setFillColor(light); deck.roundRect(x, role_y, role_w, role_h, 2 * mm, stroke=0, fill=1)
        deck.setStrokeColor(border); deck.setLineWidth(0.8); deck.roundRect(x, role_y, role_w, role_h, 2 * mm, stroke=1, fill=0)
        deck.setFillColor(color); deck.roundRect(x + 18 * mm, role_y + role_h - 6 * mm, role_w - 36 * mm, 7 * mm, 1.6 * mm, stroke=0, fill=1)
        centered_text(title, x + 18 * mm, role_y + role_h - 6 * mm, role_w - 36 * mm, 7 * mm,
                      color=colors.white, bold=True, align=PP_ALIGN.CENTER)
        if icon == "oracle_logo":
            draw_width = 23 * mm; draw_height = draw_width * logo_height / logo_width
            deck.drawImage(logo_reader, x + 5 * mm, role_y + 6 * mm, width=draw_width, height=draw_height, preserveAspectRatio=True, mask="auto")
        else:
            deck.saveState(); deck.translate(x + 5 * mm, role_y + 3 * mm); deck.scale(0.62, 0.62)
            draw_poc_support_icon(deck, icon, 0, 0, color)
            deck.restoreState()
        draw_paragraph(
            deck, _pptx_summary_text(description, 46), body,
            x + 31 * mm, role_y + role_h - 11.5 * mm, role_w - 36 * mm,
        )
    arrow_x = WIDE_CONTENT_LEFT + role_w + 1.2 * mm
    deck.setStrokeColor(muted); deck.setLineWidth(1.2); deck.line(arrow_x, role_y + 10 * mm, arrow_x + role_gap - 2.4 * mm, role_y + 10 * mm)
    deck.setFillColor(muted); deck.setFont("AssessmentJapaneseBold", 12.0)
    deck.drawCentredString(arrow_x + (role_gap - 2.4 * mm) / 2, role_y + 14.5 * mm, "週次連携")
    deck.drawCentredString(arrow_x + (role_gap - 2.4 * mm) / 2, role_y + 3.5 * mm, "QA支援")

    # 3. スケジュール
    section_label("3", "スケジュール（目安）", 57 * mm)
    table_x, table_y, table_w, table_h = WIDE_CONTENT_LEFT, 16 * mm, wide_content_width(page_width), 34 * mm
    phase_w, task_w, week_w = 43 * mm, 65 * mm, (table_w - 108 * mm) / 10
    header_h, row_h = 7 * mm, (table_h - 7 * mm) / 4
    deck.setFillColor(colors.HexColor("#F9FAFB")); deck.rect(table_x, table_y, table_w, table_h, stroke=0, fill=1)
    deck.setStrokeColor(colors.HexColor("#C9D4DF")); deck.setLineWidth(0.5); deck.rect(table_x, table_y, table_w, table_h, stroke=1, fill=0)
    for boundary in (phase_w, phase_w + task_w): deck.line(table_x + boundary, table_y, table_x + boundary, table_y + table_h)
    for week in range(11):
        x = table_x + phase_w + task_w + week * week_w; deck.line(x, table_y, x, table_y + table_h)
    deck.line(table_x, table_y + table_h - header_h, table_x + table_w, table_y + table_h - header_h)
    for row in range(1, 4):
        y = table_y + table_h - header_h - row * row_h; deck.line(table_x, y, table_x + table_w, y)
    header_y = table_y + table_h - header_h
    centered_text("フェーズ", table_x, header_y, phase_w, header_h, color=dark, bold=True, align=PP_ALIGN.CENTER)
    centered_text("主なタスク", table_x + phase_w, header_y, task_w, header_h, color=dark, bold=True, align=PP_ALIGN.CENTER)
    for week in range(10):
        centered_text(f"W{week + 1}", table_x + phase_w + task_w + week_w * week, header_y, week_w, header_h,
                      color=dark, bold=True, align=PP_ALIGN.CENTER)
    schedule = [("1. 計画・設計", "対象・比較条件の設計", navy, 0, 2), ("2. データ準備", "環境・データ準備", teal, 1, 2), ("3. 構築・検証", "実装・代表業務で検証", purple, 3, 5), ("4. 評価・整理", "精度・費用・運用整理", orange, 8, 2)]
    for row, (phase, task, color, start, duration) in enumerate(schedule):
        center_y = table_y + table_h - header_h - row_h * (row + 0.5)
        deck.setFillColor(color); deck.circle(table_x + 6 * mm, center_y, 2.4 * mm, stroke=0, fill=1)
        row_y = table_y + table_h - header_h - row_h * (row + 1)
        centered_text(phase, table_x + 11 * mm, row_y, phase_w - 13 * mm, row_h, color=dark, bold=True)
        centered_text(task, table_x + phase_w + 4 * mm, row_y, task_w - 8 * mm, row_h, color=dark)
        bar_x = table_x + phase_w + task_w + start * week_w + 1.5 * mm; bar_width = duration * week_w - 3 * mm
        deck.setFillColor(color); deck.roundRect(bar_x, center_y - 2.6 * mm, bar_width, 5.2 * mm, 1.3 * mm, stroke=0, fill=1)
    draw_footer(deck, page_width, page_number)


def draw_poc_support_page(deck: PptCanvas, page_width: float, page_height: float,
                          assessment: dict, page_number: int) -> None:
    """顧客サービス名を反映したPoC実施・支援のPPTXスライド。"""
    _draw_pptx_poc_support_page(deck, page_width, page_height, assessment, page_number)


def draw_oracle_closing_page(deck: PptCanvas, page_width: float, page_height: float) -> None:
    """白地にOracleロゴのみを配置するクロージングページ。"""
    deck.setFillColor(colors.white)
    deck.rect(0, 0, page_width, page_height, stroke=0, fill=1)

    logo_reader = ImageReader(str(ORACLE_LOGO_IMAGE))
    logo_width, logo_height = logo_reader.getSize()
    draw_width = 105 * mm
    draw_height = draw_width * logo_height / logo_width
    deck.drawImage(
        logo_reader,
        (page_width - draw_width) / 2,
        (page_height - draw_height) / 2,
        width=draw_width,
        height=draw_height,
        preserveAspectRatio=True,
        mask="auto",
    )


def _technical_focus_text(value: object, limit: int) -> str:
    """技術設計スライド用に、文を途中省略せず表示コピーを短く整える。"""
    text = _clean_pptx_display_text(value)
    if len(text) <= limit:
        return text
    for marker in ("。", "；", ";"):
        position = text.rfind(marker, 0, limit + 1)
        if position >= int(limit * 0.55):
            return text[: position + 1]
    # 一文が長い場合は、読点や区切りの直後で完結させる。省略記号は使わず、
    # スライドだけを読んでも未完の文に見えないよう句点を補う。
    for marker in ("、", "，", ",", "／", "・"):
        position = text.rfind(marker, 0, limit + 1)
        if position >= int(limit * 0.45):
            return text[:position].rstrip("、，,。／・") + "。"
    return text[:limit].rstrip("、，,。／・") + "。"


def _draw_pptx_poc_logic_detail_page(deck: PptCanvas, page_width: float, page_height: float,
                                      detail: dict, poc: dict, priority: int, page_number: int) -> None:
    """合意済みの優先PoCレイアウトを保ち、内容だけを技術実装へ更新する。"""
    dark = colors.HexColor("#1D252C")
    muted = colors.HexColor("#4C5961")
    accents = [colors.HexColor(value) for value in POC_THEME_ACCENT_HEX]
    accent = accents[max(0, min(priority - 1, len(accents) - 1))]
    flow_fills = (
        colors.HexColor("#D3E3D8"),
        colors.HexColor("#F3DDD8"),
        colors.HexColor("#DCEAF1"),
    )
    flow_fill = flow_fills[max(0, min(priority - 1, len(flow_fills) - 1))]
    neutral_fill = colors.HexColor("#FAF9F7")
    content_x, content_w = 18 * mm, page_width - 36 * mm
    theme = _pptx_summary_text(display_theme_name(detail, poc), 48)
    title = f"AI技術テーマ {priority}／3｜{theme}"
    technical = poc_technical_design_for(detail, poc)

    # この3ページは、ユーザー承認済みの旧版と同じ白地・24pt見出しを使う。
    # 共通ヘッダー帯へ置き換えると本文構成が同じでも別レイアウトに見えるため、
    # 技術内容の更新時もこのページ固有の外観を固定する。
    deck.setFillColor(colors.white)
    deck.rect(0, 0, page_width, page_height, stroke=0, fill=1)
    draw_paragraph(
        deck, title,
        ParagraphStyle(
            "poc_approved_title", fontName="AssessmentJapaneseBold", fontSize=24.0,
            leading=28.8, textColor=dark, wordWrap="CJK",
        ),
        content_x, page_height - 17.5 * mm, content_w,
    )
    lead_style = ParagraphStyle(
        "poc_technical_lead", fontName="AssessmentJapanese", fontSize=15.0,
        leading=18.0, textColor=dark, wordWrap="CJK",
    )
    draw_paragraph(
        deck, "実装概要：" + _technical_focus_text(technical["implementation_summary"], 102),
        lead_style, content_x, WIDE_EXPLANATORY_LEAD_TOP + 3 * mm, content_w,
    )
    draw_paragraph(
        deck, "対象データ：" + _technical_focus_text(technical["input_data"], 108),
        ParagraphStyle(
            "poc_technical_scope", fontName="AssessmentJapanese", fontSize=14.0,
            leading=16.8, textColor=dark, wordWrap="CJK",
        ),
        content_x, 143 * mm, content_w,
    )

    # 前版と同じ5箱の横方向フローを保持する。案件ごとに処理名だけを差し替える。
    deck.setFillColor(neutral_fill); deck.rect(content_x + 3 * mm, 126 * mm, 127 * mm, 6.5 * mm, stroke=0, fill=1)
    deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 14.0)
    deck.drawString(content_x + 4 * mm, 128 * mm, "実装対象：データ準備から既存業務への返却まで")
    modality = technical_modality_for(poc, detail)["kind"]
    flow_by_modality = {
        "anomaly_ml": ("入力・特徴量を整備", "OMLで予測・判定", "結果・根拠を保存"),
        "optimization": ("入力・制約を整備", "OML＋制約で\n順位算出", "順位・根拠を保存"),
        "rag": ("文書・権限を整備", "Vector検索＋\n回答案生成", "回答・参照を保存"),
        "document_understanding": ("文書・画像を受付", "項目抽出・照合", "抽出・修正を保存"),
        "generative_assist": ("入力・参照を整備", "生成AIで処理", "結果・根拠を保存"),
    }
    flow_labels = list(flow_by_modality.get(modality, flow_by_modality["generative_assist"]))
    flow_labels += ["結果・根拠を\n既存画面へ返却", "担当者が確認し\n業務へ反映"]
    box_y, box_h = 105.4 * mm, 14.7 * mm
    box_offsets = (5.6, 64.6, 123.5, 182.4, 248.2)
    for index, label in enumerate(flow_labels):
        x = content_x + box_offsets[index] * mm
        box_w = (53.3 if index == 4 else 51.6) * mm
        deck.setFillColor(flow_fill if index < 4 else neutral_fill)
        deck.setStrokeColor(accent); deck.setLineWidth(0.8)
        deck.roundRect(x, box_y, box_w, box_h, 1.5 * mm, stroke=1, fill=1)
        draw_paragraph(
            deck, label,
            ParagraphStyle(
                f"poc_technical_flow_{index}", fontName="AssessmentJapaneseBold",
                fontSize=14.0, leading=15.5, alignment=1, textColor=dark, wordWrap="CJK",
            ),
            x + 1.5 * mm, box_y + box_h - 1.2 * mm, box_w - 3 * mm,
        )
        if index < 4:
            arrow_x = x + box_w + 1.5 * mm
            deck.setFillColor(colors.HexColor("#24576C"))
            arrow_w = 11.2 * mm if index == 3 else 4.3 * mm
            deck.rect(arrow_x, box_y + 5.8 * mm, arrow_w, 3.0 * mm, stroke=0, fill=1)

    # 前版の3列×3行テーブルを維持し、評価論ではなく実装設計を記載する。
    table_rows = [
        ["技術観点", "Oracle Database／OCIでの実装", "既存システムとの接続・運用"],
        [
            "データ・AI処理",
            _technical_focus_text(technical["oracle_technologies"], 72).replace("／", "／\n"),
            _technical_focus_text(technical["output_interface"], 60),
        ],
        [
            "DBオブジェクト",
            _technical_focus_text(technical["database_objects"], 72).replace("／", "／\n"),
            "入力・結果・根拠・担当者の修正履歴を、既存画面またはAPIから追跡可能にします。",
        ],
        [
            "配置・統制・評価",
            _technical_focus_text(technical["implementation_boundary"], 54)
            + "\n接続候補：検証データ複製・Database Link参照・API連携",
            _technical_focus_text(technical["control_design"], 40)
            + "\n評価：" + _technical_focus_text(
                ("検索再現率・引用正確性・回答不能率・人手確認率"
                 if modality == "rag" else technical["validation_plan"]), 34,
            ),
        ],
    ]
    deck.add_table(
        table_rows, content_x, 26.2 * mm,
        [42 * mm, 128.3 * mm, 128.3 * mm],
        [11.4 * mm, 19.0 * mm, 19.0 * mm, 19.4 * mm],
        accent, [neutral_fill, neutral_fill, neutral_fill], 14.0,
        cell_margin_mm=2.5, cell_vertical_margin_mm=0.8, line_spacing_pt=14.5,
    )
    bottom_note = (
        "補足：実装位置は既存DB内／ADBへのオフロードから、"
        "バージョン・データ量・権限・遅延要件で選択します。"
    )
    draw_paragraph(
        deck, bottom_note,
        ParagraphStyle(
            "poc_technical_note", fontName="AssessmentJapaneseBold", fontSize=14.0,
            leading=16.8, textColor=muted, wordWrap="CJK",
        ),
        content_x, 22.3 * mm, content_w,
    )
    draw_footer(deck, page_width, page_number)

def draw_poc_logic_detail_page(deck: PptCanvas, page_width: float, page_height: float,
                               detail: dict, poc: dict, priority: int, page_number: int) -> None:
    """1つの優先PoCを、業務・技術・検証の観点でPPTX一ページに具体化する。"""
    _draw_pptx_poc_logic_detail_page(deck, page_width, page_height, detail, poc, priority, page_number)


def service_use_case_groups_for(assessment: dict) -> list[dict]:
    """サービス別の一覧を返す。旧形式の応答は代表サービス1件として扱う。"""
    groups = assessment.get("service_use_case_groups")
    if isinstance(groups, list) and groups:
        return groups[:3]
    return [{"service_name": str(assessment.get("service_name")), "service_type": str(assessment.get("service_genre") or "業務サービス"), "use_cases": assessment["use_cases"]}]


def _industry_reference_urls(research: dict | None, assessment: dict) -> list[str]:
    """リサーチ監査で利用可とした業界根拠と、対象製品の公式情報を選ぶ。"""
    if not isinstance(research, dict):
        return []
    sources = [
        item for item in research.get("industry_sources", [])
        if isinstance(item, dict) and item.get("fetch_status") == "fetched"
        and str(item.get("url") or "").startswith(("https://", "http://"))
        and "google." not in str(item.get("url") or "").lower()
    ]
    if not sources:
        return []
    company = re.sub(
        r"^(?:株式会社|有限会社|合同会社)", "",
        _clean_pptx_display_text(assessment.get("company_name")),
    ).strip()
    service = _clean_pptx_display_text(assessment.get("service_name"))

    def relevance(item: dict) -> tuple[int, int]:
        title = _clean_pptx_display_text(item.get("title"))
        score = (4 if service and service in title else 0) + (2 if company and company in title else 0)
        return score, -len(str(item.get("url") or ""))

    selected: list[dict] = []
    official = max(sources, key=relevance)
    if relevance(official)[0] > 0:
        selected.append(official)
    audits = [item for item in research.get("research_audit", []) if isinstance(item, dict)]
    usable_ids: list[str] = []
    for audit in reversed(audits):
        details = audit.get("audit") if isinstance(audit.get("audit"), dict) else {}
        usable_ids = [str(value) for value in details.get("usable_source_ids", []) if str(value)]
        if usable_ids:
            break
    by_id = {str(item.get("id") or ""): item for item in sources}
    for source_id in usable_ids:
        item = by_id.get(source_id)
        if item is not None and item not in selected:
            selected.append(item)
    if len(selected) < 3:
        for item in sorted(sources, key=relevance, reverse=True):
            if relevance(item)[0] <= 0 or item in selected:
                continue
            selected.append(item)
            if len(selected) == 3:
                break
    return [str(item["url"]) for item in selected[:2]]


def draw_story_header(deck: PptCanvas, page_width: float, page_height: float,
                      title: str, section: str = "ASSESSMENT STORY") -> None:
    _executive_evidence_header(deck, page_width, page_height, title, section)


def industry_trend_digest_for(research: dict | None, assessment: dict) -> dict:
    """Return a verified three-topic digest or explicit research gaps.

    Older JSON remains renderable without promoting vendor or IR snippets
    into unsupported industry AI developments.
    """
    research = research if isinstance(research, dict) else {}
    explicit = research.get("industry_trend_digest")
    if isinstance(explicit, dict):
        rows = explicit.get("rows")
        if isinstance(rows, list) and len(rows) == 3 and all(
            isinstance(row, dict)
            and all(str(row.get(key) or "").strip() for key in ("topic", "development", "implication"))
            for row in rows
        ):
            return copy.deepcopy(explicit)

    # Legacy JSON remains renderable, but vendor/IR snippets are never promoted
    # to industry AI developments. Unverified content is shown as a research gap.
    scope = research.get("industry_scope") or {}
    domain = _clean_pptx_display_text(scope.get("domain_label")) or "対象業界"
    return {
        "domain_label": domain,
        "title": f"{domain}の最新動向とAI活用",
        "lead": "入力の顧客像・対象業務を基に、利用者側の業界でのAI活用を確認します。",
        "rows": [{"topic": f"業界AI活用の確認 {i}",
                  "development": "対象業界のAI活用を裏付ける公開原典は追加調査の対象です。",
                  "implication": "利用者の業務とAI処理の対応を確認します。", "url": ""}
                 for i in range(1, 4)],
        "proposal": "業界事例の業務・データ・実装条件を照合し、対象サービスへの適用を検討します。",
        "verified_at": "",
    }


def draw_source_urls(deck: PptCanvas, page_width: float, urls: list[str], *, top: float = 25 * mm) -> None:
    """Shared 9pt source line above the inherited footer for research slides."""
    style = ParagraphStyle("assessment_source_urls", fontName="AssessmentJapanese",
                           fontSize=9.0, leading=10.2, textColor=colors.HexColor("#4C5961"),
                           wordWrap="CJK")
    text = "出典 " + "　".join(f"※{i} {url}" for i, url in enumerate(urls, 1))
    draw_paragraph(deck, text, style, WIDE_CONTENT_LEFT, top,
                   wide_content_width(page_width), minimum_font_size=9.0,
                   shape_name=PPTX_SOURCE_NOTE_SHAPE_NAME)


def draw_industry_trends_page(deck: PptCanvas, page_width: float, page_height: float,
                              assessment: dict, research: dict | None,
                              page_number: int) -> None:
    """Render three current industry developments before the use-case catalog."""
    digest = industry_trend_digest_for(research, assessment)
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    draw_pptx_base_header(
        deck, page_width, page_height,
        str(digest.get("title") or "業界・サービスの最新動向とAI活用"),
        assessment_subsection_label("industry"),
    )
    lead_style = ParagraphStyle(
        "industry_trends_lead", fontName="AssessmentJapanese", fontSize=14.0,
        leading=16.4, textColor=dark, wordWrap="CJK",
    )
    draw_paragraph(deck, str(digest.get("lead") or ""), lead_style,
                   WIDE_CONTENT_LEFT, 154 * mm, wide_content_width(page_width))
    rows = [["注目トピック・発表時期", "新たにできること・具体的な使われ方", "対象サービスへの示唆"]]
    rows.extend([
        [str(item.get("topic") or ""), str(item.get("development") or ""), str(item.get("implication") or "")]
        for item in digest["rows"][:3]
    ])
    deck.add_table(
        rows, WIDE_CONTENT_LEFT, 47 * mm,
        [71 * mm, 143 * mm, wide_content_width(page_width) - 214 * mm],
        # リード文を2行まで安全に置けるよう、表の上端を7mm下げる。
        # 行内の14pt本文は維持し、元スライドと同じ3トピック構成にする。
        [12 * mm, 27 * mm, 27 * mm, 27 * mm],
        colors.HexColor("#365F47"), [colors.HexColor("#F8F8F7")] * 3,
        14.0, cell_margin_mm=2.6, cell_vertical_margin_mm=1.6,
        line_spacing_pt=15.5,
    )
    proposal_style = ParagraphStyle(
        "industry_trends_proposal", fontName="AssessmentJapaneseBold", fontSize=14.0,
        leading=16.2, textColor=dark, wordWrap="CJK",
    )
    draw_paragraph(deck, str(digest.get("proposal") or ""), proposal_style,
                   WIDE_CONTENT_LEFT, 44 * mm, wide_content_width(page_width))
    source_urls = [str(item.get("url") or "") for item in digest["rows"] if str(item.get("url") or "")]
    policy_source_url = str(digest.get("policy_source_url") or "")
    if policy_source_url:
        source_urls.append(policy_source_url)
    draw_source_urls(deck, page_width, source_urls)
    draw_footer(deck, page_width, page_number)


def draw_industry_value_story_page(deck: PptCanvas, page_width: float, page_height: float,
                                   story: dict, front_matter: dict, page_number: int,
                                   measurement_design: dict | None = None,
                                   business_impact: dict | None = None) -> None:
    """業界・サービスの論点からAI実装の価値までを、編集可能なフローで示す。"""
    dark = colors.HexColor("#1D252C")
    muted = colors.HexColor("#4C5961")
    accents = [colors.HexColor("#C74634"), colors.HexColor("#367A9B"), colors.HexColor("#467653")]
    fills = [colors.HexColor("#FFF7F1"), colors.HexColor("#F2F7FA"), colors.HexColor("#F1F7F2")]
    title = "業界・サービスの変化を、AIで事業価値へつなぐ"
    draw_story_header(deck, page_width, page_height, title, section=assessment_subsection_label("industry"))
    lead = ParagraphStyle("industry_value_wide_lead", fontName="AssessmentJapanese", fontSize=10.0,
                          leading=12.4, textColor=muted, wordWrap="CJK")
    detail = ParagraphStyle("industry_value_wide_detail", fontName="AssessmentJapanese", fontSize=10.0,
                            leading=12.0, textColor=muted, wordWrap="CJK")
    draw_paragraph(deck, consultative_page_lead(
        story["headline"], title=title,
    ), lead, 18 * mm, 156 * mm, page_width - 36 * mm)
    deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 10.5)
    deck.drawString(18 * mm, 136 * mm, "AI実装へ接続する、3つの事業・業務論点")
    lane_x, lane_w = 18 * mm, page_width - 36 * mm
    number_w, change_w, judgment_w = 14 * mm, 75 * mm, 76 * mm
    value_x = lane_x + number_w + change_w + judgment_w
    normalized_design = normalize_poc_measurement_design(measurement_design)
    impact_items = (
        business_impact.get("items", [])
        if isinstance(business_impact, dict)
        and business_impact.get("status") == "decision_thresholds"
        and isinstance(business_impact.get("items"), list)
        and len(business_impact.get("items")) == 3
        else []
    )
    deck.setFillColor(colors.HexColor("#68747B")); deck.setFont("AssessmentJapaneseBold", 9.2)
    deck.drawString(lane_x + number_w + 4 * mm, 130 * mm, "観察した変化")
    deck.drawString(lane_x + number_w + change_w + 4 * mm, 130 * mm, "AIで支援する判断")
    deck.drawString(
        value_x + 4 * mm, 130 * mm,
        "AI導入効果目標" if impact_items else "PoCで確認する価値",
    )
    quantitative_items = (
        normalized_design.get("items", [])
        if normalized_design.get("status") in {"decision_thresholds", "external_verified"}
        else []
    )
    outcomes = story.get("outcomes", [])
    for index, driver in enumerate(story["drivers"][:3]):
        y, lane_h = 102 * mm - index * 28 * mm, 24 * mm
        accent = accents[index]
        deck.setFillColor(fills[index]); deck.rect(lane_x, y, lane_w, lane_h, stroke=0, fill=1)
        deck.setFillColor(accent); deck.rect(lane_x, y, 4 * mm, lane_h, stroke=0, fill=1)
        # 番号は小さな識別子として扱い、内容より強く見せない。
        deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 12.5)
        deck.drawCentredString(lane_x + 9 * mm, y + 14.5 * mm, str(index + 1))
        deck.setStrokeColor(colors.white); deck.setLineWidth(0.45)
        for separator_x in (lane_x + number_w + change_w, lane_x + number_w + change_w + judgment_w):
            deck.line(separator_x, y + 3 * mm, separator_x, y + lane_h - 3 * mm)
        deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 10.0)
        draw_paragraph(deck, str(driver["label"]), ParagraphStyle(
            "industry_lane_title", fontName="AssessmentJapaneseBold", fontSize=10.0,
            leading=12.0, textColor=dark, wordWrap="CJK"),
            lane_x + number_w + 4 * mm, y + lane_h - 7 * mm, change_w - 8 * mm)
        # 会社・製品全体の計画試算を示す経路では「仮説」を
        # 強調せず、試算の前提となる事業上の着眼点として表示する。
        judgment_label = "着眼点" if impact_items else basis_label(driver["basis"])
        draw_paragraph(deck, f"{judgment_label}｜{driver['detail']}", detail,
                       lane_x + number_w + change_w + 4 * mm, y + lane_h - 6 * mm, judgment_w - 8 * mm)
        outcome = outcomes[index] if index < len(outcomes) and isinstance(outcomes[index], dict) else {}
        impact_item = impact_items[index] if index < len(impact_items) and isinstance(impact_items[index], dict) else None
        quantitative_item = quantitative_items[index] if index < len(quantitative_items) and isinstance(quantitative_items[index], dict) else None
        deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 10.0)
        if impact_item:
            impact_label = (
                "経営計画接続｜公開目標"
                if impact_item.get("basis_type") == "verified_management_target"
                else "AI導入効果目標｜計画試算"
            )
            deck.drawString(value_x + 4 * mm, y + lane_h - 7 * mm, impact_label)
            deck.setFont("AssessmentJapaneseBold", 16.0)
            deck.drawString(value_x + 4 * mm, y + 7.6 * mm, str(impact_item["headline_metric"]))
            deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 10.0)
            deck.drawString(
                value_x + 36 * mm, y + 7.1 * mm,
                _pptx_summary_text(impact_item.get("kpi"), 38),
            )
        elif quantitative_item:
            deck.drawString(value_x + 4 * mm, y + lane_h - 7 * mm, "事業効果目標")
            deck.setFont("AssessmentJapaneseBold", 16.0)
            deck.drawString(value_x + 4 * mm, y + 7.6 * mm, str(quantitative_item["headline_metric"]))
            deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 10.0)
            deck.drawString(
                value_x + 4 * mm, y + 3.2 * mm,
                _pptx_summary_text(quantitative_item.get("kpi"), 28),
            )
        else:
            deck.drawString(value_x + 4 * mm, y + lane_h - 7 * mm, "期待する変化")
            draw_paragraph(
                deck,
                str(outcome.get("detail") or "対象KPI・比較条件・責任者を合意し、実測で判断します。"),
                detail, value_x + 4 * mm, y + lane_h - 12 * mm,
                lane_x + lane_w - value_x - 8 * mm,
            )
    decision_y, decision_h = 23 * mm, 17 * mm
    deck.setFillColor(colors.HexColor("#E7F0EA")); deck.rect(lane_x, decision_y, lane_w, decision_h, stroke=0, fill=1)
    deck.setFillColor(colors.HexColor("#467653")); deck.rect(lane_x, decision_y, 5 * mm, decision_h, stroke=0, fill=1)
    if impact_items:
        deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 10.0)
        deck.drawString(lane_x + 10 * mm, decision_y + 10.4 * mm, "次頁：会社・製品全体のAI導入効果")
        draw_paragraph(
            deck,
            "3つの計画試算を会社側KPIとして算定し、効果の因果・算定式・PoC合格ラインを確認します。",
            detail, lane_x + 78 * mm, decision_y + 11.1 * mm, lane_w - 84 * mm,
        )
    elif normalized_design.get("status") in {"decision_thresholds", "external_verified"}:
        deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 10.0)
        deck.drawString(lane_x + 10 * mm, decision_y + 10.4 * mm, "次頁：製品先行KPI")
        summary_x = lane_x + 47 * mm
        summary_w = (lane_w - 53 * mm) / 3
        for summary_index, item in enumerate(normalized_design["items"]):
            x = summary_x + summary_index * summary_w
            accent = accents[summary_index]
            product = item.get("product_kpi") if isinstance(item.get("product_kpi"), dict) else {}
            deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 13.0)
            deck.drawString(x, decision_y + 10.0 * mm, str(product.get("target") or ""))
            deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 10.0)
            deck.drawString(
                x, decision_y + 4.4 * mm,
                _pptx_summary_text(product.get("kpi"), 22),
            )
    else:
        deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 10.0)
        deck.drawString(lane_x + 10 * mm, decision_y + 10.4 * mm, "今回の意思決定")
        draw_paragraph(
            deck,
            str(story.get("decision_question") or "対象業務・データ・成功指標を合意し、優先テーマからPoCを開始します。"),
            detail, lane_x + 47 * mm, decision_y + 11.1 * mm, lane_w - 53 * mm,
        )


def draw_poc_measurement_design_page(deck: PptCanvas, page_width: float, page_height: float,
                                     page_number: int) -> None:
    """未検証の改善率を置かず、PoCで数値を確定する方法を一枚で示す。"""
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    accents = [colors.HexColor("#467653"), colors.HexColor("#367A9B"),
               colors.HexColor("#C74634"), colors.HexColor("#6A5B8C")]
    fills = [colors.HexColor("#EEF4F0"), colors.HexColor("#EDF4F8"),
             colors.HexColor("#FFF4E9"), colors.HexColor("#F2EFF8")]
    body = ParagraphStyle("measurement_design_body", fontName="AssessmentJapanese", fontSize=12.0,
                          leading=14.4, textColor=muted, wordWrap="CJK")
    draw_pptx_base_header(deck, page_width, page_height, "PoCで確定する定量KPIと算定方法",
                          assessment_subsection_label("value_model"))
    draw_paragraph(
        deck,
        "公開根拠または顧客承認済みの基準値がない段階では、現状値・比較条件・算定式・必要データから整理する進め方が適しています。"
        "同じ母集団と判定基準で実測し、確認できた差分を本番投資と横展開の判断材料として扱います。",
        body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP, wide_content_width(page_width),
    )
    gap_x, gap_y = 6 * mm, 6 * mm
    card_w = (wide_content_width(page_width) - gap_x) / 2
    card_h = 49 * mm
    top_y = 83 * mm
    for index, item in enumerate(poc_measurement_design_items()):
        row, column = divmod(index, 2)
        x = WIDE_CONTENT_LEFT + column * (card_w + gap_x)
        y = top_y - row * (card_h + gap_y)
        deck.setFillColor(fills[index]); deck.roundRect(x, y, card_w, card_h, 2.5 * mm, stroke=0, fill=1)
        deck.setFillColor(accents[index]); deck.rect(x, y + card_h - 5 * mm, card_w, 5 * mm, stroke=0, fill=1)
        deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 12.0)
        deck.drawString(x + 5 * mm, y + card_h - 12 * mm, f"{item['dimension']}｜{item['kpi']}")
        labels = (
            ("現状値", item["baseline"]),
            ("比較条件", item["comparison"]),
            ("算定式", item["formula"]),
            ("必要データ", item["data"]),
        )
        line_y = y + card_h - 19 * mm
        for label, value in labels:
            deck.setFillColor(accents[index]); deck.setFont("AssessmentJapaneseBold", 12.0)
            deck.drawString(x + 5 * mm, line_y, label)
            draw_paragraph(
                deck, _pptx_summary_text(value, 36), body,
                x + 29 * mm, line_y + 1 * mm, card_w - 34 * mm,
            )
            line_y -= 8.0 * mm
    draw_footer(deck, page_width, page_number)


def _business_value_measurement(area: str, index: int) -> tuple[str, str]:
    """価値領域名から、断定値を置かないKPIと算定式を割り当てる。"""
    if re.search(r"納期|欠品|売上|機会", area):
        return "納期遵守率・遅延件数・緊急対応件数", "損失抑制効果＝遅延・欠品対象の粗利額×改善率"
    if re.search(r"在庫|資金", area):
        return "在庫金額・回転日数・長期滞留在庫", "運転資金改善額＝対象在庫金額×在庫削減率"
    if re.search(r"生産|効率|工数|時間|標準", area):
        return "計画・確認・検索時間、手動修正率", "工数削減効果＝対象件数×現行平均時間×時間単価×削減率"
    if re.search(r"競争|製品|契約|収益|継続", area):
        return "AI提案率・有償付帯率・継続利用率", "付帯収益＝AI提案対象契約数×有償付帯率×平均月額"
    defaults = (
        ("品質・正答率・手戻り率", "品質改善率＝（AI利用時－現行）÷現行"),
        ("処理時間・判断時間・待ち時間", "時間削減率＝（現行時間－AI利用時）÷現行時間"),
        ("利用率・採用率・継続率", "継続利用率＝継続利用者数÷対象利用者数"),
    )
    return defaults[min(index, len(defaults) - 1)]


def draw_assessment_business_impact_page(deck: PptCanvas, page_width: float, page_height: float,
                                         assessment: dict, page_number: int,
                                         research: dict | None = None) -> None:
    """顧客業務価値と提供企業の事業価値を、未承認の効果値なしで接続する。"""
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    accents = [colors.HexColor("#C74634"), colors.HexColor("#367A9B"), colors.HexColor("#B86A00")]
    body = ParagraphStyle("impact_bridge_body", fontName="AssessmentJapanese", fontSize=14.0,
                          leading=16.1, textColor=muted, wordWrap="CJK")
    bold = ParagraphStyle("impact_bridge_bold", fontName="AssessmentJapaneseBold", fontSize=14.0,
                          leading=16.1, textColor=dark, wordWrap="CJK")
    domain_label = _clean_pptx_display_text(
        industry_trend_digest_for(research, assessment).get("domain_label")
    ) or "対象サービス"
    value = assessment.get("business_value") if isinstance(assessment.get("business_value"), dict) else {}
    areas = [x for x in value.get("impact_areas", []) if isinstance(x, dict)][:3]
    while len(areas) < 3:
        areas.append({"area": ("品質", "時間", "利用・定着")[len(areas)],
                      "ai_enabled": "AI支援による変化を現行業務と比較します。",
                      "expected_impact": "同じ条件でKPIを確認し、適用範囲を検討します。"})
    draw_pptx_base_header(deck, page_width, page_height,
                          f"{domain_label}へのAI活用で期待できるビジネスインパクト",
                          assessment_subsection_label("value_model"))
    draw_paragraph(
        deck,
        "導入企業の業務効果と、提供企業の製品・事業価値を分けて測定します。"
        "現時点では効果値を置かず、顧客の現状値と同じ条件で比較できる算定式を準備します。",
        body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP, wide_content_width(page_width),
    )
    gap = 6 * mm
    card_w = (wide_content_width(page_width) - gap * 2) / 3
    card_y, card_h = 42 * mm, 88 * mm
    for index, item in enumerate(areas):
        x = WIDE_CONTENT_LEFT + index * (card_w + gap)
        accent = accents[index]
        area = _clean_pptx_display_text(item.get("area")) or f"価値領域{index + 1}"
        kpi, formula = _business_value_measurement(area, index)
        deck.setFillColor(colors.white); deck.setStrokeColor(colors.HexColor("#D8DEE2")); deck.setLineWidth(0.8)
        deck.roundRect(x, card_y, card_w, card_h, 3 * mm, stroke=1, fill=1)
        deck.setFillColor(accent); deck.rect(x, card_y + card_h - 2 * mm, card_w, 2 * mm, stroke=0, fill=1)
        draw_number_badge(deck, x + 10 * mm, card_y + card_h - 11 * mm, 5 * mm, index + 1, accent, 14.0)
        draw_paragraph(deck, area, bold, x + 19 * mm, card_y + card_h - 7 * mm, card_w - 25 * mm)
        deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 14.0)
        deck.drawString(x + 6 * mm, card_y + card_h - 24 * mm, "AIで変わる判断")
        draw_paragraph(deck, _clean_pptx_display_text(item.get("ai_enabled")), body,
                       x + 6 * mm, card_y + card_h - 30 * mm, card_w - 12 * mm)
        deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 14.0)
        deck.drawString(x + 6 * mm, card_y + 35 * mm, "確認するKPI")
        draw_paragraph(deck, kpi, body, x + 6 * mm, card_y + 29 * mm, card_w - 12 * mm)
        deck.setFillColor(colors.HexColor("#F6F7F8")); deck.roundRect(
            x + 5 * mm, card_y + 5 * mm, card_w - 10 * mm, 16 * mm, 2 * mm, stroke=0, fill=1,
        )
        draw_paragraph(deck, formula, bold, x + 8 * mm, card_y + 17 * mm, card_w - 16 * mm)
    deck.setFillColor(colors.HexColor("#EEF4F0")); deck.roundRect(
        WIDE_CONTENT_LEFT, 16 * mm, wide_content_width(page_width), 19 * mm, 2.5 * mm, stroke=0, fill=1,
    )
    deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 14.0)
    deck.drawString(WIDE_CONTENT_LEFT + 6 * mm, 28 * mm, "提供企業側の確認")
    draw_paragraph(
        deck,
        "導入企業で確認できた利用・採用・業務効果を、標準機能化、有償付帯、提供原価、継続利用の判断材料へつなげます。"
        "目標値は契約・利用・支援実績を確認して設定します。",
        body, WIDE_CONTENT_LEFT + 50 * mm, 29 * mm, wide_content_width(page_width) - 56 * mm,
    )
    draw_footer(deck, page_width, page_number)


def draw_poc_quantitative_target_page(deck: PptCanvas, page_width: float, page_height: float,
                                      design: dict, service_name: str, page_number: int) -> None:
    """サービス固有の定量目標を、実績主張と分離したPoC判定基準として示す。"""
    normalized = normalize_poc_measurement_design(design)
    if normalized.get("status") not in {"decision_thresholds", "external_verified"}:
        raise ValueError("検証済みの定量効果契約を描画できません。JSON契約を確認してください。")
    verified_external = normalized.get("status") == "external_verified"
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    accents = [colors.HexColor(value) for value in BUSINESS_VALUE_CARD_ACCENT_HEX]
    fills = [colors.HexColor("#FFF3EE"), colors.HexColor("#EDF4F8"), colors.HexColor("#EEF4F0")]
    body = ParagraphStyle(
        "quantitative_target_body", fontName="AssessmentJapanese", fontSize=12.0,
        leading=14.4, textColor=muted, wordWrap="CJK",
    )
    body_bold = ParagraphStyle(
        "quantitative_target_body_bold", fontName="AssessmentJapaneseBold", fontSize=12.0,
        leading=14.4, textColor=dark, wordWrap="CJK",
    )
    draw_pptx_base_header(
        deck, page_width, page_height,
        "AI実装で期待できる定量効果とPoC判定基準" if verified_external else "AI実装で目指す定量効果とPoC判定基準",
        assessment_subsection_label("value_model"),
    )
    lead = normalized.get("lead") or (
        f"{service_name}の業務データと優先PoCを踏まえ、AI実装で目指す定量目標を具体化します。"
        "同じ母集団と比較期間でPoC実測し、達成度を本番化と横展開の判断に用います。"
    )
    draw_paragraph(deck, lead, body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP,
                   wide_content_width(page_width))

    targets = normalized.get("management_targets", [])
    if targets:
        target_summary = " ｜ ".join(
            " ".join(part for part in (
                _clean_pptx_display_text(item.get("label")),
                _clean_pptx_display_text(item.get("target")),
                f"({_clean_pptx_display_text(item.get('delta'))})" if item.get("delta") else "",
            ) if part)
            for item in targets[:3]
        )
        deck.setFillColor(colors.HexColor("#EEF4F0")); deck.roundRect(
            WIDE_CONTENT_LEFT, 132 * mm, wide_content_width(page_width), 11 * mm,
            2.2 * mm, stroke=0, fill=1,
        )
        deck.setFillColor(colors.HexColor("#467653")); deck.setFont("AssessmentJapaneseBold", 12.0)
        deck.drawString(WIDE_CONTENT_LEFT + 5 * mm, 139 * mm, "中期経営計画との接続")
        draw_paragraph(
            deck, target_summary,
            ParagraphStyle(
                "quantitative_target_summary", fontName="AssessmentJapanese", fontSize=12.0,
                leading=14.4, textColor=dark, wordWrap="CJK",
            ),
            WIDE_CONTENT_LEFT + 45 * mm, 139 * mm, wide_content_width(page_width) - 50 * mm,
        )

    gap = 6 * mm
    card_w = (wide_content_width(page_width) - gap * 2) / 3
    # P6は3列の中で「テーマ → 事業効果 → 因果 → 製品KPI → 判定」を
    # 必ずこの順に積む。各ブロックを他ブロックの背景上に重ねない。
    card_y = 14.5 * mm
    card_h = 113.5 * mm
    label_style = ParagraphStyle(
        "quantitative_target_label", fontName="AssessmentJapaneseBold", fontSize=12.0,
        leading=14.4, textColor=dark, wordWrap="CJK",
    )
    for index, item in enumerate(normalized["items"]):
        x = WIDE_CONTENT_LEFT + index * (card_w + gap)
        top = card_y + card_h
        accent, fill = accents[index], fills[index]
        deck.setFillColor(fill); deck.roundRect(x, card_y, card_w, card_h, 3 * mm, stroke=0, fill=1)
        deck.setFillColor(accent); deck.rect(x, top - 5 * mm, card_w, 5 * mm, stroke=0, fill=1)

        draw_paragraph(deck, f"{item['priority']}｜{str(item['theme'])}", body_bold,
                       x + 6 * mm, top - 10 * mm, card_w - 12 * mm)
        metric_style = ParagraphStyle(
            f"quantitative_target_metric_{index}", fontName="AssessmentJapaneseBold",
            fontSize=22.0, leading=24.0, textColor=accent, wordWrap="CJK",
        )
        accent_label_style = ParagraphStyle(
            f"quantitative_target_accent_label_{index}",
            fontName="AssessmentJapaneseBold", fontSize=12.0, leading=14.4,
            textColor=accent, wordWrap="CJK",
        )
        product = item["product_kpi"]
        customer = item["customer_outcome_kpi"]
        provider = item["provider_business_kpi"]
        display_layer_name = str(item.get("display_kpi_layer") or "product_kpi")
        display_layer = {
            "product_kpi": product,
            "customer_outcome_kpi": customer,
            "provider_business_kpi": provider,
        }[display_layer_name]
        display_labels = {
            "product_kpi": "公開効果｜製品KPI",
            "customer_outcome_kpi": "顧客業務成果（AI寄与）",
            "provider_business_kpi": "提供者事業成果（AI寄与）",
        }
        draw_paragraph(deck, display_labels[display_layer_name], accent_label_style,
                       x + 6 * mm, top - 21.5 * mm, card_w - 12 * mm)
        draw_paragraph(deck, item["headline_metric"], metric_style,
                       x + 6 * mm, top - 28 * mm, card_w - 12 * mm)
        # KPI名は省略記号で切らず、最大2行のエリアで折り返す。
        draw_paragraph(deck, str(display_layer["kpi"]), accent_label_style,
                       x + 6 * mm, top - 40 * mm, card_w - 12 * mm)

        # 事業成果と製品KPIの因果は、長い原文を強制分断せず、
        # 次の製品KPI帯へつなぐ役割だけを一行で示す。原典照合済みの
        # 公開効果を表示する場合は、因果文の代わりにカードごとの出典を残す。
        bridge_text = (
            f"出典: {_pptx_summary_text(normalize_public_source_title(item.get('source_title')), 28)} "
            f"[{item.get('source_id')}]"
            if verified_external else
            "因果｜製品KPI達成 → 事業効果を同条件で実測"
        )
        draw_paragraph(deck, _pptx_summary_text(bridge_text, 28), body,
                       x + 6 * mm, top - 47.5 * mm, card_w - 12 * mm)

        # 製品KPIは大数値と重ならない通常フローの横帯とする。
        product_box_y, product_box_h = 47.5 * mm, 19 * mm
        deck.setFillColor(colors.white)
        deck.setStrokeColor(colors.HexColor("#D9E0E3")); deck.setLineWidth(0.45)
        deck.rect(x + 4 * mm, product_box_y, card_w - 8 * mm, product_box_h,
                 stroke=1, fill=1)
        draw_paragraph(deck, "製品先行KPI（提供者が直接管理）", label_style,
                       x + 7 * mm, product_box_y + 15 * mm, card_w - 14 * mm)
        product_metric_style = ParagraphStyle(
            f"quantitative_product_metric_{index}", fontName="AssessmentJapaneseBold",
            fontSize=17.0, leading=18.5, textColor=accent, wordWrap="CJK",
        )
        draw_paragraph(deck, str(product["target"]), product_metric_style,
                       x + 7 * mm, product_box_y + 8.7 * mm, 35 * mm)
        draw_paragraph(deck, str(product["kpi"]), body_bold,
                       x + 44 * mm, product_box_y + 8.7 * mm, card_w - 51 * mm)

        # 長文を2×2に押し込むのをやめ、経営判断に必要な数値・責任・
        # 停止原則だけを3行で示す。完全な原文はJSONと詳細PoCページに保持する。
        gate_h = 31 * mm
        deck.setFillColor(colors.HexColor("#F5F7F8")); deck.rect(
            x, card_y, card_w, gate_h, stroke=0, fill=1,
        )
        deck.setFillColor(accent); deck.rect(x, card_y, 2.5 * mm, gate_h, stroke=0, fill=1)
        gate_style = ParagraphStyle(
            f"quantitative_gate_{index}", fontName="AssessmentJapaneseBold", fontSize=12.0,
            leading=14.2, textColor=dark, wordWrap="CJK",
        )
        gate = item["poc_gate"]
        display_owner = display_layer["metric_owner"]
        confirm_needed = (
            display_owner["status"] == "confirm"
            or product["metric_owner"]["status"] == "confirm"
        )
        gate_lines = (
            f"事業効果Go｜{item['headline_metric']}＋{product['target']}",
            f"判定責任者｜{'提供者事業責任者' if display_owner['scope'] == 'provider' else '顧客・提供者共同'}",
            f"判定期間｜{_pptx_summary_text(gate['measurement_period'], 8)}{'・要確認' if confirm_needed else ''}",
            "停止条件｜品質・安全・運用逸脱",
        )
        for line_index, line_text in enumerate(gate_lines):
            draw_paragraph(
                deck, line_text, gate_style, x + 6 * mm,
                card_y + (gate_h - 3.0 * mm - line_index * 7.0 * mm), card_w - 12 * mm,
            )
    draw_footer(deck, page_width, page_number)


def draw_ai_product_business_impact_page(deck: PptCanvas, page_width: float, page_height: float,
                                         model: dict, assessment: dict,
                                         service_name: str, page_number: int) -> None:
    """Render the business-impact bridge as three independent value cards.

    The preceding evidence page and the following use-case catalog are both
    information-dense.  This page therefore avoids another grid or full-card
    color fill.  Each card is read vertically from value area to AI-enabled
    change, KPI/formula and the PoC decision gate.
    """
    normalized = normalize_ai_product_business_impact_candidate(model, assessment)
    if not normalized:
        raise ValueError("AI製品化の会社事業KPI契約を描画できません。元の入力からJSONを再生成してください。")
    role = normalized.get("organization_role")
    provider_case = role in {"provider", "mixed"}
    title = (
        "AI導入効果目標を、売上・利益・継続収益へ接続する"
        if provider_case else
        "AI導入効果目標を、売上・利益・事業成長へ接続する"
    )
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    # 色は面ではなく、価値領域を追うための細い識別子としてだけ使う。
    accents = [colors.HexColor("#C74634"), colors.HexColor("#367A9B"), colors.HexColor("#B86A00")]
    pale_fills = [colors.HexColor("#FFF5F1"), colors.HexColor("#F2F7FA"), colors.HexColor("#FFF8EC")]
    border = colors.HexColor("#D9E0E3")
    neutral = colors.HexColor("#F6F7F8")
    body = ParagraphStyle(
        "ai_product_business_body", fontName="AssessmentJapanese", fontSize=14.0,
        leading=16.2, textColor=muted, wordWrap="CJK",
    )
    body_bold = ParagraphStyle(
        "ai_product_business_bold", fontName="AssessmentJapaneseBold", fontSize=14.0,
        leading=16.2, textColor=dark, wordWrap="CJK",
    )
    draw_pptx_base_header(
        deck, page_width, page_height, title, assessment_subsection_label("value_model"),
    )
    generated_lead = normalized.get("lead") or (
        f"{service_name}へAIを組み込んで得られる顧客価値を、有償化・標準化・継続利用へ転換します。"
        "売上・利益・継続収益を評価対象企業のKPIとして測定し、本番化と横展開を判断します。"
    )
    lead_core = _pptx_summary_text(generated_lead, 82)
    lead = _pptx_summary_text(
        f"計画試算として、{lead_core} "
        "同じKPI定義と比較期間で検証し、事業化を判断します。",
        132,
    )
    draw_paragraph(deck, lead, body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP,
                   wide_content_width(page_width))

    targets = normalized.get("management_targets", [])
    if targets:
        target_summary = " ｜ ".join(
            " ".join(part for part in (
                _clean_pptx_display_text(item.get("label")),
                _clean_pptx_display_text(item.get("target")),
                f"({_clean_pptx_display_text(item.get('delta'))})" if item.get("delta") else "",
            ) if part)
            for item in targets[:3]
        )
        deck.setFillColor(neutral); deck.roundRect(
            WIDE_CONTENT_LEFT, 132 * mm, wide_content_width(page_width), 11 * mm,
            2.2 * mm, stroke=0, fill=1,
        )
        deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 14.0)
        deck.drawString(WIDE_CONTENT_LEFT + 5 * mm, 138 * mm, "中期経営計画との接続")
        draw_paragraph(
            deck, target_summary,
            ParagraphStyle(
                "ai_product_business_target_summary", fontName="AssessmentJapanese", fontSize=14.0,
                leading=16.2, textColor=dark, wordWrap="CJK",
            ),
            WIDE_CONTENT_LEFT + 53 * mm, 138.5 * mm,
            wide_content_width(page_width) - 58 * mm,
        )

    gap = 6 * mm
    card_w = (wide_content_width(page_width) - gap * 2) / 3
    card_y = 15.5 * mm
    card_h = (112.0 if targets else 120.0) * mm
    for index, item in enumerate(normalized["items"]):
        x = WIDE_CONTENT_LEFT + index * (card_w + gap)
        top = card_y + card_h
        accent, pale_fill = accents[index], pale_fills[index]
        deck.setFillColor(colors.white); deck.setStrokeColor(border); deck.setLineWidth(0.7)
        deck.roundRect(x, card_y, card_w, card_h, 3 * mm, stroke=1, fill=1)
        deck.setFillColor(accent); deck.rect(x, top - 2 * mm, card_w, 2 * mm, stroke=0, fill=1)
        draw_number_badge(deck, x + 10 * mm, top - 10 * mm, 5 * mm, index + 1, accent, 14.0)

        accent_label = ParagraphStyle(
            f"ai_product_business_label_{index}", fontName="AssessmentJapaneseBold",
            fontSize=14.0, leading=16.2, textColor=accent, wordWrap="CJK",
        )
        metric_style = ParagraphStyle(
            f"ai_product_business_metric_{index}", fontName="AssessmentJapaneseBold",
            fontSize=21.0, leading=22.0, textColor=accent, wordWrap="CJK",
        )
        impact_label = (
            "経営計画接続｜公開目標"
            if item.get("basis_type") == "verified_management_target"
            else "AI導入効果目標｜計画試算"
        )
        draw_paragraph(deck, impact_label, accent_label,
                       x + 19 * mm, top - 7 * mm, card_w - 25 * mm)
        draw_paragraph(deck, str(item["label"]), body_bold,
                       x + 19 * mm, top - 15 * mm, card_w - 25 * mm)

        # 大きな数値は淡い局所帯に置き、カード全体は白地のまま保つ。
        deck.setFillColor(pale_fill); deck.roundRect(
            x + 5 * mm, top - 43 * mm, card_w - 10 * mm, 21 * mm,
            2 * mm, stroke=0, fill=1,
        )
        draw_paragraph(deck, str(item["headline_metric"]), metric_style,
                       x + 9 * mm, top - 27 * mm, card_w - 18 * mm)
        draw_paragraph(deck, str(item["kpi"]), accent_label,
                       x + 9 * mm, top - 37 * mm, card_w - 18 * mm)

        gate = item["decision_gate"]
        leading = item["leading_indicator"]
        # 14ptで責任者まで省略せず表示すると4〜5行になるため、判定欄には
        # 十分な高さを確保する。カード全体の色は増やさず、下端の局所帯だけに置く。
        gate_h = 34 * mm
        causal_label_offset = 49
        causal_body_offset = 56
        formula_label_offset = 70
        formula_body_offset = 77
        causal_limit = 30 if targets else 38
        formula_limit = 26 if targets else 34
        deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 14.0)
        deck.drawString(x + 6 * mm, top - causal_label_offset * mm, "AIで変わる判断")
        draw_paragraph(deck, _pptx_summary_text(item["causal_chain"], causal_limit), body,
                       x + 6 * mm, top - causal_body_offset * mm, card_w - 12 * mm)
        deck.setFillColor(accent); deck.setFont("AssessmentJapaneseBold", 14.0)
        deck.drawString(x + 6 * mm, top - formula_label_offset * mm, "算定式")
        draw_paragraph(deck, _pptx_formula_text(
            item["formula"], item["headline_metric"], formula_limit,
        ), body,
                       x + 6 * mm, top - formula_body_offset * mm, card_w - 12 * mm)

        deck.setFillColor(neutral); deck.roundRect(
            x + 4 * mm, card_y + 4 * mm, card_w - 8 * mm, gate_h - 4 * mm,
            2 * mm, stroke=0, fill=1,
        )
        deck.setFillColor(accent); deck.rect(
            x + 4 * mm, card_y + 4 * mm, 2 * mm, gate_h - 4 * mm,
            stroke=0, fill=1,
        )
        gate_style = ParagraphStyle(
            f"ai_product_business_gate_{index}", fontName="AssessmentJapaneseBold",
            fontSize=14.0, leading=15.4, textColor=dark, wordWrap="CJK",
        )
        # 責任者・KPIの表示を版面都合で短文化しない。長い値は折返しと自動高さで
        # 保持し、必要であれば見た目の領域を越えても原文を失わないことを優先する。
        decision_owner = _clean_pptx_display_text(gate["decision_owner"])
        leading_kpi = _pptx_kpi_label(leading["kpi"], 12)
        gate_text_width = card_w - 8 * mm
        gate_headline = (
            f"経営目標｜{item['headline_metric']}へのAI寄与確認"
            if item.get("basis_type") == "verified_management_target"
            else f"PoC合格ライン｜{item['headline_metric']}達成"
        )
        gate_lines = (
            gate_headline,
            f"先行KPI｜{leading['target']} {leading_kpi}",
            f"判定｜{_pptx_period_label(gate['measurement_period'], 12)}・{decision_owner}",
        )
        gate_text = "\n".join(gate_lines)
        # 1つの段落として組み、実際の折返し行数に応じた高さを確保する。
        # 文字を省略せず、色帯の内側に自然に収める。
        draw_paragraph(
            deck, gate_text, gate_style,
            x + 9 * mm, card_y + gate_h - 1.2 * mm, gate_text_width - 5 * mm,
        )
    draw_footer(deck, page_width, page_number)


def compact_service_name_for_heading(service_name: str) -> str:
    """一覧見出しでも正式なサービス名を省略せずに返す。"""
    return _clean_pptx_display_text(service_name)


def ellipsize_for_width(deck: PptCanvas, text: str, font_name: str,
                        font_size: float, max_width: float) -> str:
    """幅不足でも文字を削らない、旧API互換の表示テキスト関数。"""
    _ = deck, font_name, font_size, max_width
    return _clean_pptx_display_text(text)


def _use_case_list_plain_text(value: object) -> str:
    """一覧セル用に改行・装飾を除いた表示文字列を返す。"""
    return re.sub(r"\s+", " ", _plain_text(str(value or ""))).strip()


def compact_use_case_description_for_pptx(description: object, *, max_chars: int = 42) -> str:
    """一覧説明を文字数上限で要約せず、原文のまま返す。"""
    _ = max_chars
    return _use_case_list_plain_text(description)


def fit_use_case_technology_text(deck: PptCanvas, technology: object,
                                 column_width: float) -> str:
    """AI技術名を、列幅にかかわらず全て保持する。"""
    _ = deck, column_width
    return _use_case_list_plain_text(technology)


def fit_complete_use_case_text(deck: PptCanvas, text: str, max_width: float) -> str:
    """一覧セルのテキストを削らず返し、折返し・行高へ委ねる。"""
    _ = deck, max_width
    return _use_case_list_plain_text(text)


def use_case_list_cell_text(deck: PptCanvas, value: object, column_width: float) -> str:
    """PPTX表セルのテキストを、幅判定で削らず返す。"""
    _ = deck, column_width
    return _use_case_list_plain_text(value)


def build_use_case_page_heading(deck: PptCanvas, page_width: float, service_name: str, service_type: str,
                                *, prefix: str = "", counter_label: str = "") -> str:
    """付録接頭辞・サービス番号を含め、見出しを削らず返す。"""
    _ = deck, page_width, counter_label
    suffix = "のAIユースケース一覧"
    prefix_text = f"{prefix}：" if prefix else ""
    compact_name = compact_service_name_for_heading(service_name)
    return f"{prefix_text}{compact_name}（{service_type}）{suffix}"


def _draw_midterm_management_target_card_text(
        deck: PptCanvas, x: float, y: float, width: float, height: float,
        target: dict, index: int, *, green: colors.Color, dark: colors.Color,
        muted: colors.Color) -> None:
    """中期経営目標の4行を、カード内で一体として縦中央に配置する。"""
    lines = [
        (_pptx_summary_text(target.get("label"), 14), 12.0, 14.2, green, True),
        (_pptx_summary_text(target.get("target"), 12), 13.5, 15.0, dark, True),
        (_pptx_summary_text(target.get("period"), 12), 12.0, 14.2, muted, False),
        (_pptx_summary_text(target.get("comparison"), 16), 12.0, 14.2, muted, False),
    ]
    lines = [line for line in lines if line[0]]
    if not lines:
        return

    inset_x = 4 * mm
    inset_y = 2 * mm
    box = deck.slide.shapes.add_textbox(
        deck._x(x + inset_x),
        deck._y(y + height - inset_y),
        deck._width(width - 2 * inset_x),
        deck._height(height - 2 * inset_y),
    )
    box.name = f"AI_ASSESS_MIDTERM_TARGET_TEXT_{index + 1}"
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    for line_index, (text, font_size, leading, color, bold) in enumerate(lines):
        paragraph = frame.paragraphs[0] if line_index == 0 else frame.add_paragraph()
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        paragraph.line_spacing = Pt(leading)
        run = paragraph.add_run()
        run.text = text
        run.font.name = PptCanvas.FONT_MAP["AssessmentJapaneseBold" if bold else "AssessmentJapanese"]
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.color.rgb = deck._rgb(color)


def midterm_ai_takeaway_for(analysis: dict) -> str:
    """Summarize the input-derived strategy/AI connections, without new claims."""
    explicit = _clean_pptx_display_text(analysis.get("ai_takeaway"))
    if explicit:
        return explicit
    themes = list(dict.fromkeys(
        _clean_pptx_display_text(item.get("related_use_case"))
        for item in analysis.get("ai_alignment", [])[:3] if isinstance(item, dict)
        and _clean_pptx_display_text(item.get("related_use_case"))
    ))
    return ("中期計画の重点方針に沿い、" + "、".join(themes) + "をAI活用の軸として具体化します。"
            if themes else "中期計画と対象業務の接点を基に、AIが支援する判断と実装範囲を具体化します。")


def draw_midterm_plan_alignment_page(deck: PptCanvas, page_width: float, page_height: float,
                                     analysis: dict, page_number: int,
                                     *, section: str | None = None) -> None:
    """中期経営計画の重点方針と、今回のAI提案の因果を経営層向けに示す。"""
    dark, green, orange = colors.HexColor("#1D252C"), colors.HexColor("#467653"), colors.HexColor("#C74634")
    muted = colors.HexColor("#4C5961")
    body = ParagraphStyle("midterm_body", fontName="AssessmentJapanese", fontSize=12.0,
                          leading=14.4, textColor=muted, wordWrap="CJK")
    small = ParagraphStyle("midterm_small", fontName="AssessmentJapanese", fontSize=12.0,
                           leading=14.2, textColor=muted, wordWrap="CJK")
    draw_pptx_base_header(
        deck, page_width, page_height, "中期経営計画とAI実装の整合性",
        section or assessment_subsection_label("delivery"),
    )
    draw_paragraph(deck, consultative_page_lead(
        str(analysis.get("plan_summary", "")), title="中期経営計画とAI実装の整合性", limit=150,
    ), body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP, wide_content_width(page_width))
    source = analysis.get("source", {}) if isinstance(analysis.get("source"), dict) else {}
    targets = [item for item in analysis.get("management_targets", []) if isinstance(item, dict)][:4]
    has_targets = bool(targets)
    if targets:
        target_gap = 4 * mm
        target_w = (wide_content_width(page_width) - target_gap * (len(targets) - 1)) / len(targets)
        for index, target in enumerate(targets):
            x = WIDE_CONTENT_LEFT + index * (target_w + target_gap)
            deck.setFillColor(colors.HexColor("#F2F6F3")); deck.roundRect(x, 116 * mm, target_w, 26 * mm, 2 * mm, stroke=0, fill=1)
            deck.slide.shapes[-1].name = f"AI_ASSESS_MIDTERM_TARGET_CARD_{index + 1}"
            _draw_midterm_management_target_card_text(
                deck, x, 116 * mm, target_w, 26 * mm, target, index,
                green=green, dark=dark, muted=muted,
            )
    # 数値目標が原典照合を通過した場合は従来座標を保つ。数値目標を表示しない
    # 場合は、空いた上段を出典帯・方針カードに再配分し、16:9の縦方向の密度を揃える。
    source_y = (101 if has_targets else 128) * mm
    source_label = "経営目標（AI効果とは別）" if has_targets else "参照した公開資料"
    source_text_x = WIDE_CONTENT_LEFT + (66 if has_targets else 48) * mm
    deck.setFillColor(colors.HexColor("#EEF4F0")); deck.roundRect(
        WIDE_CONTENT_LEFT, source_y, wide_content_width(page_width), 12 * mm, 2 * mm, stroke=0, fill=1,
    )
    source_label_width = (60 if has_targets else 42) * mm
    source_label_style = ParagraphStyle(
        "midterm_source_label", fontName="AssessmentJapaneseBold", fontSize=12.0,
        leading=14.2, textColor=green, wordWrap="CJK",
    )
    draw_paragraph(deck, source_label, source_label_style,
                   WIDE_CONTENT_LEFT + 4 * mm, source_y + 8.5 * mm, source_label_width - 6 * mm)
    # 資料名は上段に、完全なURLは下端の共通出典欄に表示する。
    source_text = f"出典：{normalize_public_source_title(source.get('title', ''))}"
    source_text_width = page_width - WIDE_CONTENT_RIGHT - source_text_x - 4 * mm
    draw_paragraph(
        deck,
        ellipsize_for_width(deck, source_text, "AssessmentJapanese", 12.0, source_text_width),
        small, source_text_x, source_y + 8.5 * mm, source_text_width,
    )
    alignments = analysis.get("ai_alignment", [])[:3]
    card_height = (23 if has_targets else 25) * mm
    y = (84 if has_targets else 101) * mm
    card_step = (23 if has_targets else 27) * mm
    heading_offset = card_height - 4 * mm
    detail_offset = card_height - 11.7 * mm
    for index, item in enumerate(alignments, 1):
        deck.setFillColor(colors.HexColor("#F5F7F8")); deck.roundRect(18 * mm, y, page_width - 36 * mm, card_height, 2 * mm, stroke=0, fill=1)
        draw_number_badge(deck, 27 * mm, y + card_height / 2, 4.2 * mm, index, orange)
        columns = (
            (35 * mm, 66 * mm, "中期計画の重点方針", _pptx_summary_text(item.get("plan_priority"), 24), dark),
            (105 * mm, 61 * mm, "AIで支援する役割", _pptx_summary_text(item.get("ai_role"), 26), green),
            (170 * mm, page_width - 188 * mm, "優先する理由・関連テーマ",
             f"{_pptx_reason_summary(item.get('why_now'), limit=28)} "
             f"関連：{_pptx_summary_text(item.get('related_use_case'), 22)}", dark),
        )
        for x, width, heading, detail, accent in columns:
            heading_style = ParagraphStyle(
                f"midterm_alignment_heading_{index}_{int(x)}", fontName="AssessmentJapaneseBold",
                fontSize=12.0, leading=14.2, textColor=accent, wordWrap="CJK",
            )
            draw_paragraph(deck, heading, heading_style, x, y + heading_offset, width)
            draw_paragraph(deck, detail, small, x, y + detail_offset, width)
        y -= card_step
    caveat_y = 25 * mm
    caveat_height = (12 if has_targets else 16) * mm
    deck.setFillColor(colors.HexColor("#FFF7EE")); deck.roundRect(
        WIDE_CONTENT_LEFT, caveat_y, wide_content_width(page_width), caveat_height, 2 * mm, stroke=0, fill=1,
    )
    caveat_text_y = caveat_y + caveat_height - 4 * mm
    draw_paragraph(
        deck, "AI活用方針",
        ParagraphStyle("midterm_takeaway_label", fontName="AssessmentJapaneseBold", fontSize=12.0,
                       leading=14.2, textColor=dark, wordWrap="CJK"),
        WIDE_CONTENT_LEFT + 4 * mm, caveat_text_y + 1.0 * mm, 34 * mm,
    )
    draw_paragraph(deck, midterm_ai_takeaway_for(analysis), small,
                   WIDE_CONTENT_LEFT + 41 * mm, caveat_text_y, wide_content_width(page_width) - 45 * mm)
    if source.get("url"):
        draw_source_urls(deck, page_width, [str(source["url"])], top=22 * mm)
    draw_footer(deck, page_width, page_number)


def use_case_list_column_widths(page_width: float) -> list[float]:
    """一覧表の描画・行高計算で共用する列幅を返す。"""
    left_margin, right_margin = 16 * mm, 16 * mm
    widths = list(USE_CASE_LIST_COLUMN_WIDTHS)
    widths.append(page_width - left_margin - right_margin - sum(widths))
    return widths


def use_case_list_native_row(case: dict) -> list[str]:
    """一覧の1行を、表示専用の短文化をせずに構成する。"""
    return [
        str(case.get("no", "")),
        _use_case_list_plain_text(case.get("use_case", "")),
        _use_case_list_plain_text(case.get("ai_technology", "")),
        _use_case_list_plain_text(case.get("description", "")),
    ]


def use_case_list_row_heights(cases: list[dict], widths: list[float]) -> list[float]:
    """Meiryo UIの実幅に近い見積りで、必要な行だけを拡張する。

    共通表の保守的な文字数推定は、境界付近のセルを実表示より多く折り返し、
    15件一覧を不要に複数ページへ分けていた。一覧表ではPptCanvasと同じ文字幅
    係数を使い、明示改行と実幅超過だけを追加行として数える。
    """

    def measured_lines(value: object, width: float) -> int:
        text = _use_case_list_plain_text(value)
        if not text:
            return 1
        usable_width = max(
            20.0,
            width - 2 * USE_CASE_LIST_CELL_MARGIN_MM * mm,
        ) * USE_CASE_LIST_WRAP_TOLERANCE
        measured_width = sum(
            MIN_PPTX_FONT_SIZE * (1 if ord(char) < 128 else 1.85) * 0.52
            for char in text
        )
        return max(1, int((measured_width + usable_width - 0.01) // usable_width))

    heights: list[float] = []
    for case in cases:
        lines = max(
            measured_lines(value, widths[index])
            for index, value in enumerate(use_case_list_native_row(case))
        )
        # PowerPoint側も12pt固定行間・上下0.4mmで描画する。単一行は共通の
        # 基準行高とし、真に折り返す行だけ必要量へ拡張する。
        required = lines * MIN_PPTX_FONT_SIZE + 0.8 * mm
        heights.append(max(USE_CASE_LIST_ROW_HEIGHT, required))
    return heights


def use_case_list_chunks(group: dict, *, cases_per_slide: int = USE_CASES_PER_LIST_SLIDE,
                         page_width: float = PPTX_WIDESCREEN_WIDTH) -> list[list[dict]]:
    """候補を最大15件かつ必要な表の高さで分割し、文字を削らない。"""
    cases = [case for case in group.get("use_cases", []) if isinstance(case, dict)]
    if not cases:
        return [[]]

    max_cases = max(1, int(cases_per_slide))
    widths = use_case_list_column_widths(page_width)
    row_heights = use_case_list_row_heights(cases, widths)
    # 見出し下からフッター上までの実領域を分割判定に使う。基準行高の
    # 15倍で判定すると、ごく少数の折返し行があるだけで14+1ページに
    # なり、表上の大きな空白を生むため、描画幾何と一致させる。
    available_body_height = (
        PPTX_WIDESCREEN_HEIGHT
        - USE_CASE_LIST_TABLE_TOP
        - USE_CASE_LIST_TABLE_BOTTOM
        - USE_CASE_LIST_HEADER_HEIGHT
    )
    chunks: list[list[dict]] = []
    current: list[dict] = []
    current_height = 0.0
    for case, row_height in zip(cases, row_heights):
        starts_new_page = current and (
            len(current) >= max_cases
            or current_height + row_height > available_body_height + 0.01
        )
        if starts_new_page:
            chunks.append(current)
            current, current_height = [], 0.0
        current.append(case)
        current_height += row_height
    if current:
        chunks.append(current)
    return chunks


def draw_use_case_list_page(deck: PptCanvas, page_width: float, page_height: float,
                            group: dict, cases: list[dict], page_number: int,
                            total_groups: int, service_index: int, list_page_index: int,
                            list_page_count: int, *, appendix: bool = False,
                            priority_pocs: list[dict] | None = None,
                            priority_icon_path: Path | None = None) -> None:
    """候補一覧を16:9・本文最小12ptで読めるページ単位に出力する。"""
    dark, green = colors.HexColor("#1D252C"), colors.HexColor("#467653")
    service_type = str(group.get("service_type") or "業務サービス").strip()
    # 1サービス15件は必ず1枚なので、単一サービスに冗長な「1/1」を出さない。
    # 複数サービスの場合も、対象サービス数と候補ページ数を混同しない表記にする。
    service_counter = f"サービス {service_index} / {total_groups}" if total_groups > 1 else ""
    first_no = str(cases[0].get("no", "")) if cases else ""
    last_no = str(cases[-1].get("no", "")) if cases else ""
    list_counter = f"候補 {first_no}–{last_no} / {len(group.get('use_cases', []))}" if list_page_count > 1 else ""
    counter_label = " | ".join(label for label in (service_counter, list_counter) if label)
    heading = build_use_case_page_heading(
        deck, page_width, str(group["service_name"]), service_type,
        prefix="参考" if appendix else "", counter_label=counter_label,
    )
    draw_pptx_base_header(deck, page_width, page_height, heading, assessment_subsection_label("use_cases"))
    if counter_label:
        deck.setFillColor(colors.HexColor("#4C5961")); deck.setFont("AssessmentJapanese", 12)
        deck.drawRightString(page_width - 18 * mm, page_height - 20 * mm, counter_label)
    # No.列を固定し、列幅は分割判定と同じ共通定義を使用する。収まりきらない
    # テキストは削らず、折返し・可変行高・追加ページで吸収する。
    left_margin, right_margin = 16 * mm, 16 * mm
    widths = use_case_list_column_widths(page_width)
    priority_ids = {
        str(item.get("use_case_id") or "").strip()
        for item in priority_pocs or []
        if str(item.get("use_case_id") or "").strip()
    }
    priority_numbers = {
        str(item.get("use_case_no") or "").strip()
        for item in priority_pocs or []
        if str(item.get("use_case_no") or "").strip()
    }

    def is_priority_case(case: dict) -> bool:
        """IDで候補を特定し、旧単一サービスJSONだけ番号で補完する。"""
        case_id = str(case.get("use_case_id") or "").strip()
        if case_id and priority_ids:
            return case_id in priority_ids
        return total_groups == 1 and str(case.get("no") or "").strip() in priority_numbers

    native_rows = [["No.", "ユースケース", "主なAI技術", "業務価値・処理"]]
    priority_row_indexes: list[int] = []
    for row_index, case in enumerate(cases):
        if is_priority_case(case):
            priority_row_indexes.append(row_index)
        native_rows.append(use_case_list_native_row(case))
    body_row_heights = use_case_list_row_heights(cases, widths)
    heights = [USE_CASE_LIST_HEADER_HEIGHT, *body_row_heights]
    row_colors = [USE_CASE_LIST_BODY_COLOR] * len(cases)
    # タイトル下の表を0.5mm下げ、下側の余白を活用する。フッターとの9mm
    # 安全域は維持し、優先候補アイコンも表基準で同じ量だけ追随する。
    table_bottom = USE_CASE_LIST_TABLE_BOTTOM
    deck.add_table(native_rows, left_margin, table_bottom, widths, heights,
                  green, row_colors or [USE_CASE_LIST_BODY_COLOR], 12.0,
                  cell_margin_mm=USE_CASE_LIST_CELL_MARGIN_MM,
                  cell_vertical_margin_mm=0.4,
                  line_spacing_pt=MIN_PPTX_FONT_SIZE)

    # 優先候補の行内色分けは行わず、左余白に同一のひらめきアイコンを置く。
    # アイコンが未配置の開発環境でも表の生成自体は維持し、asset配置後は
    # 同じ生成経路で自動的にマーカーが反映される。
    icon_path = Path(priority_icon_path) if priority_icon_path is not None else USE_CASE_LIST_PRIORITY_ICON_IMAGE
    if priority_row_indexes and icon_path.is_file():
        table_top = table_bottom + sum(heights)
        icon_x = left_margin - USE_CASE_LIST_PRIORITY_ICON_GAP - USE_CASE_LIST_PRIORITY_ICON_SIZE
        for row_index in priority_row_indexes:
            row_center_y = (
                table_top - USE_CASE_LIST_HEADER_HEIGHT
                - sum(body_row_heights[:row_index])
                - body_row_heights[row_index] / 2
            )
            deck.drawImage(
                str(icon_path), icon_x, row_center_y - USE_CASE_LIST_PRIORITY_ICON_SIZE / 2,
                USE_CASE_LIST_PRIORITY_ICON_SIZE, USE_CASE_LIST_PRIORITY_ICON_SIZE,
                preserveAspectRatio=True, mask="auto",
            )
    draw_footer(deck, page_width, page_number)


def default_document_slide_plan(assessment: dict, *,
                                include_cost_estimate: bool = True) -> tuple[SlideSpec, ...]:
    """標準PPTXのページ数と描画順が共有する論理スライド計画を返す。"""
    has_midterm = isinstance(assessment.get("midterm_plan_analysis"), dict) and bool(
        assessment["midterm_plan_analysis"].get("ai_alignment")
    )
    use_case_page_keys = [
        f"use_case_list:{service_index}:{list_page_index}"
        for service_index, group in enumerate(service_use_case_groups_for(assessment), 1)
        for list_page_index, _cases in enumerate(use_case_list_chunks(group), 1)
    ]
    return build_default_slide_plan(
        has_midterm=has_midterm,
        use_case_page_keys=use_case_page_keys,
        include_cost_estimate=include_cost_estimate,
    )


def draw_assessment_overview_page(deck: PptCanvas, page_width: float, page_height: float,
                                  assessment: dict, page_number: int) -> None:
    """公開情報と入力から得た暫定評価、理由、未確認事項を最初に示す。"""
    dark, muted = colors.HexColor("#1D252C"), colors.HexColor("#4C5961")
    accents = [colors.HexColor("#467653"), colors.HexColor("#367A9B"), colors.HexColor("#C74634")]
    body = ParagraphStyle(
        "assessment_overview_body", fontName="AssessmentJapanese", fontSize=14.0,
        leading=16.4, textColor=muted, wordWrap="CJK",
    )
    bold = ParagraphStyle(
        "assessment_overview_bold", fontName="AssessmentJapaneseBold", fontSize=14.0,
        leading=16.4, textColor=dark, wordWrap="CJK",
    )
    draw_pptx_base_header(
        deck, page_width, page_height, "今回の暫定評価と確認事項",
        assessment_subsection_label("overview"),
    )
    summary = _clean_pptx_display_text(assessment.get("executive_summary"))
    draw_paragraph(
        deck,
        consultative_page_lead(
            "アセスメント概要｜" + (summary or "対象サービスの業務・データを基に、AIが補完しやすい判断と実装条件を整理しています。"),
            title="今回の暫定評価と確認事項",
        ),
        body, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP, wide_content_width(page_width),
    )
    business_value = assessment.get("business_value") if isinstance(assessment.get("business_value"), dict) else {}
    areas = [x for x in business_value.get("impact_areas", []) if isinstance(x, dict)][:3]
    themes = priority_pocs_for(assessment)[:3]
    readiness = assessment.get("poc_start_readiness") if isinstance(assessment.get("poc_start_readiness"), dict) else {}
    open_items = [x for x in readiness.get("items", []) if isinstance(x, dict) and int(x.get("open_gate_count") or 0) > 0]
    def compact_reason(value: object) -> str:
        text = _clean_pptx_display_text(value).rstrip("。")
        text = re.sub(r"^[^、]{1,18}では、", "", text)
        text = re.sub(r"(.+)が[^。]{1,18}に分散しています$", r"\1の情報分散", text)
        text = re.sub(r"を担当者が横断して調整する必要があります$", "の横断調整", text)
        text = re.sub(r"が重要です$", "", text)
        return text.replace("、", "・")

    reason_text = "\n".join(
        f"・{_clean_pptx_display_text(x.get('area'))}：{compact_reason(x.get('business_rationale'))}"
        for x in areas
    ) or "・業務判断、利用可能なデータ、既存機能との差分を確認しています。"
    theme_text = "\n".join(
        f"・{display_theme_name(next((d for d in poc_logic_details_for(assessment) if _matching_source_poc(d, x)), {}), x)}"
        for x in themes
    ) or "・代表技術テーマは、顧客の関心とデータを確認して具体化します。"
    confirm_text = (
        "・対象業務と利用者\n・利用可能な期間・件数・品質・権限\n"
        "・現状値と比較条件\n・人の承認と例外時の戻し方\n・評価責任者と事業化判断者"
        if open_items else
        "・対象範囲、比較条件、責任者を最終確認します。"
    )
    cards = [
        ("暫定評価", "AIは確定処理を置き換えるより、予兆・候補順位・根拠提示を既存業務へ加える形が適合しやすいと考えられます。", theme_text),
        ("判断理由", "受注・在庫・工程・原価などを横断して扱う製品特性から、例外確認と判断支援に接続しやすい構成です。", reason_text),
        ("PoC前に確認したい事項", "現時点では分析仮説を含みます。次の事項をお客様と確認し、最初に扱う代表技術テーマを1つ具体化します。", confirm_text),
    ]
    gap = 6 * mm
    card_w = (wide_content_width(page_width) - gap * 2) / 3
    card_y, card_h = 24 * mm, 101 * mm
    for index, (heading, lead, detail) in enumerate(cards):
        x = WIDE_CONTENT_LEFT + index * (card_w + gap)
        accent = accents[index]
        deck.setFillColor(colors.white); deck.setStrokeColor(colors.HexColor("#D8DEE2")); deck.setLineWidth(0.8)
        deck.roundRect(x, card_y, card_w, card_h, 3 * mm, stroke=1, fill=1)
        deck.setFillColor(accent); deck.rect(x, card_y + card_h - 3 * mm, card_w, 3 * mm, stroke=0, fill=1)
        draw_number_badge(deck, x + 10 * mm, card_y + card_h - 12 * mm, 5 * mm, index + 1, accent, 14.0)
        draw_paragraph(deck, heading, bold, x + 19 * mm, card_y + card_h - 7 * mm, card_w - 25 * mm)
        draw_paragraph(deck, lead, body, x + 6 * mm, card_y + card_h - 25 * mm, card_w - 12 * mm)
        deck.setFillColor(accent); deck.rect(x + 6 * mm, card_y + 8 * mm, card_w - 12 * mm, 0.7 * mm, stroke=0, fill=1)
        draw_paragraph(deck, detail, body, x + 6 * mm, card_y + 53 * mm, card_w - 12 * mm)
    draw_footer(deck, page_width, page_number)


def default_document_page_count(assessment: dict, *, include_cost_estimate: bool = True) -> int:
    """標準PPTXのページ数を、描画順と同じ計画から返す。"""
    return len(default_document_slide_plan(
        assessment, include_cost_estimate=include_cost_estimate,
    ))


def standard_document_page_count(assessment: dict, *, include_cost_estimate: bool = True) -> int:
    """唯一の標準PPTX出力について、描画順と同じページ数を返す。"""
    return default_document_page_count(assessment, include_cost_estimate=include_cost_estimate)


def _qualitative_poc_selection_copy(priority_poc: dict, decision_item: dict,
                                    scorecard_item: dict) -> tuple[str, str, str]:
    """優先PoCを順位や点数ではなく、顧客との対話に使う定性コピーへ変換する。

    選定の再現性は ``poc_priority_decision`` の内部契約で維持する。一方、顧客向けの
    選定ページは採点表ではなく、「サービスにどう適合するか」と「PoCで何を共同で
    具体化するか」を示す。この分離により、個社ごとの事情を数値の優劣へ誤って還元
    しない。
    """
    dimensions = decision_item.get("dimensions") if isinstance(
        decision_item.get("dimensions"), dict
    ) else {}
    score_reasons = scorecard_item.get("score_reasons") if isinstance(
        scorecard_item.get("score_reasons"), dict
    ) else {}

    def dimension_reason(key: str) -> str:
        dimension = dimensions.get(key) if isinstance(dimensions.get(key), dict) else {}
        return _clean_pptx_display_text(
            score_reasons.get(key) or dimension.get("reason")
        )

    theme = _clean_pptx_display_text(priority_poc.get("theme")) or "対象AIユースケース"
    selection_reason = (
        dimension_reason("business_value")
        or _clean_pptx_display_text(priority_poc.get("reason"))
        or f"{theme}が対象サービスの業務価値につながる可能性を具体化しやすいためです。"
    )
    poc_focus = (
        dimension_reason("feasibility")
        or _clean_pptx_display_text(priority_poc.get("depends_on"))
        or "既存業務との連携、人による最終確認、運用上の前提を確認します。"
    )
    preparation_point = (
        _clean_pptx_display_text(scorecard_item.get("data_next_action"))
        or _clean_pptx_display_text(priority_poc.get("first_step"))
        or _clean_pptx_display_text(priority_poc.get("depends_on"))
        or "対象データ、利用範囲、権限、評価条件をお客様と確認しながら整えます。"
    )
    return selection_reason, poc_focus, preparation_point


def draw_default_poc_selection_page(deck: PptCanvas, page_width: float, page_height: float,
                                    assessment: dict, page_number: int) -> None:
    """15件の候補から直接選定したPoC 3テーマと、その理由を定性的に示す。"""
    dark, green = colors.HexColor("#1D252C"), colors.HexColor("#467653")
    muted = colors.HexColor("#4C5961")
    pale_green, card_fill = colors.HexColor("#EEF4F0"), colors.white
    theme_accents = [colors.HexColor(value) for value in POC_THEME_ACCENT_HEX]
    page_title = "15のAIユースケースから整理した AI技術アプローチの代表3テーマ"
    draw_pptx_base_header(deck, page_width, page_height, page_title,
                          assessment_subsection_label("poc_selection"))
    service_name = _clean_pptx_display_text(assessment.get("service_name")) or "対象サービス"
    lead = ParagraphStyle("default_poc_selection_lead", fontName="AssessmentJapanese", fontSize=14,
                          leading=15.6, textColor=colors.HexColor("#4C5961"), wordWrap="CJK")
    draw_paragraph(
        deck,
        consultative_page_lead(
            f"{service_name}の15件を共通の観点で評価し、異なるAI技術方式を比較できる代表3テーマを整理しています。"
            "実際のPoC対象は、お客様の関心領域と利用可能なデータを確認しながら具体化します。",
            title=page_title,
        ),
        lead, WIDE_CONTENT_LEFT, WIDE_EXPLANATORY_LEAD_TOP + 1.2 * mm, wide_content_width(page_width),
    )
    front = consulting_front_matter_for(assessment)
    decision = (
        normalize_poc_priority_decision(assessment.get("poc_priority_decision"), assessment, front)
        or priority_decision_for(assessment, front)
    )
    decision_items = [item for item in decision.get("items", []) if isinstance(item, dict)]
    scorecard = poc_selection_scorecard_for(assessment, front)
    scorecard_by_id = {
        _front_text(item.get("use_case_id"), 8): item
        for item in scorecard.get("items", []) if isinstance(item, dict)
    } if isinstance(scorecard, dict) else {}
    # 15件から3件へ直接選定したことと、全候補へ共通適用した観点を一つの帯で示す。
    band_x, band_y = WIDE_CONTENT_LEFT, 125 * mm
    band_w, band_h = wide_content_width(page_width), 18.5 * mm
    deck.setFillColor(pale_green); deck.roundRect(
        band_x, band_y, band_w, band_h, 2 * mm, stroke=0, fill=1,
    )
    deck.setFillColor(green); deck.setFont("AssessmentJapaneseBold", 14.0)
    deck.drawString(band_x + 5 * mm, band_y + 9.2 * mm, "整理観点")
    criteria = ("業務価値", "既存機能との差分", "データ準備", "比較検証の成立性")
    criteria_x = band_x + 31 * mm
    criteria_gap = 2.5 * mm
    criteria_w = (band_w - 36 * mm - criteria_gap * 3) / 4
    for index, criterion in enumerate(criteria):
        x = criteria_x + index * (criteria_w + criteria_gap)
        deck.setFillColor(colors.white); deck.setStrokeColor(colors.HexColor("#C9D8CF")); deck.setLineWidth(0.7)
        deck.roundRect(x, band_y + 4.0 * mm, criteria_w, 10.5 * mm, 1.8 * mm, stroke=1, fill=1)
        deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 14.0)
        box = deck.slide.shapes.add_textbox(deck._x(x), deck._y(band_y + 14.5 * mm),
                                             deck._width(criteria_w), deck._height(10.5 * mm))
        frame = box.text_frame
        frame.clear(); frame.word_wrap = False
        frame.auto_size = MSO_AUTO_SIZE.NONE
        frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
        paragraph = frame.paragraphs[0]; paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run(); run.text = criterion
        run.font.name = PptCanvas.FONT_MAP["AssessmentJapaneseBold"]
        run.font.size = Pt(14); run.font.bold = True; run.font.color.rgb = deck._rgb(dark)
    # 三つのテーマを横並び・同じ意匠で置き、P1/P2/P3やスコアによる上下関係を
    # 作らない。各カードは「なぜ適合するか」「どこを具体化するか」「最初に何を
    # 確認するか」を示す定性推薦として読む。
    # 日本語の折返しがPowerPoint側で1行増えても、見出しや次の段へ侵入しない
    # ように、カード間の余白を少し詰めて本文幅を確保し、下端までカードを延ばす。
    card_gap = 2.8 * mm
    card_width = (wide_content_width(page_width) - 2 * card_gap) / 3
    card_h, card_y = 107.5 * mm, 14.0 * mm
    title_style = ParagraphStyle(
        "default_poc_selection_card_title", fontName="AssessmentJapaneseBold", fontSize=15.0,
        leading=18.0, textColor=dark, wordWrap="CJK",
    )
    body_style = ParagraphStyle(
        "default_poc_selection_qualitative_body", fontName="AssessmentJapanese", fontSize=14,
        leading=16.2, textColor=muted, wordWrap="CJK",
    )
    for index, item in enumerate(priority_pocs_for(assessment, front)[:3]):
        x = WIDE_CONTENT_LEFT + index * (card_width + card_gap)
        # 色は優先順位や評価の優劣ではなく、後続するPoC詳細ページの同じテーマを
        # たどるための視覚的な対応付けとしてだけ用いる。
        theme_accent = theme_accents[index]
        deck.setFillColor(card_fill); deck.setStrokeColor(colors.HexColor("#D8DEE2")); deck.setLineWidth(0.8)
        deck.roundRect(x, card_y, card_width, card_h, 2 * mm, stroke=1, fill=1)
        deck.setFillColor(theme_accent); deck.rect(
            x, card_y + card_h - 4.5 * mm, card_width, 4.5 * mm, stroke=0, fill=1,
        )
        column_heading_style = ParagraphStyle(
            f"default_poc_selection_column_heading_{index}", fontName="AssessmentJapaneseBold", fontSize=14,
            leading=16.2, textColor=theme_accent, wordWrap="CJK",
        )
        # 日本語本文の最終句読点だけが次行へ落ちるケースを避けるため、左右余白を
        # 3mmにそろえて可読幅を確保する。14ptと全文は維持する。
        content_x, content_width = x + 3 * mm, card_width - 6 * mm
        draw_number_badge(deck, content_x + 4.0 * mm, card_y + card_h - 13.0 * mm,
                          4.0 * mm, index + 1, theme_accent, 14.0)
        detail = next((row for row in poc_logic_details_for(assessment)
                       if _matching_source_poc(row, item)), {})
        modality_label = technical_modality_for(item, detail).get("label") or "AI技術"
        draw_paragraph(
            deck, f"技術方式｜{modality_label}", column_heading_style,
            content_x + 9.5 * mm, card_y + card_h - 8.3 * mm, content_width - 9.5 * mm,
        )
        title_bottom = draw_paragraph(
            deck, _pptx_summary_text(display_theme_name(detail, item), 72) or "ご提案テーマ", title_style,
            content_x + 9.5 * mm, card_y + card_h - 16.4 * mm, content_width - 9.5 * mm,
        )
        decision_item = next((row for row in decision_items
                              if str(row.get("use_case_id") or "") == str(item.get("use_case_id") or "")
                              or str(row.get("priority") or "") == f"P{index + 1}"), {})
        use_case_id = _front_text(item.get("use_case_id"), 8)
        scorecard_item = scorecard_by_id.get(use_case_id, {})
        selection_reason, poc_focus, preparation_point = _qualitative_poc_selection_copy(
            item, decision_item, scorecard_item,
        )
        preparation_point = re.sub(
            r"^(?:対象[^、]{1,16}を棚卸し、|過去[^、]{1,16}について、)", "", preparation_point,
        ).replace("文書アクセス権", "アクセス権")
        current_y = title_bottom - 1.0 * mm
        for heading, body in (
            ("代表とする理由", selection_reason),
            ("PoCで確かめること", poc_focus),
            ("準備のポイント", preparation_point),
        ):
            heading_bottom = draw_paragraph(
                deck, heading, column_heading_style, content_x, current_y, content_width,
            )
            body_bottom = draw_paragraph(
                deck, body, body_style, content_x, heading_bottom - 0.7 * mm, content_width,
            )
            current_y = body_bottom - 0.8 * mm
    draw_footer(deck, page_width, page_number)


def assessment_intro_copy_for(assessment: dict) -> dict:
    """P2の説明文を、案件の会社・サービス情報から組み立てる。

    P2は全案件共通のアセスメント方法を説明するページだが、固定の「ISV様」や
    特定製品名を埋め込まない。対象サービスの名称・業務種別を差し込み、後続の
    ユースケース選定、PoC、OCI実証へ自然につながる顧客向けコピーにする。
    """
    company_name = re.sub(
        r"(?:様|御中)$", "", str(assessment.get("company_name") or "ご提案先企業").strip(),
    ).strip()
    formal_company_name = company_name
    company_name = re.sub(r"^(?:株式会社|有限会社|合同会社)", "", company_name).strip()
    service_name = str(assessment.get("service_name") or "対象サービス").strip()
    groups = assessment.get("service_use_case_groups")
    first_group = groups[0] if isinstance(groups, list) and groups and isinstance(groups[0], dict) else {}
    service_context = str(
        first_group.get("service_type")
        or assessment.get("service_genre")
        or service_name
        or "対象サービス"
    ).strip()
    compact_context = _clean_pptx_display_text(service_context)
    return {
        "title": "AI Use Case Assessment",
        "subtitle": f"{formal_company_name}様のAI構築をご支援",
        "blocks": [
            {
                "title": "AIのビジネス価値を共同で構築",
                "description": (
                    f"{company_name}様と、{service_name}の顧客体験、業務品質、生産性、事業成長への効果を整理し、"
                    "AIで変える判断と評価指標を具体化します。"
                ),
            },
            {
                "title": "AIユースケースと成功条件を共同で設計",
                "description": (
                    f"{service_name}の利用者、業務フロー、利用可能なデータを起点に候補を比較し、"
                    "検証対象、評価指標、PoCの成功条件を具体化します。"
                ),
            },
            {
                "title": "OCI環境でPoCを設計・検証",
                "description": (
                    "選定テーマを対象に、業務連携・アプリケーション層、Autonomous AI Database、"
                    "OCI Generative AIを組み合わせ、実装方式と評価結果を次の判断へつなげます。"
                ),
            },
        ],
        "image_banner": f"{compact_context}のAI活用を設計・実証",
    }


def assessment_document_title(assessment: dict) -> str:
    """Return the customer-facing package/PDF title for an assessment deck."""
    service_name = _clean_pptx_display_text(
        assessment.get("service_name") or "対象サービス"
    )
    return (
        "AI Use Case Assessment｜"
        f"{service_name}におけるAI活用・PoCのご提案"
    )


def draw_default_assessment_intro_page(
    deck: PptCanvas, page_width: float, page_height: float, assessment: dict,
) -> None:
    """参照画像の左右分割構成を、ネイティブPPTX図形としてP2へ描く。"""
    if not ASSESSMENT_INTRO_IMAGE.is_file():
        raise FileNotFoundError(ASSESSMENT_INTRO_IMAGE)

    copy_data = assessment_intro_copy_for(assessment)
    dark = colors.HexColor("#252B2B")
    muted = colors.HexColor("#353B3B")
    split_x = page_width / 2
    left_x = 20 * mm
    left_width = split_x - left_x - 14 * mm

    # 960×1080の縦長写真を右半面へ等倍比率で配置する。16:9の半面と同じ
    # 8:9比率なのでトリミングや引き伸ばしは発生しない。
    deck.drawImage(
        ImageReader(str(ASSESSMENT_INTRO_IMAGE)), split_x, 0,
        width=page_width - split_x, height=page_height,
        preserveAspectRatio=True,
    )

    deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 24)
    deck.drawString(left_x, page_height - 25 * mm, str(copy_data["title"]))
    subtitle_style = ParagraphStyle("assessment_intro_subtitle", fontName="AssessmentJapanese",
                                    fontSize=16, leading=19, textColor=dark, wordWrap="CJK")
    draw_paragraph(deck, copy_data["subtitle"], subtitle_style, left_x,
                   page_height - 35 * mm, left_width)

    heading_style = ParagraphStyle(
        "assessment_intro_heading", fontName="AssessmentJapaneseBold", fontSize=16,
        leading=18.8, textColor=dark, wordWrap="CJK",
    )
    body_style = ParagraphStyle(
        "assessment_intro_body", fontName="AssessmentJapanese", fontSize=12,
        leading=15.2, textColor=muted, wordWrap="CJK",
    )
    block_heading_y = (page_height - 55 * mm, page_height - 95 * mm, page_height - 135 * mm)
    for block, heading_y in zip(copy_data["blocks"], block_heading_y):
        description_y = draw_paragraph(
            deck, str(block["title"]), heading_style, left_x, heading_y, left_width,
        )
        draw_paragraph(
            deck, str(block["description"]), body_style,
            left_x, description_y - 3 * mm, left_width,
        )

    # 写真内の元プレースホルダーを完全に覆う帯。文字は案件のサービス種別を
    # 差し込み、短い一文に保って一行で中央配置する。
    banner_y, banner_height = 71.25 * mm, 19 * mm
    deck.setFillColor(colors.HexColor("#252B2B"))
    deck.rect(split_x, banner_y, page_width - split_x, banner_height, stroke=0, fill=1)
    deck.setFillColor(colors.white); deck.setFont("AssessmentJapanese", 18)
    deck._text(
        str(copy_data["image_banner"]), split_x,
        banner_y + banner_height / 2 - 1.7 * mm,
        page_width - split_x, PP_ALIGN.CENTER,
    )


def create_default_assessment_document(assessment: dict, architecture_image: Path,
                                       cost_estimate: dict | None = None,
                                       research: dict | None = None) -> PptCanvas:
    """通常実行で出す、営業・初回提案向けの標準AIアセスメントを生成する。"""
    register_japanese_font()
    page_width, page_height = PPTX_WIDESCREEN_WIDTH, PPTX_WIDESCREEN_HEIGHT
    deck = PptCanvas(page_width, page_height)
    company_name = str(assessment.get("company_name") or "ご提案先企業様")
    service_name = str(assessment.get("service_name") or "対象サービス")
    # PowerPointからPDFへ書き出した際の文書タイトルは、このcore propertyを
    # 引き継ぐ。空欄やOracleベースの旧Sales Session名を残さない。
    deck.presentation.core_properties.title = assessment_document_title(assessment)
    deck.presentation.core_properties.subject = "AI Use Case Assessment"
    dark, green, orange = colors.HexColor("#1D252C"), colors.HexColor("#467653"), colors.HexColor("#C74634")
    body = ParagraphStyle("default_body", fontName="AssessmentJapanese", fontSize=10,
                          leading=12.8, textColor=colors.HexColor("#4C5961"), wordWrap="CJK")
    slide_plan = default_document_slide_plan(
        assessment, include_cost_estimate=bool(cost_estimate),
    )

    def expect_slide(key: str, page_number: int) -> None:
        if page_number < 1 or page_number > len(slide_plan):
            raise RuntimeError(f"スライド計画外のページ番号です: {page_number} / {key}")
        expected = slide_plan[page_number - 1]
        if expected.key != key:
            raise RuntimeError(
                f"スライド順序が計画と一致しません: page={page_number}, "
                f"expected={expected.key}, actual={key}"
            )

    def finish(page_number: int) -> None:
        draw_footer(deck, page_width, page_number)
        deck.showPage()

    # 1. 表紙
    expect_slide("cover", 1)
    deck.setFillColor(colors.HexColor("#F8F8F7")); deck.rect(0, 0, page_width, page_height, stroke=0, fill=1)
    cover_reader = ImageReader(str(COVER_BACKGROUND_IMAGE))
    cover_width, cover_height = cover_reader.getSize()
    cover_scale = max(page_width / cover_width, page_height / cover_height)
    deck.drawImage(cover_reader, (page_width - cover_width * cover_scale) / 2,
                   (page_height - cover_height * cover_scale) / 2,
                   width=cover_width * cover_scale, height=cover_height * cover_scale,
                   preserveAspectRatio=True)
    logo_reader = ImageReader(str(ORACLE_LOGO_IMAGE)); logo_width, logo_height = logo_reader.getSize()
    logo_draw_width = 40 * mm
    deck.drawImage(logo_reader, 18 * mm, page_height - 26 * mm, width=logo_draw_width,
                   height=logo_draw_width * logo_height / logo_width, preserveAspectRatio=True, mask="auto")
    deck.setFillColor(dark); deck.setFont("AssessmentJapaneseBold", 28)
    deck.drawString(24 * mm, page_height - 67 * mm, "AI Use Case Assessment")
    deck.setFillColor(orange); deck.rect(24 * mm, page_height - 74 * mm, 75 * mm, 1.8 * mm, stroke=0, fill=1)
    deck.setFillColor(dark); deck.setFont("AssessmentJapanese", 18); deck.drawString(24 * mm, page_height - 88 * mm, "AI活用アセスメント報告書")
    deck.setFont("AssessmentJapaneseBold", 20); deck.drawString(24 * mm, page_height - 116 * mm, f"{company_name}　御中")
    deck.setFillColor(green); deck.roundRect(24 * mm, 36 * mm, 150 * mm, 28 * mm, 2.5 * mm, stroke=0, fill=1)
    deck.setFillColor(colors.HexColor("#DCE9E1")); deck.setFont("AssessmentJapanese", 9.5); deck.drawString(31 * mm, 54 * mm, "対象サービス")
    deck.setFillColor(colors.white); deck.setFont("AssessmentJapaneseBold", 16); deck.drawString(31 * mm, 43 * mm, service_name)
    deck.showPage()

    # 2. 既存PPTXを移植せず、共通写真と編集可能なネイティブ図形だけで
    # 「OCI AI Use Case Assessmentとは」を構成する。
    expect_slide("intro", 2)
    draw_default_assessment_intro_page(deck, page_width, page_height, assessment)
    finish(2)

    page = 3
    midterm = assessment.get("midterm_plan_analysis")
    if isinstance(midterm, dict) and midterm.get("ai_alignment"):
        expect_slide("midterm", page)
        draw_midterm_plan_alignment_page(deck, page_width, page_height, midterm, page,
                                         section=assessment_subsection_label("plan_alignment"))
        deck.showPage(); page += 1

    # 中計の有無によらず、公開情報による業界動向と、対象サービスへ
    # 接続するビジネスインパクトをユースケース一覧の直前に並べる。
    expect_slide("industry_value", page)
    draw_industry_trends_page(deck, page_width, page_height, assessment, research, page)
    deck.showPage(); page += 1

    expect_slide("business_impact", page)
    draw_assessment_business_impact_page(
        deck, page_width, page_height, assessment, page, research,
    )
    deck.showPage(); page += 1

    # 全候補を先に示し、その次に優先3テーマを明示する。
    groups = service_use_case_groups_for(assessment)
    for service_index, group in enumerate(groups, 1):
        chunks = use_case_list_chunks(group)
        for list_page_index, cases in enumerate(chunks, 1):
            expect_slide(f"use_case_list:{service_index}:{list_page_index}", page)
            draw_use_case_list_page(deck, page_width, page_height, group, cases, page, len(groups), service_index,
                                    list_page_index, len(chunks), appendix=False,
                                    priority_pocs=priority_pocs_for(assessment))
            deck.showPage(); page += 1
    expect_slide("poc_selection", page)
    draw_default_poc_selection_page(deck, page_width, page_height, assessment, page)
    deck.showPage(); page += 1

    for priority, (detail, poc) in enumerate(zip(poc_logic_details_for(assessment), priority_pocs_for(assessment)), 1):
        expect_slide(f"poc_detail:{priority}", page)
        draw_poc_logic_detail_page(deck, page_width, page_height, detail, poc, priority, page)
        deck.showPage(); page += 1
    expect_slide("architecture", page)
    draw_fixed_architecture_reference_page(deck, page_width, page_height, assessment, architecture_image, page)
    deck.showPage(); page += 1
    if cost_estimate:
        expect_slide("cost", page)
        draw_cost_estimate_page(deck, page_width, page_height, cost_estimate, page)
        deck.showPage(); page += 1
    expect_slide("support", page)
    draw_poc_support_page(deck, page_width, page_height, assessment, page)
    deck.showPage()
    # 提案の余韻を残すOracleロゴのクロージングを必ず最終ページに置く。
    expect_slide("closing", page + 1)
    draw_oracle_closing_page(deck, page_width, page_height)
    if len(deck.presentation.slides) != len(slide_plan):
        raise RuntimeError(
            f"生成スライド数が計画と一致しません: expected={len(slide_plan)}, "
            f"actual={len(deck.presentation.slides)}"
        )
    deck.save()
    return deck


def create_pptx(assessment: dict, architecture_image: Path, output_path: Path,
                cost_estimate: dict | None = None,
                research: dict | None = None) -> None:
    """標準の編集可能な16:9 AIアセスメントPPTXを出力する。"""
    deck = create_default_assessment_document(
        assessment, architecture_image, cost_estimate, research,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    deck.write(output_path)
