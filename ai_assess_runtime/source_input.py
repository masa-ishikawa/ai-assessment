"""TXT/XLSX入力の解決・抽出・正規化。"""

import re
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree

from .paths import INPUT_DIR, PROJECT_DIR

INPUT_TEMPLATE_SHEET = "サービス入力"
ISV_INPUT_SHEET_NAMES = ("入力ファイル", "入力フォーム")
ISV_NON_ASSESSMENT_QUESTION_IDS = frozenset({
    "F001", "F002", "F107", "F108", "F109", "F1010",
})
EXCEL_INPUT_FIELDS = (
    "会社名",
    "サービス名",
    "サービス概要",
    "現在取り組んでいるAI",
    "AIに期待すること",
)
REQUIRED_EXCEL_INPUT_FIELDS = EXCEL_INPUT_FIELDS[:3]

MAX_TEXT_INPUT_BYTES = 2 * 1024 * 1024
MAX_XLSX_INPUT_BYTES = 25 * 1024 * 1024
MAX_XLSX_MEMBER_BYTES = 25 * 1024 * 1024
MAX_XLSX_UNCOMPRESSED_BYTES = 150 * 1024 * 1024
MAX_XLSX_MEMBERS = 10_000
MAX_XLSX_COMPRESSION_RATIO = 1_000


def _validate_input_size(input_file: Path, max_bytes: int, label: str) -> None:
    """入力をメモリへ展開する前に、圧縮前ファイルサイズを制限する。"""
    try:
        size = input_file.stat().st_size
    except OSError as error:
        raise ValueError(f"{label}を読み取れません: {input_file}") from error
    if size > max_bytes:
        raise ValueError(f"{label}が許容サイズ（{max_bytes // (1024 * 1024)}MB）を超えています。")


def _validate_xlsx_archive(
    workbook_zip: zipfile.ZipFile,
    *,
    max_members: int = MAX_XLSX_MEMBERS,
    max_member_bytes: int = MAX_XLSX_MEMBER_BYTES,
    max_uncompressed_bytes: int = MAX_XLSX_UNCOMPRESSED_BYTES,
    max_compression_ratio: int = MAX_XLSX_COMPRESSION_RATIO,
) -> None:
    """XLSXの異常な展開量、暗号化、過剰圧縮を読み込み前に拒否する。"""
    members = workbook_zip.infolist()
    if len(members) > max_members:
        raise ValueError("Excel内のファイル数が許容上限を超えています。")
    total = 0
    for member in members:
        if member.flag_bits & 0x1:
            raise ValueError("暗号化されたExcelファイルは読み込めません。")
        if member.file_size > max_member_bytes:
            raise ValueError("Excel内のデータが許容サイズを超えています。")
        total += member.file_size
        if total > max_uncompressed_bytes:
            raise ValueError("Excelの展開後サイズが許容上限を超えています。")
        if member.file_size and member.compress_size == 0:
            raise ValueError("Excel内の圧縮情報が不正です。")
        if member.compress_size and member.file_size / member.compress_size > max_compression_ratio:
            raise ValueError("Excel内に過剰圧縮されたデータがあります。")

def safe_filename(name: str) -> str:
    sanitized = re.sub(r"[\\/:*?\"<>|]", "_", name).strip()
    return re.sub(r"\s+", "_", sanitized) or "ai_assessment"


def extract_company_name(source_text: str) -> str:
    """自由記述の入力から、表紙に表示する会社名を取得する。"""
    pattern = r"^\s*(?:会社名|企業名|顧客名|お客様名|提案先)\s*[:：]\s*(.+?)\s*$"
    match = re.search(pattern, source_text, flags=re.MULTILINE)
    if match:
        return match.group(1).strip()

    # タイトル（例: 「# 株式会社サンプル：AI Assess用…」）を次善の明示情報として使う。
    match = re.search(r"^\s*#+\s*(株式会社[^：:\n]+)", source_text, flags=re.MULTILINE)
    if match:
        return match.group(1).strip()
    return "ご提案先企業様"


def derive_display_company_name(source_text: str, fallback: str) -> str:
    """規則で取れる社名を優先し、自由記述ではLLMが抽出した社名へフォールバックする。"""
    extracted = extract_company_name(source_text)
    if extracted != "ご提案先企業様":
        return extracted
    fallback = fallback.strip()
    return fallback if fallback else extracted


def extract_target_service_name(source_text: str) -> str | None:
    """入力で明示された表紙用の対象サービス名・製品名を取得する。"""
    pattern = r"^\s*(?:対象サービス名|対象サービス|サービス名|製品名)\s*[:：]\s*(.+?)\s*$"
    match = re.search(pattern, source_text, flags=re.MULTILINE)
    return match.group(1).strip() if match else None


def extract_target_services(source_text: str) -> list[str]:
    """明示名または「AI Assessの対象範囲」等の番号付きリストから対象サービス名を抽出する。"""
    explicit_name = extract_target_service_name(source_text)
    if explicit_name:
        return [explicit_name]

    section_match = re.search(
        r"^\s*##\s*.*?(?:対象範囲|対象サービス).*?\n(.*?)(?=^\s*##\s|\Z)",
        source_text,
        flags=re.MULTILINE | re.DOTALL,
    )
    text = section_match.group(1) if section_match else source_text
    services: list[str] = []
    for line in text.splitlines():
        match = re.match(r"^\s*\d+[.．、]\s*(.+?)\s*$", line)
        if not match:
            continue
        item = match.group(1)
        quoted = re.search(r"[「\"]([^」\"]+)[」\"]", item)
        name = quoted.group(1) if quoted else re.split(r"[（(]", item, maxsplit=1)[0]
        name = re.sub(r"^(?:既存の|新規の)", "", name).strip(" ・、。")
        if name:
            services.append(name)
    return services


def derive_display_service_name(source_text: str, fallback: str) -> str:
    """入力の優先度・対象一覧から、表紙用の代表サービス名を動的に決定する。"""
    explicit_name = extract_target_service_name(source_text)
    if explicit_name:
        return explicit_name

    target_services = extract_target_services(source_text)
    # 「最優先の既存データ活用対象」のように明記された見出しを、代表サービスとして優先する。
    priority_match = re.search(r"^\s*#{2,6}\s*(?:[A-Z][.．]\s*)?(.+?)（最優先[^）]*）", source_text, flags=re.MULTILINE)
    if priority_match:
        representative = priority_match.group(1).strip()
    elif target_services:
        representative = target_services[0]
    else:
        return fallback

    # 複数対象の場合は、代表例であることが分かるよう必ず「等」を付与する。
    return representative if len(target_services) <= 1 or representative.endswith("等") else f"{representative} 等"


def extract_service_genre(source_text: str, fallback: str) -> str:
    """入力の明示値を優先して、一覧表の表題に用いるサービスジャンルを取得する。"""
    genre_pattern = r"^\s*(?:サービスジャンル|業種・ジャンル|サービス種別)\s*[:：]\s*(.+?)\s*$"
    match = re.search(genre_pattern, source_text, flags=re.MULTILINE)
    if match:
        return match.group(1).strip()

    domain_pattern = r"^\s*サービス領域\s*[:：]?\s*\n\s*(.+?)\s*$"
    match = re.search(domain_pattern, source_text, flags=re.MULTILINE)
    if match:
        description = match.group(1).strip()
        genre = re.split(r"(?:を支援する|を提供する|サービスです|SaaS)", description, maxsplit=1)[0].strip(" ・、。")
        if genre:
            return genre
    return fallback


def _excel_column(cell_reference: str) -> str:
    """Excelセル参照から列記号を取り出す。"""
    match = re.match(r"([A-Z]+)", cell_reference)
    return match.group(1) if match else ""


def _xlsx_cell_text(cell: ElementTree.Element, shared_strings: list[str], namespace: str) -> str:
    """xlsxのセル値を、共有文字列・inlineStrを含めて文字列として返す。"""
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.iter(f"{{{namespace}}}t")).strip()

    value_node = cell.find(f"{{{namespace}}}v")
    if value_node is None or value_node.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value_node.text)].strip()
        except (IndexError, ValueError):
            raise ValueError("Excelの共有文字列を読み取れません。指定テンプレートを使ってください。")
    return value_node.text.strip()


def read_xlsx_rows(input_file: Path, accepted_sheet_names: tuple[str, ...]) -> tuple[str, dict[int, dict[str, str]]]:
    """Read an .xlsx sheet into row/column text without depending on Excel automation."""
    _validate_input_size(input_file, MAX_XLSX_INPUT_BYTES, "Excel入力")
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    try:
        with zipfile.ZipFile(input_file) as workbook_zip:
            _validate_xlsx_archive(workbook_zip)
            workbook_xml = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
            relationships_xml = ElementTree.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
            relationships = {relation.get("Id"): relation.get("Target") for relation in relationships_xml.findall(f"{{{package_rel_ns}}}Relationship")}
            sheet = next((item for item in workbook_xml.findall(f".//{{{main_ns}}}sheet") if item.get("name") in accepted_sheet_names), None)
            if sheet is None:
                names = "、".join(f"「{name}」" for name in accepted_sheet_names)
                raise ValueError(f"シート名{names}が見つかりません。指定テンプレートを使用してください。")
            relation_id = sheet.get(f"{{{office_rel_ns}}}id")
            target = relationships.get(relation_id)
            if not target:
                raise ValueError("Excelのシート情報を読み取れません。指定テンプレートを使用してください。")
            sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in workbook_zip.namelist():
                shared_xml = ElementTree.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
                shared_strings = ["".join(node.text or "" for node in item.iter(f"{{{main_ns}}}t")) for item in shared_xml.findall(f"{{{main_ns}}}si")]
            sheet_xml = ElementTree.fromstring(workbook_zip.read(sheet_path))
    except (KeyError, zipfile.BadZipFile, ElementTree.ParseError) as error:
        raise ValueError("有効な .xlsx ファイルを指定してください。") from error
    rows: dict[int, dict[str, str]] = {}
    for cell in sheet_xml.findall(f".//{{{main_ns}}}c"):
        reference = cell.get("r", "")
        row_match = re.search(r"(\d+)$", reference)
        if row_match:
            rows.setdefault(int(row_match.group(1)), {})[_excel_column(reference)] = _xlsx_cell_text(cell, shared_strings, main_ns)
    return str(sheet.get("name")), rows


def extract_isv_assessment_input(input_file: Path) -> dict:
    """Extract every ISV form answer while retaining the required-answer contract.

    The assessment generator historically exposed only ``required_answers``.  Keep
    that key for compatibility, and add ``answers`` so optional customer context is
    not discarded before the LLM analysis step.
    """
    sheet_name, rows = read_xlsx_rows(input_file, ISV_INPUT_SHEET_NAMES)
    questions: list[dict[str, str | int | bool]] = []
    for row_number, row in sorted(rows.items()):
        requirement = unicodedata.normalize("NFKC", row.get("C", "")).strip()
        question = row.get("D", "").strip()
        question_id = row.get("B", "").strip()
        if requirement not in {"必須", "任意"} or not question or not question_id or question_id == "質問ID":
            continue
        questions.append({
            "row": row_number,
            "question_id": question_id,
            "required": requirement == "必須",
            "question": question,
            "answer": row.get("E", "").strip(),
        })
    required_questions = [item for item in questions if item["required"]]
    return {
        "format": "isv-ai-use-case-assessment-input/v1",
        "sheet_name": sheet_name,
        "answers": questions,
        "required_answers": required_questions,
        "missing_required": [item["question"] for item in required_questions if not item["answer"]],
    }


def normalize_isv_assessment_context(extracted: dict) -> dict:
    """ISVフォームを、個人情報・同意回答を除いた分析用コンテキストへ正規化する。"""
    answer_items = extracted.get("answers")
    if not isinstance(answer_items, list):
        # v1拡張前の抽出JSONも後方互換で読み込む。
        answer_items = extracted.get("required_answers", [])
    answers = {
        str(item.get("question_id", "")).strip(): str(item.get("answer", "")).strip()
        for item in answer_items
        if isinstance(item, dict) and str(item.get("question_id", "")).strip() not in ISV_NON_ASSESSMENT_QUESTION_IDS
    }
    questions = {
        str(item.get("question_id", "")).strip(): str(item.get("question", "")).strip()
        for item in answer_items
        if isinstance(item, dict) and str(item.get("question_id", "")).strip() not in ISV_NON_ASSESSMENT_QUESTION_IDS
    }

    def value(question_id: str) -> str:
        return answers.get(question_id, "")

    projects: dict[str, dict[str, str]] = {}
    priority_use_cases: dict[str, dict[str, str]] = {}
    consumed_question_ids = {
        "F101", "F103", "F104", "F201", "F301", "F304", "F305", "F501", "F601", "F602",
    }
    for question_id, question in questions.items():
        project_match = re.match(r"プロジェクト(\d+)：(.+)", question)
        if project_match:
            consumed_question_ids.add(question_id)
            project = projects.setdefault(project_match.group(1), {})
            label = project_match.group(2)
            if "製品名" in label or "プロジェクト" in label:
                project["name"] = value(question_id)
            elif "段階" in label:
                project["stage"] = value(question_id)
            elif "優先度" in label:
                project["priority"] = value(question_id)
            elif "狙い" in label or "確認したい" in label:
                project["ai_goal"] = value(question_id)
            continue
        use_case_match = re.match(r"ユースケース(\d+)：(.+)", question)
        if use_case_match:
            consumed_question_ids.add(question_id)
            use_case = priority_use_cases.setdefault(use_case_match.group(1), {})
            label = use_case_match.group(2)
            if "名称" in label:
                use_case["name"] = value(question_id)
            elif "業務領域" in label:
                use_case["business_domain"] = value(question_id)
            elif "優先度" in label:
                use_case["priority"] = value(question_id)
            elif "利用シナリオ" in label:
                use_case["scenario"] = value(question_id)
            elif "必要データ" in label:
                use_case["data"] = value(question_id)
            elif "想定AI技術" in label:
                use_case["ai_technology"] = value(question_id)
            elif "期待効果" in label or "KPI" in label:
                use_case["expected_impact"] = value(question_id)

    additional_answers = [
        {
            "question_id": question_id,
            "question": questions[question_id],
            "answer": answer,
        }
        for question_id, answer in answers.items()
        if answer and question_id not in consumed_question_ids
    ]
    required_answer_count = sum(
        1 for item in answer_items
        if isinstance(item, dict)
        and bool(item.get("required"))
        and str(item.get("question_id", "")).strip() not in ISV_NON_ASSESSMENT_QUESTION_IDS
        and str(item.get("answer", "")).strip()
    )

    return {
        "format": "isv-ai-use-case-assessment-context/v1",
        "source": {
            "sheet_name": str(extracted.get("sheet_name", "")),
            "required_answer_count": required_answer_count,
            "assessment_answer_count": sum(1 for answer in answers.values() if answer),
        },
        "company": {
            "name": value("F101"),
            "business_overview": value("F103"),
            "target_customers": value("F104"),
        },
        "assessment_scope": {
            "scope": value("F201"),
            "projects": [project for _, project in sorted(projects.items()) if any(project.values())],
        },
        "service": {
            "name": value("F301"),
            "business_functions": value("F304"),
            "data_domains": value("F305"),
            "business_ip": value("F501"),
        },
        "ai_objectives": {"business_value": value("F601"), "desired_capability": value("F602")},
        "priority_use_cases": [use_case for _, use_case in sorted(priority_use_cases.items()) if any(use_case.values())],
        "additional_answers": additional_answers,
    }


def build_isv_source_text(context: dict) -> str:
    """正規化済みISVコンテキストを、既存の分析・PPTX生成経路へ渡す本文にする。"""
    company = context.get("company", {})
    scope = context.get("assessment_scope", {})
    service = context.get("service", {})
    objectives = context.get("ai_objectives", {})
    sections = [
        f"会社名: {company.get('name', '')}",
        f"サービス名: {service.get('name', '')}",
        "製品概要:\n" + "\n".join(part for part in (
            str(company.get("business_overview", "")).strip(),
            f"主な顧客・利用部門: {company.get('target_customers', '')}" if company.get("target_customers") else "",
            f"代表的な機能・業務範囲: {service.get('business_functions', '')}" if service.get("business_functions") else "",
            f"関連データ領域: {service.get('data_domains', '')}" if service.get("data_domains") else "",
            f"保有データ・業務IP: {service.get('business_ip', '')}" if service.get("business_ip") else "",
        ) if part),
    ]
    if scope.get("scope"):
        sections.append(f"今回のAssessment対象範囲: {scope['scope']}")
    for project in scope.get("projects", []):
        details = " / ".join(f"{label}: {project[key]}" for key, label in (
            ("name", "製品・プロジェクト"), ("stage", "段階"), ("priority", "優先度"), ("ai_goal", "AI活用の狙い"),
        ) if project.get(key))
        if details:
            sections.append("対象プロジェクト: " + details)
    if objectives.get("business_value"):
        sections.append(f"AI活用目的: {objectives['business_value']}")
    if objectives.get("desired_capability"):
        sections.append(f"AIで実現したいこと: {objectives['desired_capability']}")
    use_cases = context.get("priority_use_cases", [])
    if use_cases:
        lines = []
        for use_case in use_cases:
            lines.append(" / ".join(f"{label}: {use_case[key]}" for key, label in (
                ("name", "名称"), ("business_domain", "業務領域"), ("priority", "優先度"),
                ("scenario", "利用シナリオ"), ("data", "必要データ"),
                ("ai_technology", "想定AI技術"), ("expected_impact", "期待効果/KPI"),
            ) if use_case.get(key)))
        sections.append("顧客指定の優先ユースケース（必ず提案・PoC候補へ反映）:\n- " + "\n- ".join(lines))
    additional_answers = context.get("additional_answers", [])
    if isinstance(additional_answers, list):
        lines = [
            f"- {item.get('question', '')}: {item.get('answer', '')}"
            for item in additional_answers
            if isinstance(item, dict) and str(item.get("question", "")).strip() and str(item.get("answer", "")).strip()
        ]
        if lines:
            sections.append("追加ヒアリング回答（顧客入力）:\n" + "\n".join(lines))
    return "\n\n".join(section for section in sections if section.strip())


def load_isv_assessment_context(input_file: Path) -> dict | None:
    """ISVフォームなら正規化コンテキストを返し、旧形式ならNoneを返す。"""
    try:
        extracted = extract_isv_assessment_input(input_file)
    except ValueError:
        return None
    missing = list(extracted["missing_required"])
    if missing:
        raise ValueError("Excelの必須設問が未入力です: " + "、".join(str(item) for item in missing))
    return normalize_isv_assessment_context(extracted)


def load_excel_service_input(input_file: Path) -> str:
    """固定ExcelテンプレートからLLMへ渡すサービス説明を組み立てる。"""
    _validate_input_size(input_file, MAX_XLSX_INPUT_BYTES, "Excel入力")
    # 新しいISV入力フォームは、必須行を構造化してからAIへ渡す。
    isv_context = load_isv_assessment_context(input_file)
    if isv_context is not None:
        return build_isv_source_text(isv_context)

    # 旧サービス入力テンプレートも既存案件のため読み込みを維持する。
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    try:
        with zipfile.ZipFile(input_file) as workbook_zip:
            _validate_xlsx_archive(workbook_zip)
            workbook_xml = ElementTree.fromstring(workbook_zip.read("xl/workbook.xml"))
            relationships_xml = ElementTree.fromstring(workbook_zip.read("xl/_rels/workbook.xml.rels"))
            relationships = {
                relation.get("Id"): relation.get("Target")
                for relation in relationships_xml.findall(f"{{{package_rel_ns}}}Relationship")
            }
            sheet = next(
                (item for item in workbook_xml.findall(f".//{{{main_ns}}}sheet")
                 if item.get("name") == INPUT_TEMPLATE_SHEET),
                None,
            )
            if sheet is None:
                raise ValueError(f"シート名「{INPUT_TEMPLATE_SHEET}」が見つかりません。指定テンプレートを使用してください。")
            relation_id = sheet.get(f"{{{office_rel_ns}}}id")
            target = relationships.get(relation_id)
            if not target:
                raise ValueError("Excelのシート情報を読み取れません。指定テンプレートを使用してください。")
            sheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target

            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in workbook_zip.namelist():
                shared_xml = ElementTree.fromstring(workbook_zip.read("xl/sharedStrings.xml"))
                shared_strings = [
                    "".join(node.text or "" for node in item.iter(f"{{{main_ns}}}t"))
                    for item in shared_xml.findall(f"{{{main_ns}}}si")
                ]
            sheet_xml = ElementTree.fromstring(workbook_zip.read(sheet_path))
    except (KeyError, zipfile.BadZipFile, ElementTree.ParseError) as error:
        raise ValueError("有効な .xlsx ファイルを指定してください。") from error

    rows: dict[int, dict[str, str]] = {}
    for cell in sheet_xml.findall(f".//{{{main_ns}}}c"):
        reference = cell.get("r", "")
        row_match = re.search(r"(\d+)$", reference)
        if not row_match:
            continue
        rows.setdefault(int(row_match.group(1)), {})[_excel_column(reference)] = _xlsx_cell_text(
            cell, shared_strings, main_ns
        )

    values: dict[str, str] = {}
    for row in rows.values():
        # テンプレートの表示用注記（例:「会社名\n（必須）」）は判定対象から除く。
        field_name = re.split(r"[\n（(]", row.get("A", "").strip(), maxsplit=1)[0].strip()
        if field_name in EXCEL_INPUT_FIELDS:
            values[field_name] = row.get("B", "").strip()

    missing = [field for field in REQUIRED_EXCEL_INPUT_FIELDS if not values.get(field)]
    if missing:
        raise ValueError("Excelの必須項目が未入力です: " + "、".join(missing))

    sections = [
        f"会社名: {values['会社名']}",
        f"サービス名: {values['サービス名']}",
        "製品概要:\n" + values["サービス概要"],
    ]
    if values.get("現在取り組んでいるAI"):
        sections.append("現在取り組んでいるAI:\n" + values["現在取り組んでいるAI"])
    if values.get("AIに期待すること"):
        sections.append("お客様がAIで実現したいこと:\n" + values["AIに期待すること"])
    return "\n\n".join(sections)


def load_source_text(input_file: Path) -> str:
    """UTF-8テキストまたは固定Excelテンプレートから入力を読み込む。"""
    if input_file.suffix.lower() == ".xlsx":
        return load_excel_service_input(input_file)
    if input_file.suffix.lower() == ".xls":
        raise ValueError(".xls は未対応です。固定テンプレートを .xlsx 形式で保存してください。")
    _validate_input_size(input_file, MAX_TEXT_INPUT_BYTES, "テキスト入力")
    return input_file.read_text(encoding="utf-8").strip()

def resolve_input_file(input_file: Path) -> Path:
    """入力ファイル名だけの指定を、assessment_inputsを優先して解決する。"""
    if input_file.is_absolute():
        return input_file
    input_candidate = INPUT_DIR / input_file
    if input_candidate.is_file():
        return input_candidate
    project_candidate = PROJECT_DIR / input_file
    if project_candidate.is_file():
        return project_candidate
    # 既存の運用（カレントフォルダからの相対パス）も維持する。
    return input_file
