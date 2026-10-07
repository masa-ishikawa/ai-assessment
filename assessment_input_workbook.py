"""Read and populate the fixed AI Assessment input workbook safely.

The workbook is an OOXML ZIP package.  This module intentionally uses only the
Python standard library and rewrites only the selected input worksheet part;
all other package parts are copied without changing their contents.  This is
important for retaining validations, styles, comments, drawings, and any other
Excel features that a general-purpose workbook round trip might normalize.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
import csv
import io
import os
from pathlib import Path
import posixpath
import re
import shutil
import tempfile
from typing import Iterable, Mapping
from xml.etree import ElementTree
import zipfile


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"

INPUT_SHEET_NAMES = ("入力フォーム", "入力ファイル")
ANSWER_COLUMN = "E"

_MAIN = f"{{{MAIN_NS}}}"
_OFFICE_REL = f"{{{OFFICE_REL_NS}}}"
_PACKAGE_REL = f"{{{PACKAGE_REL_NS}}}"
_CELL_REFERENCE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_QUESTION_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_.-]*[0-9][A-Za-z0-9_.-]*$")
_RANGE_FORMULA_RE = re.compile(
    r"^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?"
    r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)"
    r"(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$"
)


@dataclass(frozen=True, slots=True)
class FormQuestion:
    """One question discovered from the input form."""

    id: str
    row: int
    required: bool
    question: str
    guide: str
    input_type: str
    choices: tuple[str, ...]
    current_answer: str
    multi_select_delimiter: str | None = None

    @property
    def answer_cell(self) -> str:
        return f"{ANSWER_COLUMN}{self.row}"

    def as_dict(self) -> dict[str, object]:
        result = asdict(self)
        result["choices"] = list(self.choices)
        result["answer_cell"] = self.answer_cell
        return result


@dataclass(frozen=True, slots=True)
class FormSchema:
    """Workbook metadata and all questions in row order."""

    workbook_path: Path
    sheet_name: str
    worksheet_path: str
    questions: tuple[FormQuestion, ...]

    def as_dict(self) -> dict[str, object]:
        return {
            "workbook_path": str(self.workbook_path),
            "sheet_name": self.sheet_name,
            "worksheet_path": self.worksheet_path,
            "questions": [question.as_dict() for question in self.questions],
        }


@dataclass(frozen=True, slots=True)
class AnswerOutcome:
    """Result for one requested question ID."""

    id: str
    row: int
    cell: str
    previous_answer: str
    requested_answer: str
    status: str

    def as_dict(self) -> dict[str, object]:
        return asdict(self)


@dataclass(frozen=True, slots=True)
class WriteSummary:
    """Structured result of copying a template and applying answers."""

    template_path: Path
    output_path: Path
    sheet_name: str
    worksheet_path: str
    overwrite: bool
    requested_count: int
    written: tuple[str, ...]
    preserved: tuple[str, ...]
    outcomes: tuple[AnswerOutcome, ...]

    @property
    def written_count(self) -> int:
        return len(self.written)

    @property
    def preserved_count(self) -> int:
        return len(self.preserved)

    def as_dict(self) -> dict[str, object]:
        return {
            "template_path": str(self.template_path),
            "output_path": str(self.output_path),
            "sheet_name": self.sheet_name,
            "worksheet_path": self.worksheet_path,
            "overwrite": self.overwrite,
            "requested_count": self.requested_count,
            "written_count": self.written_count,
            "preserved_count": self.preserved_count,
            "written": list(self.written),
            "preserved": list(self.preserved),
            "outcomes": [outcome.as_dict() for outcome in self.outcomes],
        }


def read_form_schema(workbook_path: str | os.PathLike[str]) -> FormSchema:
    """Return the input form schema, including current E-column answers.

    Questions are discovered from their ID, required marker, question text,
    current answer, guide, input type, and choice metadata in columns B-I.  For
    list validations, choices are resolved from the actual validation formula
    (including ranges on ``選択肢一覧``).  Column I is used as a safe fallback.
    """

    path = Path(workbook_path)
    if not path.is_file():
        raise FileNotFoundError(f"Workbook not found: {path}")

    try:
        with zipfile.ZipFile(path) as archive:
            workbook_root = _read_xml(archive, "xl/workbook.xml")
            sheet_name, worksheet_path = _find_input_sheet(archive, workbook_root)
            shared_strings = _read_shared_strings(archive)
            worksheet_root = _read_xml(archive, worksheet_path)
            form_cells = _read_cell_values(worksheet_root, shared_strings)
            validations = _list_validations(worksheet_root)
            sheet_paths = _sheet_paths(archive, workbook_root)
            defined_names = _defined_names(workbook_root)

            sheet_cell_cache: dict[str, dict[str, str]] = {sheet_name: form_cells}

            def cells_for_sheet(name: str) -> dict[str, str] | None:
                if name in sheet_cell_cache:
                    return sheet_cell_cache[name]
                sheet_path = sheet_paths.get(name)
                if not sheet_path:
                    return None
                root = _read_xml(archive, sheet_path)
                values = _read_cell_values(root, shared_strings)
                sheet_cell_cache[name] = values
                return values

            questions: list[FormQuestion] = []
            seen_ids: set[str] = set()
            row_values = _rows_from_cells(form_cells)
            for row, values in sorted(row_values.items()):
                question_id = values.get("B", "").strip()
                question_text = values.get("D", "").strip()
                if not _is_question_row(question_id, values.get("C", ""), question_text):
                    continue
                if question_id in seen_ids:
                    raise ValueError(f"Duplicate question ID in input form: {question_id}")
                seen_ids.add(question_id)

                answer_cell = f"{ANSWER_COLUMN}{row}"
                validation = next(
                    (item for item in validations if _sqref_contains(item.get("sqref", ""), answer_cell)),
                    None,
                )
                metadata_choices = _split_semicolon_choices(values.get("I", ""))
                validation_choices = _validation_choices(
                    validation,
                    current_sheet=sheet_name,
                    defined_names=defined_names,
                    cells_for_sheet=cells_for_sheet,
                )
                choices = validation_choices or metadata_choices
                raw_input_type = values.get("H", "").strip()
                input_type = raw_input_type or ("Choice" if validation is not None else "Text")
                multi_select = _is_multi_select(
                    input_type=input_type,
                    guide=values.get("F", ""),
                    validation=validation,
                )
                questions.append(FormQuestion(
                    id=question_id,
                    row=row,
                    required=_is_required(values.get("C", "")),
                    question=question_text,
                    guide=values.get("F", "").strip(),
                    input_type=input_type,
                    choices=choices,
                    current_answer=values.get(ANSWER_COLUMN, ""),
                    multi_select_delimiter=";" if multi_select else None,
                ))
    except zipfile.BadZipFile as error:
        raise ValueError(f"Not a valid XLSX file: {path}") from error

    if not questions:
        raise ValueError(f"No questions found on sheet {sheet_name!r} in {path}")
    return FormSchema(
        workbook_path=path,
        sheet_name=sheet_name,
        worksheet_path=worksheet_path,
        questions=tuple(questions),
    )


def list_questions(workbook_path: str | os.PathLike[str]) -> tuple[FormQuestion, ...]:
    """Convenience wrapper returning just the questions."""

    return read_form_schema(workbook_path).questions


def write_answers(
    template_path: str | os.PathLike[str],
    output_path: str | os.PathLike[str],
    answers: Mapping[str, str],
    *,
    overwrite: bool = False,
) -> WriteSummary:
    """Copy ``template_path`` and populate E cells by question ID.

    Existing non-empty answers are retained unless ``overwrite`` is true.  The
    source and destination must be distinct paths.  Unknown question IDs and
    non-string values are rejected before the destination is created or
    replaced.
    """

    template = Path(template_path)
    output = Path(output_path)
    if not template.is_file():
        raise FileNotFoundError(f"Template workbook not found: {template}")
    if _paths_refer_to_same_file(template, output):
        raise ValueError("Output workbook must be different from the template workbook.")
    if not isinstance(answers, Mapping):
        raise TypeError("answers must be a mapping of question IDs to strings")

    normalized_answers: dict[str, str] = {}
    for question_id, answer in answers.items():
        if not isinstance(question_id, str) or not question_id.strip():
            raise TypeError("Every answer key must be a non-empty question ID string")
        if not isinstance(answer, str):
            raise TypeError(f"Answer for {question_id!r} must be a string")
        normalized_answers[question_id.strip()] = answer

    schema = read_form_schema(template)
    question_by_id = {question.id: question for question in schema.questions}
    missing_ids = sorted(set(normalized_answers) - set(question_by_id))
    if missing_ids:
        raise ValueError("Question IDs not found in template: " + ", ".join(missing_ids))

    outcomes: list[AnswerOutcome] = []
    answers_by_cell: dict[str, str] = {}
    written: list[str] = []
    preserved: list[str] = []
    for question_id, requested_answer in normalized_answers.items():
        question = question_by_id[question_id]
        if question.current_answer.strip() and not overwrite:
            status = "preserved"
            preserved.append(question_id)
        else:
            status = "written"
            written.append(question_id)
            answers_by_cell[question.answer_cell] = requested_answer
        outcomes.append(AnswerOutcome(
            id=question_id,
            row=question.row,
            cell=question.answer_cell,
            previous_answer=question.current_answer,
            requested_answer=requested_answer,
            status=status,
        ))

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    if answers_by_cell:
        _replace_worksheet_answers(output, schema.worksheet_path, answers_by_cell)

    return WriteSummary(
        template_path=template,
        output_path=output,
        sheet_name=schema.sheet_name,
        worksheet_path=schema.worksheet_path,
        overwrite=overwrite,
        requested_count=len(normalized_answers),
        written=tuple(written),
        preserved=tuple(preserved),
        outcomes=tuple(outcomes),
    )


def _read_xml(archive: zipfile.ZipFile, member: str) -> ElementTree.Element:
    try:
        return ElementTree.fromstring(archive.read(member))
    except KeyError as error:
        raise ValueError(f"Required XLSX part is missing: {member}") from error
    except ElementTree.ParseError as error:
        raise ValueError(f"Invalid XML in XLSX part: {member}") from error


def _read_shared_strings(archive: zipfile.ZipFile) -> tuple[str, ...]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return ()
    root = _read_xml(archive, "xl/sharedStrings.xml")
    return tuple(
        "".join(text.text or "" for text in item.iter(f"{_MAIN}t"))
        for item in root.findall(f"{_MAIN}si")
    )


def _sheet_paths(
    archive: zipfile.ZipFile,
    workbook_root: ElementTree.Element,
) -> dict[str, str]:
    relationships_root = _read_xml(archive, "xl/_rels/workbook.xml.rels")
    relationships = {
        relationship.get("Id", ""): relationship.get("Target", "")
        for relationship in relationships_root.findall(f"{_PACKAGE_REL}Relationship")
    }
    paths: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{_MAIN}sheet"):
        name = sheet.get("name", "")
        relationship_id = sheet.get(f"{_OFFICE_REL}id", "")
        target = relationships.get(relationship_id, "")
        if name and target:
            member = _resolve_package_target("xl/workbook.xml", target)
            if member in archive.namelist():
                paths[name] = member
    return paths


def _find_input_sheet(
    archive: zipfile.ZipFile,
    workbook_root: ElementTree.Element,
) -> tuple[str, str]:
    paths = _sheet_paths(archive, workbook_root)
    for sheet_name in INPUT_SHEET_NAMES:
        if sheet_name in paths:
            return sheet_name, paths[sheet_name]
    expected = " or ".join(repr(name) for name in INPUT_SHEET_NAMES)
    raise ValueError(f"Input worksheet {expected} was not found.")


def _resolve_package_target(source_member: str, target: str) -> str:
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_member), target))


def _read_cell_values(
    worksheet_root: ElementTree.Element,
    shared_strings: tuple[str, ...],
) -> dict[str, str]:
    values: dict[str, str] = {}
    for cell in worksheet_root.findall(f".//{_MAIN}c"):
        reference = cell.get("r", "").replace("$", "").upper()
        if not _CELL_REFERENCE_RE.fullmatch(reference):
            continue
        cell_type = cell.get("t", "")
        if cell_type == "inlineStr":
            value = "".join(text.text or "" for text in cell.iter(f"{_MAIN}t"))
        else:
            value_node = cell.find(f"{_MAIN}v")
            value = value_node.text or "" if value_node is not None else ""
            if cell_type == "s" and value:
                try:
                    value = shared_strings[int(value)]
                except (IndexError, ValueError) as error:
                    raise ValueError(f"Invalid shared-string index in cell {reference}: {value}") from error
            elif cell_type == "b" and value:
                value = "TRUE" if value == "1" else "FALSE"
        values[reference] = value
    return values


def _rows_from_cells(cells: Mapping[str, str]) -> dict[int, dict[str, str]]:
    rows: dict[int, dict[str, str]] = {}
    for reference, value in cells.items():
        match = _CELL_REFERENCE_RE.fullmatch(reference)
        if match:
            rows.setdefault(int(match.group(2)), {})[match.group(1).upper()] = value
    return rows


def _is_question_row(question_id: str, required_marker: str, question: str) -> bool:
    del required_marker  # The column header itself says "必須", so ID shape is authoritative.
    return bool(question_id and question and _QUESTION_ID_RE.fullmatch(question_id))


def _is_required(marker: str) -> bool:
    return marker.strip().casefold() in {"必須", "required", "yes", "true", "1"}


def _list_validations(worksheet_root: ElementTree.Element) -> tuple[dict[str, str], ...]:
    validations: list[dict[str, str]] = []
    for validation in worksheet_root.findall(f".//{_MAIN}dataValidation"):
        if validation.get("type", "") != "list":
            continue
        item = dict(validation.attrib)
        item["formula1"] = (validation.findtext(f"{_MAIN}formula1") or "").strip()
        validations.append(item)
    return tuple(validations)


def _sqref_contains(square_reference: str, target: str) -> bool:
    target_match = _CELL_REFERENCE_RE.fullmatch(target.replace("$", "").upper())
    if not target_match:
        return False
    target_column = _column_number(target_match.group(1))
    target_row = int(target_match.group(2))
    for item in square_reference.split():
        endpoints = item.split(":", 1)
        start_match = _CELL_REFERENCE_RE.fullmatch(endpoints[0].replace("$", ""))
        end_match = _CELL_REFERENCE_RE.fullmatch(endpoints[-1].replace("$", ""))
        if not start_match or not end_match:
            continue
        start_column = _column_number(start_match.group(1))
        end_column = _column_number(end_match.group(1))
        start_row = int(start_match.group(2))
        end_row = int(end_match.group(2))
        if (
            min(start_column, end_column) <= target_column <= max(start_column, end_column)
            and min(start_row, end_row) <= target_row <= max(start_row, end_row)
        ):
            return True
    return False


def _defined_names(workbook_root: ElementTree.Element) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in workbook_root.findall(f".//{_MAIN}definedName"):
        name = item.get("name", "")
        formula = (item.text or "").strip()
        if name and formula and name not in result:
            result[name] = formula
    return result


def _validation_choices(
    validation: Mapping[str, str] | None,
    *,
    current_sheet: str,
    defined_names: Mapping[str, str],
    cells_for_sheet,
) -> tuple[str, ...]:
    if not validation:
        return ()
    formula = validation.get("formula1", "").strip()
    if formula.startswith("="):
        formula = formula[1:].strip()
    if not formula:
        return ()
    if len(formula) >= 2 and formula[0] == formula[-1] == '"':
        return _split_literal_choices(formula[1:-1])
    formula = defined_names.get(formula, formula)
    if formula.startswith("="):
        formula = formula[1:].strip()
    parsed_range = _parse_range_formula(formula, current_sheet)
    if parsed_range is None:
        return ()
    sheet_name, start_column, start_row, end_column, end_row = parsed_range
    cells = cells_for_sheet(sheet_name)
    if cells is None:
        return ()
    choices: list[str] = []
    for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
        for column in range(min(start_column, end_column), max(start_column, end_column) + 1):
            value = cells.get(f"{_column_letters(column)}{row}", "").strip()
            if value and value not in choices:
                choices.append(value)
    return tuple(choices)


def _parse_range_formula(
    formula: str,
    current_sheet: str,
) -> tuple[str, int, int, int, int] | None:
    match = _RANGE_FORMULA_RE.fullmatch(formula.strip())
    if not match:
        return None
    quoted_sheet, unquoted_sheet, start_column, start_row, end_column, end_row = match.groups()
    if quoted_sheet is not None:
        sheet_name = quoted_sheet.replace("''", "'")
    else:
        sheet_name = (unquoted_sheet or current_sheet).strip()
    end_column = end_column or start_column
    end_row = end_row or start_row
    return (
        sheet_name,
        _column_number(start_column),
        int(start_row),
        _column_number(end_column),
        int(end_row),
    )


def _split_literal_choices(value: str) -> tuple[str, ...]:
    delimiter = ";" if ";" in value and "," not in value else ","
    reader = csv.reader(io.StringIO(value), delimiter=delimiter)
    return _unique_nonempty(next(reader, []))


def _split_semicolon_choices(value: str) -> tuple[str, ...]:
    return _unique_nonempty(re.split(r"[;；]", value))


def _unique_nonempty(values: Iterable[str]) -> tuple[str, ...]:
    result: list[str] = []
    for item in values:
        value = item.strip()
        if value and value not in result:
            result.append(value)
    return tuple(result)


def _is_multi_select(
    *,
    input_type: str,
    guide: str,
    validation: Mapping[str, str] | None,
) -> bool:
    if "複数" in input_type.casefold() or "multi" in input_type.casefold():
        return True
    if "複数選択" in guide:
        return True
    if not input_type.strip() and validation:
        validation_text = " ".join(
            validation.get(name, "") for name in ("promptTitle", "prompt", "errorTitle", "error")
        )
        return "区切" in validation_text and (";" in validation_text or "；" in validation_text)
    return False


def _column_number(letters: str) -> int:
    result = 0
    for character in letters.upper():
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def _column_letters(number: int) -> str:
    if number < 1:
        raise ValueError("Column number must be positive")
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _paths_refer_to_same_file(first: Path, second: Path) -> bool:
    try:
        if first.exists() and second.exists() and os.path.samefile(first, second):
            return True
    except OSError:
        pass
    return first.resolve(strict=False) == second.resolve(strict=False)


def _replace_worksheet_answers(
    workbook_path: Path,
    worksheet_path: str,
    answers_by_cell: Mapping[str, str],
) -> None:
    ElementTree.register_namespace("", MAIN_NS)
    with zipfile.ZipFile(workbook_path) as archive:
        try:
            worksheet_bytes = archive.read(worksheet_path)
        except KeyError as error:
            raise ValueError(f"Input worksheet part is missing: {worksheet_path}") from error
    try:
        worksheet_root = ElementTree.fromstring(worksheet_bytes)
    except ElementTree.ParseError as error:
        raise ValueError(f"Invalid XML in input worksheet part: {worksheet_path}") from error

    cells = {
        cell.get("r", "").replace("$", "").upper(): cell
        for cell in worksheet_root.findall(f".//{_MAIN}c")
        if cell.get("r")
    }
    for reference, answer in answers_by_cell.items():
        normalized_reference = reference.replace("$", "").upper()
        cell = cells.get(normalized_reference)
        if cell is None:
            cell = _create_cell(worksheet_root, normalized_reference)
            cells[normalized_reference] = cell
        _set_inline_string(cell, answer)

    replacement = ElementTree.tostring(
        worksheet_root,
        encoding="utf-8",
        xml_declaration=True,
        short_empty_elements=True,
    )
    _replace_zip_member(workbook_path, worksheet_path, replacement)


def _create_cell(worksheet_root: ElementTree.Element, reference: str) -> ElementTree.Element:
    match = _CELL_REFERENCE_RE.fullmatch(reference)
    if not match:
        raise ValueError(f"Invalid answer cell reference: {reference}")
    column_number = _column_number(match.group(1))
    row_number = int(match.group(2))
    sheet_data = worksheet_root.find(f"{_MAIN}sheetData")
    if sheet_data is None:
        raise ValueError("Input worksheet does not contain sheetData")
    row = next((item for item in sheet_data.findall(f"{_MAIN}row") if item.get("r") == str(row_number)), None)
    if row is None:
        row = ElementTree.Element(f"{_MAIN}row", {"r": str(row_number)})
        inserted = False
        for index, existing in enumerate(sheet_data.findall(f"{_MAIN}row")):
            if int(existing.get("r", "0")) > row_number:
                sheet_data.insert(index, row)
                inserted = True
                break
        if not inserted:
            sheet_data.append(row)
    cell = ElementTree.Element(f"{_MAIN}c", {"r": reference})
    inserted = False
    for index, existing in enumerate(row.findall(f"{_MAIN}c")):
        existing_match = _CELL_REFERENCE_RE.fullmatch(existing.get("r", ""))
        if existing_match and _column_number(existing_match.group(1)) > column_number:
            row.insert(index, cell)
            inserted = True
            break
    if not inserted:
        row.append(cell)
    return cell


def _set_inline_string(cell: ElementTree.Element, value: str) -> None:
    cell.set("t", "inlineStr")
    for child in list(cell):
        if child.tag in {f"{_MAIN}f", f"{_MAIN}v", f"{_MAIN}is"}:
            cell.remove(child)
    inline_string = ElementTree.SubElement(cell, f"{_MAIN}is")
    text = ElementTree.SubElement(inline_string, f"{_MAIN}t")
    if value != value.strip():
        text.set(f"{{{XML_NS}}}space", "preserve")
    text.text = value


def _replace_zip_member(workbook_path: Path, member: str, replacement: bytes) -> None:
    temporary_file = tempfile.NamedTemporaryFile(
        prefix=f".{workbook_path.name}.",
        suffix=".tmp",
        dir=workbook_path.parent,
        delete=False,
    )
    temporary_path = Path(temporary_file.name)
    temporary_file.close()
    try:
        with zipfile.ZipFile(workbook_path, "r") as source, zipfile.ZipFile(temporary_path, "w") as target:
            target.comment = source.comment
            found = False
            for info in source.infolist():
                data = replacement if info.filename == member else source.read(info)
                if info.filename == member:
                    found = True
                target.writestr(info, data)
            if not found:
                raise ValueError(f"XLSX part was not found while writing: {member}")
        os.replace(temporary_path, workbook_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


__all__ = [
    "AnswerOutcome",
    "FormQuestion",
    "FormSchema",
    "WriteSummary",
    "list_questions",
    "read_form_schema",
    "write_answers",
]
