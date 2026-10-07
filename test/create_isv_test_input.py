#!/usr/bin/env python3
"""ISV AI Use Case Assessment入力テンプレートへ、検証用の必須回答を投入する。

会社名と製品名だけを指定すると、公開Web検索の候補を補足情報として使いつつ、
AIアセスメントの動作確認に必要な必須設問を埋めたコピーを作成する。公開情報の
取得に失敗しても、明示的に「テスト値」とした汎用回答で出力できる。
"""

import argparse
import json
import re
import shutil
import sys
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse
from xml.etree import ElementTree

try:
    import httpx
    from bs4 import BeautifulSoup
except ImportError as error:  # pragma: no cover - runtime guidance for a missing venv
    raise SystemExit("必要な依存関係がありません。.venv/bin/python を使用してください。") from error


PROJECT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_TEMPLATE = PROJECT_DIR / "assessment_inputs" / "ISV_AI_Use_Case_Assessment_Input_Template.xlsx"
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "output" / "test_inputs"
SHEET_NAMES = ("入力ファイル", "入力フォーム")


def compact(value: str, limit: int = 360) -> str:
    """セルへ入れやすい長さに整形する。"""
    return re.sub(r"\s+", " ", value).strip()[:limit]


def search_public_web(company_name: str, product_name: str) -> list[dict[str, str]]:
    """DuckDuckGoの公開検索結果を最大3件だけ返す。失敗時は空配列。"""
    query = f"{company_name} {product_name} 製品 サービス"
    headers = {"User-Agent": "Mozilla/5.0 (compatible; AIAssessTestInput/1.0)"}
    try:
        with httpx.Client(timeout=12, follow_redirects=True, headers=headers) as client:
            response = client.get("https://html.duckduckgo.com/html/", params={"q": query})
            response.raise_for_status()
    except httpx.HTTPError:
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    results: list[dict[str, str]] = []
    for node in soup.select(".result"):
        link = node.select_one("a.result__a")
        if not link or not link.get("href"):
            continue
        parsed = urlparse(str(link["href"]))
        url = unquote(parse_qs(parsed.query).get("uddg", [str(link["href"])])[0])
        if not url.startswith("https://"):
            continue
        snippet = node.select_one(".result__snippet")
        results.append({
            "title": compact(link.get_text(" ", strip=True), 180),
            "url": url,
            "snippet": compact(snippet.get_text(" ", strip=True) if snippet else "", 360),
        })
        if len(results) == 3:
            break
    return results


def build_required_answers(company_name: str, product_name: str,
                           sources: list[dict[str, str]]) -> dict[str, str]:
    """公開情報の要約を補助に、必須設問へ安全な検証用回答を割り当てる。"""
    source_summary = " ".join(item.get("snippet", "") for item in sources if item.get("snippet"))
    overview = (
        f"{company_name}が提供する{product_name}に関するAIユースケースを検証するためのテスト入力。"
        + (f" 公開情報の要約: {source_summary}" if source_summary else " 公開情報は実行時に取得できなかったため、以下はテスト用の仮値。")
    )
    return {
        "F001": "確認しました",
        "F002": "確認しました",
        "F101": company_name,
        "F103": compact(overview),
        "F104": f"{product_name}を利用する業務部門・管理者・現場担当者（テスト値）。",
        "F107": "テスト入力者",
        "F108": "test@example.com",
        "F201": "既存SaaS/クラウドサービス",
        "F211": product_name,
        "F212": "本番運用中（テスト値）",
        "F213": "高",
        "F214": f"{product_name}の既存データを用い、利用者の検索・判断・確認業務をAIで支援できるか検討する。",
        "F301": product_name,
        "F304": f"{product_name}の代表的な業務機能・画面・ワークフロー（テスト値。公開情報とヒアリングで要確認）。",
        "F305": "顧客・案件・取引・操作履歴・マスタ・帳票・文書・問い合わせ履歴（テスト値）。",
        "F501": "製品仕様、業務マスタ、取引・操作履歴、帳票、マニュアル、FAQ（テスト値）。",
        "F601": "新サービス開発・収益増",
        "F602": "自然言語検索",
        "F711": "業務ナレッジ検索・判断支援",
        "F712": "利用者支援・業務効率化",
        "F713": "高",
        "F714": f"{product_name}のマニュアル、FAQ、業務データを横断検索し、利用者へ根拠付きの回答・次の確認事項を提示する。",
        "F715": "マニュアル、FAQ、操作ログ、業務マスタ、取引・案件履歴（テスト値）。",
        "F1010": "共有可",
    }


def fill_template(template_path: Path, output_path: Path, answers: dict[str, str]) -> int:
    """テンプレートをバイナリコピーし、入力シートのE列だけを直接更新する。

    Excelライブラリでブック全体を再保存せず、対象worksheet XMLだけを置換する。
    そのため、他シート・スタイル・図形・入力規則などのパッケージ部品は元ファイルの
    バイト列のまま保持される。
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    worksheet_path, rows = read_input_rows(output_path)
    question_rows = {row.get("B", "").strip(): row_number for row_number, row in rows.items()}
    missing = sorted(set(answers) - set(question_rows))
    if missing:
        raise ValueError("テンプレートに設問IDがありません: " + ", ".join(missing))
    write_answers_to_worksheet_xml(output_path, worksheet_path, {
        f"E{question_rows[question_id]}": answer for question_id, answer in answers.items()
    })
    return len(answers)


def read_input_rows(workbook_path: Path) -> tuple[str, dict[int, dict[str, str]]]:
    """入力シートのセル値と、xlsx内のworksheetパスを取得する。"""
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with zipfile.ZipFile(workbook_path) as archive:
        workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        relationships = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationship_map = {
            item.get("Id"): item.get("Target")
            for item in relationships.findall(f"{{{package_rel_ns}}}Relationship")
        }
        sheet = next((item for item in workbook.findall(f".//{{{main_ns}}}sheet") if item.get("name") in SHEET_NAMES), None)
        if sheet is None:
            raise ValueError("シート「入力ファイル」または「入力フォーム」が見つかりません。")
        target = relationship_map.get(sheet.get(f"{{{office_rel_ns}}}id"))
        if not target:
            raise ValueError("入力シートの関連情報を読み取れません。")
        worksheet_path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        worksheet = ElementTree.fromstring(archive.read(worksheet_path))
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = ["".join(node.text or "" for node in item.iter(f"{{{main_ns}}}t"))
                              for item in shared_root.findall(f"{{{main_ns}}}si")]

    rows: dict[int, dict[str, str]] = {}
    for cell in worksheet.findall(f".//{{{main_ns}}}c"):
        reference = cell.get("r", "")
        match = re.fullmatch(r"([A-Z]+)(\d+)", reference)
        if not match:
            continue
        cell_type = cell.get("t")
        if cell_type == "inlineStr":
            value = "".join(node.text or "" for node in cell.iter(f"{{{main_ns}}}t"))
        else:
            value_node = cell.find(f"{{{main_ns}}}v")
            value = value_node.text if value_node is not None and value_node.text else ""
            if cell_type == "s" and value:
                value = shared_strings[int(value)]
        rows.setdefault(int(match.group(2)), {})[match.group(1)] = value.strip()
    return worksheet_path, rows


def write_answers_to_worksheet_xml(workbook_path: Path, worksheet_path: str,
                                   answers_by_cell: dict[str, str]) -> None:
    """worksheet XML内の既存Eセルの値だけをinline stringとして差し替える。"""
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ElementTree.register_namespace("", main_ns)
    with zipfile.ZipFile(workbook_path) as archive:
        contents = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    root = ElementTree.fromstring(contents[worksheet_path])
    cells = {cell.get("r"): cell for cell in root.findall(f".//{{{main_ns}}}c")}
    missing_cells = sorted(set(answers_by_cell) - set(cells))
    if missing_cells:
        raise ValueError("テンプレートに回答セルがありません: " + ", ".join(missing_cells))
    for reference, answer in answers_by_cell.items():
        cell = cells[reference]
        cell.set("t", "inlineStr")
        for child in list(cell):
            cell.remove(child)
        inline = ElementTree.SubElement(cell, f"{{{main_ns}}}is")
        text = ElementTree.SubElement(inline, f"{{{main_ns}}}t")
        text.text = answer
    contents[worksheet_path] = ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)
    temporary_path = workbook_path.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(temporary_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)
    temporary_path.replace(workbook_path)


def main() -> int:
    parser = argparse.ArgumentParser(description="ISV入力テンプレートの必須設問へ検証用データを投入します。")
    parser.add_argument("--company", required=True, help="会社名。例: 株式会社エクス")
    parser.add_argument("--product", required=True, help="対象製品・サービス名。例: Factory-ONE 電脳工場")
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE, help="元となる空テンプレート")
    parser.add_argument("--output", type=Path, help="出力Excel。省略時はoutput/test_inputs配下")
    parser.add_argument("--offline", action="store_true", help="公開Webを検索せず、すべてテスト用の仮値で作成する")
    args = parser.parse_args()
    if not args.template.is_file():
        parser.error(f"テンプレートが見つかりません: {args.template}")
    safe_name = re.sub(r"[^0-9A-Za-zぁ-んァ-ン一-龯_-]+", "_", unicodedata.normalize("NFKC", args.product)).strip("_") or "test_input"
    output_path = args.output or DEFAULT_OUTPUT_DIR / f"ISV_AI_Use_Case_Assessment_Input_{safe_name}_test.xlsx"
    sources = [] if args.offline else search_public_web(args.company, args.product)
    answers = build_required_answers(args.company.strip(), args.product.strip(), sources)
    written = fill_template(args.template, output_path, answers)
    log_path = output_path.with_suffix(".sources.json")
    log_path.write_text(json.dumps({
        "generated_at": datetime.now().astimezone().isoformat(),
        "company": args.company,
        "product": args.product,
        "web_search_used": not args.offline,
        "sources": sources,
        "note": "テスト用の仮値を含みます。顧客提出用には使わないでください。",
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"テスト入力Excelを出力しました: {output_path}（必須回答 {written} 件）")
    print(f"公開情報ログ: {log_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError) as error:
        print(f"エラー: {error}", file=sys.stderr)
        raise SystemExit(1)
