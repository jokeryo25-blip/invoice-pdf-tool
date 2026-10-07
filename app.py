import re
import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


def find_first_pdf(folder: Path) -> Path | None:
    pdf_files = [
        path for path in folder.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"
    ]
    if not pdf_files:
        return None
    return sorted(pdf_files)[0]


def extract_text_from_pdf(pdf_path: Path) -> str:
    extracted_text = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if page_text:
                extracted_text.append(page_text)

    return "\n\n".join(extracted_text)


def extract_invoice_date(pdf_text: str) -> str:
    match = re.search(r"請求日\s*:\s*(.+)", pdf_text)
    return match.group(1).strip() if match else ""


def extract_company_name(pdf_text: str) -> str:
    match = re.search(r"会社名\s*:\s*(.+)", pdf_text)
    return match.group(1).strip() if match else ""


def extract_invoice_number(pdf_text: str) -> str:
    match = re.search(r"請求番号\s*:\s*(.+)", pdf_text)
    return match.group(1).strip() if match else ""


def extract_amount(pdf_text: str) -> str:
    match = re.search(r"請求金額\s*:\s*(.+)", pdf_text)
    return match.group(1).strip() if match else ""


def append_run_log(project_dir: Path, processed_count: int, error_count: int) -> None:
    logs_dir = project_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    log_path = logs_dir / "app.log"
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(
            f"実行日時: {timestamp} | 処理件数: {processed_count} | "
            f"エラー件数: {error_count}\n"
        )


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    config_path = project_dir / "config.json"
    with config_path.open(encoding="utf-8") as config_file:
        config = json.load(config_file)

    invoices_dir = project_dir / config["入力フォルダ"]
    output_dir = project_dir / config["出力フォルダ"]

    if not invoices_dir.exists():
        print(f"PDF folder not found: {invoices_dir}")
        append_run_log(project_dir, 0, 0)
        return 1

    pdf_files = sorted(
        path for path in invoices_dir.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"
    )
    if not pdf_files:
        print(f"No PDF file found in: {invoices_dir}")
        append_run_log(project_dir, 0, 0)
        return 1

    invoice_list = []
    invoice_file_names = []
    error_list = []
    for pdf_path in pdf_files:
        try:
            text = extract_text_from_pdf(pdf_path)
            invoice = {
                "請求日": extract_invoice_date(text),
                "会社名": extract_company_name(text),
                "請求番号": extract_invoice_number(text),
                "請求金額": extract_amount(text),
            }
            missing_fields = [field for field, value in invoice.items() if not value]
            if missing_fields:
                raise ValueError(f"抽出できない項目: {', '.join(missing_fields)}")
            invoice_list.append(invoice)
            invoice_file_names.append(pdf_path.name)
        except Exception as error:
            print(f"処理をスキップしました: {pdf_path.name} ({error})")
            error_list.append(
                {
                    "ファイル名": pdf_path.name,
                    "エラー内容": str(error),
                }
            )

    append_run_log(project_dir, len(pdf_files), len(error_list))

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "請求書一覧.xlsx"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Invoice"
    invoice_headers = ["請求日", "会社名", "請求番号", "請求金額"]
    headers = ["ファイル名", *invoice_headers]
    worksheet.append(headers)
    for file_name, invoice in zip(invoice_file_names, invoice_list):
        worksheet.append([file_name, *[invoice[header] for header in invoice_headers]])
        amount_cell = worksheet.cell(row=worksheet.max_row, column=headers.index("請求金額") + 1)
        amount_value = re.sub(r"[^0-9.-]", "", invoice["請求金額"])
        amount_cell.value = float(amount_value)
        amount_cell.number_format = '"¥"#,##0.##'

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    for column_cells in worksheet.columns:
        max_length = max(
            sum(
                2 if unicodedata.east_asian_width(character) in {"F", "W"} else 1
                for character in str(cell.value)
            )
            for cell in column_cells
            if cell.value is not None
        )
        worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 2

    error_worksheet = workbook.create_sheet("Error")
    error_headers = ["ファイル名", "エラー内容"]
    error_worksheet.append(error_headers)
    for error in error_list:
        error_worksheet.append([error[header] for header in error_headers])

    summary_worksheet = workbook.create_sheet("Summary")
    summary_worksheet.append(["項目", "値"])
    invoice_count = worksheet.max_row - 1
    amount_column = headers.index("請求金額") + 1
    total_amount = sum(
        worksheet.cell(row=row, column=amount_column).value or 0
        for row in range(2, worksheet.max_row + 1)
    )
    average_amount = total_amount / invoice_count if invoice_count else 0
    summary_worksheet.append(["件数", invoice_count])
    summary_worksheet.append(["合計請求額", total_amount])
    summary_worksheet.append(["平均請求額", average_amount])

    workbook.save(output_path)

    print(invoice_list)
    print(f"Excel file created: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
