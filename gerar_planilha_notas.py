#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

MAIN_DATE_FORMAT = "%Y-%m-%d"
VENCIMENTO_RE = re.compile(
    r"(?i)\b(?:data\s+de\s+vencimento|vencimento|due\s+date)\s*:\s*(\d{2}/\d{2}/\d{4})"
)

PORTUGUESE_MONTHS = {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4,
    "maio": 5, "junho": 6, "julho": 7, "agosto": 8,
    "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
}


@dataclass(frozen=True)
class NotaRow:
    numero_nf: str
    data_emissao: str
    data_vencimento: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera uma planilha Excel agrupada por mes de emissao com as colunas n_nota, data_emissao e data_vencimento."
    )
    parser.add_argument(
        "--input",
        default="nfs_202607201526.csv",
        help="CSV principal contendo as colunas numero_nf e data_emissao.",
    )
    parser.add_argument(
        "--monthly-dir",
        default=".",
        help="Diretorio base onde estao as pastas dos anos (ex: 2025, 2026) contendo os CSVs mensais.",
    )
    parser.add_argument(
        "--output",
        default="notas_agrupadas_por_mes.xlsx",
        help="Caminho do arquivo .xlsx de saida.",
    )
    return parser.parse_args()


def extract_due_date(text: str) -> str:
    if not text:
        return ""
    text_clean = text.replace("\xa0", " ").strip()
    
    # 1. Tenta padrão numérico DD/MM/YYYY
    match_numeric = VENCIMENTO_RE.search(text_clean)
    if match_numeric:
        try:
            return datetime.strptime(match_numeric.group(1), "%d/%m/%Y").strftime("%Y-%m-%d")
        except ValueError:
            pass

    # 2. Tenta padrão com mês por extenso (ex: "08 de Fevereiro 2025" ou "13 de Abril de 2025")
    pattern_text = re.compile(
        r"(?i)\b(?:data\s+de\s+vencimento|vencimento|due\s+date)\s*:\s*(\d{1,2})\s+(?:de\s+)?([a-zç]+)\s+(?:de\s+)?(\d{4})",
        re.IGNORECASE
    )
    match_text = pattern_text.search(text_clean)
    if match_text:
        day = int(match_text.group(1))
        month_name = match_text.group(2).lower()
        year = int(match_text.group(3))
        month_num = PORTUGUESE_MONTHS.get(month_name)
        if month_num:
            return f"{year:04d}-{month_num:02d}-{day:02d}"

    # 3. Tenta padrão "A Vista" / "A VISTA" / "a vista"
    pattern_a_vista = re.compile(r"(?i)\b(?:vencimento|due\s+date)\s*:\s*a\s+vista\b")
    if pattern_a_vista.search(text_clean):
        return "A Vista"

    return ""


def read_main_csv(path: Path) -> list[tuple[str, datetime]]:
    rows = []
    with path.open(newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            nf = (row.get("numero_nf") or "").strip()
            dt_raw = (row.get("data_emissao") or "").strip()
            if nf and dt_raw:
                rows.append((nf, datetime.strptime(dt_raw, MAIN_DATE_FORMAT)))
    return rows


def read_due_dates_by_years(monthly_dir: Path, years: Iterable[int], input_file: Path) -> dict[str, str]:
    due_dates: dict[str, str] = {}
    for year in years:
        year_dir = monthly_dir / str(year)
        if not year_dir.is_dir():
            continue
        
        # Filtra os arquivos CSV mensais do respectivo ano
        csv_files = sorted(
            path for path in year_dir.glob("*.csv")
            if path.resolve() != input_file.resolve()
        )
        
        for path in csv_files:
            with path.open(newline="", encoding="latin-1") as f:
                reader = csv.reader(f, delimiter=";")
                next(reader, None)  # Pula o cabeçalho
                for row in reader:
                    if len(row) >= 2:
                        nf = row[1].strip()
                        if nf and nf not in due_dates:
                            due_dates[nf] = extract_due_date(row[-1])
    return due_dates


def build_rows(main_rows: list[tuple[str, datetime]], due_dates: dict[str, str]) -> list[NotaRow]:
    rows = []
    for nf, dt in sorted(main_rows, key=lambda item: (item[1], item[0])):
        due = due_dates.get(nf, "")
        if due == "A Vista":
            due = dt.strftime("%Y-%m-%d")
        rows.append(
            NotaRow(
                numero_nf=nf,
                data_emissao=dt.strftime("%Y-%m-%d"),
                data_vencimento=due,
            )
        )
    return rows


def worksheet_xml(rows: list[NotaRow]) -> str:
    # 3 colunas fixas: A, B, C
    xml_rows = [
        '<row r="1">'
        '<c r="A1" t="inlineStr"><is><t>n_nota</t></is></c>'
        '<c r="B1" t="inlineStr"><is><t>data_emissao</t></is></c>'
        '<c r="C1" t="inlineStr"><is><t>data_vencimento</t></is></c>'
        '</row>'
    ]
    for r_idx, row in enumerate(rows, start=2):
        nf = escape(row.numero_nf)
        emissao = escape(row.data_emissao)
        vencimento = escape(row.data_vencimento)
        xml_rows.append(
            f'<row r="{r_idx}">'
            f'<c r="A{r_idx}" t="inlineStr"><is><t>{nf}</t></is></c>'
            f'<c r="B{r_idx}" t="inlineStr"><is><t>{emissao}</t></is></c>'
            f'<c r="C{r_idx}" t="inlineStr"><is><t>{vencimento}</t></is></c>'
            f'</row>'
        )

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>{"".join(xml_rows)}</sheetData>
</worksheet>"""


def workbook_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="notas" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""


def workbook_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""


def root_rels_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def content_types_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""


def write_xlsx(output_path: Path, rows: list[NotaRow]) -> None:
    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types_xml())
        archive.writestr("_rels/.rels", root_rels_xml())
        archive.writestr("xl/workbook.xml", workbook_xml())
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        archive.writestr("xl/worksheets/sheet1.xml", worksheet_xml(rows))


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).resolve()
    monthly_dir = Path(args.monthly_dir).resolve()
    output_path = Path(args.output).resolve()

    main_rows = read_main_csv(input_path)
    
    # Extrai os anos únicos a partir das datas de emissão
    years = {dt.year for _, dt in main_rows}
    
    # Busca e lê as datas de vencimento nas subpastas de cada ano
    due_dates = read_due_dates_by_years(monthly_dir, years, input_path)
    
    rows = build_rows(main_rows, due_dates)
    write_xlsx(output_path, rows)

    total = len(rows)
    unresolved = sum(1 for row in rows if not row.data_vencimento)
    print(f"Arquivo gerado: {output_path}")
    print(f"Notas processadas: {total}")
    print("Abas criadas: 1")
    print(f"Notas sem data de vencimento encontrada: {unresolved}")


if __name__ == "__main__":
    main()
