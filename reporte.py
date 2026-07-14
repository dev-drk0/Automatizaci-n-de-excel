from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Any, List
import pandas as pd
from pypdf import PdfReader

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


def extract_invoice_data(pdf_path: Path) -> Dict[str, Any]:
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"

    # --- RESPALDO POR SI ES PDF ESCANEADO (IMAGEN O FOTO) ---
    if not full_text.strip():
        return {
            "Nombre del Archivo": pdf_path.name,
            "Identificador Visual": pdf_path.stem.title(),
            "Proveedor": "⚠️ Requiere OCR (PDF Escaneado/Imagen)",
            "Fecha de Emisión": "No leíble",
            "Categoría de Gasto": "No definido",
            "Conceptos": 0,
            "Monto Total": 0.0,
            "Estado": "Error de Lectura Nátiva"
        }

    # --- PROVEEDOR ---
    clean_filename = pdf_path.stem.replace("_", " ").replace("-", " ").title()
    lines = [line.strip() for line in full_text.split("\n") if line.strip()]
    vendor = "No detectado"
    if lines:
        for line in lines[:4]:
            if not any(kw in line.lower() for kw in ["invoice", "factura", "fecha", "date", "bill to", "remit to"]):
                vendor = line
                break
        if vendor == "No detectado":
            vendor = clean_filename

    # --- FECHA ---
    date_val = "No detectada"
    numeric_date = re.search(r"(\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b)|(\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b)", full_text)
    if numeric_date:
        date_val = numeric_date.group(0)
    else:
        text_date = re.search(r"(\b\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4}\b)|(\b[a-zA-Z]+\s+\d{1,2},?\s+\d{4}\b)", full_text)
        if text_date:
            date_val = text_date.group(0)

    # --- MONTOS ---
    all_amounts = re.findall(r"[\$-]?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2}))", full_text)
    valid_amounts = []
    for amt in all_amounts:
        try:
            val = float(amt.replace(",", ""))
            if 0.5 <= val < 999999.0:
                valid_amounts.append(val)
        except ValueError:
            continue

    unique_amounts = list(dict.fromkeys(valid_amounts))
    total_amount = 0.0
    items_count = 0

    if unique_amounts:
        total_amount = max(unique_amounts)
        concept_amounts = [a for a in unique_amounts if a < total_amount]
        items_count = len(concept_amounts) if concept_amounts else 1

    category = "Gastos Operativos"
    if "cloud" in full_text.lower() or "aws" in full_text.lower() or "software" in full_text.lower():
        category = "Tecnología / Software"
    elif "marketing" in full_text.lower() or "ads" in full_text.lower():
        category = "Publicidad"

    return {
        "Nombre del Archivo": pdf_path.name,
        "Identificador Visual": clean_filename,
        "Proveedor": vendor,
        "Fecha de Emisión": date_val,
        "Categoría de Gasto": category,
        "Conceptos": items_count,
        "Monto Total": total_amount,
        "Estado": "Revisado / Pendiente"
    }


def process_multiple_files_in_memory(file_paths: List[str], target_save_path: str) -> Dict[str, Any]:
    """
    Procesa las facturas y guarda el archivo Excel EXACTAMENTE donde el usuario indicó 
    a través de la ventana nativa de guardado.
    """
    records = []
    for p in file_paths:
        path_obj = Path(p)
        if path_obj.suffix.lower() == ".pdf":
            try:
                records.append(extract_invoice_data(path_obj))
            except Exception:
                continue

    if not records:
        raise ValueError("No se pudo extraer información de ningún PDF válido.")

    df_result = pd.DataFrame(records)
    final_output = Path(target_save_path)

    # Escritura estética con Openpyxl
    with pd.ExcelWriter(final_output, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name="Reporte Consolidado")
        
        workbook = writer.book
        worksheet = writer.sheets["Reporte Consolidado"]
        
        header_fill = PatternFill(start_color="4C78A8", end_color="4C78A8", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        data_font = Font(name="Calibri", size=11, color="333333")
        total_fill = PatternFill(start_color="E6EDF5", end_color="E6EDF5", fill_type="solid")
        total_font = Font(name="Calibri", size=11, bold=True, color="000000")
        
        thin_side = Side(border_style="thin", color="CCCCCC")
        data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        for cell in worksheet[1]:
            cell.fill = header_fill; cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        num_rows = len(df_result)
        for row in worksheet.iter_rows(min_row=2, max_row=num_rows + 1):
            for cell in row:
                cell.font = data_font; cell.border = data_border
                if cell.column == 7:
                    cell.number_format = "$#,##0.00"
                    cell.alignment = Alignment(horizontal="right")

        total_row_idx = num_rows + 2
        worksheet.cell(row=total_row_idx, column=1, value="TOTAL GENERAL").font = total_font
        worksheet.cell(row=total_row_idx, column=7, value=f"=SUM(G2:G{total_row_idx-1})").font = total_font
        worksheet.cell(row=total_row_idx, column=7).number_format = "$#,##0.00"
        
        for col_idx in range(1, 9):
            cell = worksheet.cell(row=total_row_idx, column=col_idx)
            cell.fill = total_fill
            if col_idx == 7:
                cell.alignment = Alignment(horizontal="right")

        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        # Gráfica de Barras
        chart = BarChart()
        chart.type = "col"; chart.style = 10
        chart.title = "Análisis de Costos por Proveedor"
        chart.y_axis.title = "Monto ($)"
        chart.x_axis.title = "Proveedor"
        
        data_ref = Reference(worksheet, min_col=7, min_row=1, max_row=total_row_idx-1)
        cats_ref = Reference(worksheet, min_col=3, min_row=2, max_row=total_row_idx-1)
        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)
        chart.legend = None
        worksheet.add_chart(chart, "J2")

    return {
        "status": "success",
        "rows_processed": len(df_result),
        "grand_total": float(df_result["Monto Total"].sum())
    }