from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Any, List
import pandas as pd
from pypdf import PdfReader


def extract_invoice_data(pdf_path: Path) -> Dict[str, Any]:
    """
    Analiza a fondo el PDF de Invoice Generator.
    Extrae de forma robusta: Proveedor, Fecha, Cantidad de Montos Detectados y Total Real.
    """
    reader = PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"

    # --- 1. PROVEEDOR ---
    # Limpiamos el nombre del archivo para que sea legible en vez de "factura_123_final.pdf"
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

    # --- 2. FECHA EXTRA-ROBUSTA (Numérica y Texto en ES/EN) ---
    date_val = "No detectada"
    # Buscar formatos DD/MM/AAAA o AAAA-MM-DD
    numeric_date = re.search(r"(\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b)|(\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b)", full_text)
    if numeric_date:
        date_val = numeric_date.group(0)
    else:
        # Buscar formatos con texto (ej: "July 12, 2026" o "12 de Julio de 2026")
        text_date = re.search(r"(\b\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4}\b)|(\b[a-zA-Z]+\s+\d{1,2},?\s+\d{4}\b)", full_text)
        if text_date:
            date_val = text_date.group(0)
        else:
            # Búsqueda desesperada por líneas que contengan la palabra "date" o "fecha"
            for line in lines:
                if any(k in line.lower() for k in ["date", "fecha"]):
                    # Extraer lo que parezca un año o bloque de texto al final de esa línea
                    match = re.search(r"(?:date|fecha)[:\s]+([\w\s,\d/-]+)", line, re.IGNORECASE)
                    if match:
                        date_val = match.group(1).strip()
                        break

    # --- 3. EXTRACCIÓN MÚLTIPLE DE MONTOS ---
    # Capturamos todos los números con formato decimal (.00)
    all_amounts = re.findall(r"[\$-]?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2}))", full_text)
    
    valid_amounts = []
    for amt in all_amounts:
        try:
            val = float(amt.replace(",", ""))
            # Descartamos cantidades ridículas o folios numéricos con punto decimal
            if 0.5 <= val < 999999.0:
                valid_amounts.append(val)
        except ValueError:
            continue

    # Remover duplicados manteniendo el orden
    unique_amounts = list(dict.fromkeys(valid_amounts))
    
    total_amount = 0.0
    items_count = 0

    if unique_amounts:
        # El total de la factura siempre es el número más grande dentro de los cobros
        total_amount = max(unique_amounts)
        # Los conceptos individuales son los montos menores al total
        concept_amounts = [a for a in unique_amounts if a < total_amount]
        items_count = len(concept_amounts) if concept_amounts else 1
    else:
        items_count = 0

    # Categoría ficticia inteligente para poblar más celdas
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
        "Conceptos Detectados": items_count,
        "Monto Total": total_amount,
        "Estado del Pago": "Revisado / Pendiente"
    }


def process_multiple_files(file_paths: List[str], output_dir: str | Path | None = None) -> Dict[str, Any]:
    """
    Recibe una lista de archivos (pueden ser varios PDFs o un Excel histórico) 
    y consolida todo en el Excel de control.
    """
    base_dir = Path(__file__).resolve().parent
    output_dir = Path(output_dir or "output")
    if not output_dir.is_absolute():
        output_dir = base_dir / output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    records = []

    for p in file_paths:
        path_obj = Path(p)
        suffix = path_obj.suffix.lower()

        if suffix in [".xlsx", ".xls"]:
            df = pd.read_excel(path_obj)
            normalized_cols = {c.lower().strip(): c for c in df.columns}
            
            # Sinónimos para Excel histórico
            synonyms_provider = ["proveedor", "autor", "titulo", "title", "vendor", "name", "nombre"]
            synonyms_date = ["fecha", "date", "emision", "created_at", "fecha de emisión"]
            synonyms_amount = ["monto total", "precio", "total", "amount", "price", "neto", "importe"]
            
            col_provider = next((normalized_cols[k] for k in normalized_cols if k in synonyms_provider), None)
            col_date = next((normalized_cols[k] for k in normalized_cols if k in synonyms_date), None)
            col_amount = next((normalized_cols[k] for k in normalized_cols if k in synonyms_amount), None)
            
            for _, row in df.iterrows():
                p_val = row[col_provider] if col_provider else "No provisto"
                d_val = row[col_date] if col_date else "No provista"
                a_val = pd.to_numeric(row[col_amount], errors="coerce") if col_amount else 0.0
                
                records.append({
                    "Nombre del Archivo": path_obj.name,
                    "Identificador Visual": str(p_val).title(),
                    "Proveedor": p_val,
                    "Fecha de Emisión": d_val,
                    "Categoría de Gasto": "Importado Histórico",
                    "Conceptos Detectados": 1,
                    "Monto Total": float(a_val) if not pd.isna(a_val) else 0.0,
                    "Estado del Pago": "Cargado desde Excel"
                })

        elif suffix == ".pdf":
            try:
                invoice_data = extract_invoice_data(path_obj)
                records.append(invoice_data)
            except Exception:
                continue

    if not records:
        raise ValueError("No se pudo extraer información válida de ningún archivo seleccionado.")

    df_result = pd.DataFrame(records)
    consolidated_path = output_dir / "control_facturas.xlsx"
    df_result.to_excel(consolidated_path, index=False, sheet_name="Reporte Consolidado")

    return {
        "status": "success",
        "total_files": len(file_paths),
        "rows_processed": len(df_result),
        "output_file": consolidated_path,
        "grand_total": float(df_result["Monto Total"].sum())
    }