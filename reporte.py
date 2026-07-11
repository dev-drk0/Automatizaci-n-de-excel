from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Any
import pandas as pd
from pypdf import PdfReader


def extract_data_from_pdf(pdf_path: Path) -> pd.DataFrame:
    """
    Lee un archivo PDF y extrae los datos de texto de forma estructurada.
    Modifica las expresiones regulares según el formato real de tus PDFs.
    """
    reader = PdfReader(pdf_path)
    extracted_records = []

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue
            
        # Ejemplo de procesamiento línea por línea
        # Supongamos que buscamos líneas con un producto y un monto (ej: "Producto A - $1,200")
        lines = text.split("\n")
        for line in lines:
            # Una expresión regular simple para buscar texto seguido de números/precios
            match = re.search(r"([\w\s]+?)\s*[\$-]?\s*([\d\.,]+)", line)
            if match:
                product = match.group(1).strip()
                amount_str = match.group(2).replace(",", "")
                try:
                    amount = float(amount_str)
                    extracted_records.append({"product": product, "amount": amount, "source_type": "PDF"})
                except ValueError:
                    continue

    if not extracted_records:
        # Fallback por si el PDF es plano o tiene otro formato
        return pd.DataFrame(columns=["product", "amount", "source_type"])
        
    return pd.DataFrame(extracted_records)


def process_input_file(input_path: str | Path, output_dir: str | Path | None = None) -> Dict[str, Any]:
    base_dir = Path(__file__).resolve().parent
    input_path = Path(input_path)
    if not input_path.is_absolute():
        input_path = base_dir / input_path

    output_dir = Path(output_dir or "output")
    if not output_dir.is_absolute():
        output_dir = base_dir / output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    suffix = input_path.suffix.lower()
    df_result = pd.DataFrame()

    # --- PROCESAMIENTO DEPENDIENDO DEL TIPO ---
    if suffix in [".xlsx", ".xls"]:
        df = pd.read_excel(input_path)
        # Intentar estandarizar columnas si vienen con nombres distintos
        df.columns = [c.lower().strip() for c in df.columns]
        
        if "amount" in df.columns and "product" in df.columns:
            df_result = df[["product", "amount"]].copy()
            df_result["source_type"] = "Excel"
            
    elif suffix == ".csv":
        df = pd.read_csv(input_path)
        df.columns = [c.lower().strip() for c in df.columns]
        if "amount" in df.columns and "product" in df.columns:
            df_result = df[["product", "amount"]].copy()
            df_result["source_type"] = "CSV"
            
    elif suffix == ".pdf":
        df_result = extract_data_from_pdf(input_path)
        
    else:
        raise ValueError(format(f"Formato de archivo no soportado: {suffix}"))

    if df_result.empty:
        raise ValueError("No se pudieron extraer datos válidos del archivo proporcionado.")

    # Limpieza de datos genérica
    df_result["amount"] = pd.to_numeric(df_result["amount"], errors="coerce")
    df_result = df_result.dropna(subset=["amount"])

    # Acomodar/Consolidar los datos: Agrupación de control
    consolidated_path = output_dir / "datos_procesados.xlsx"
    df_result.to_excel(consolidated_path, index=False, sheet_name="Datos Extraídos")

    return {
        "status": "success",
        "file_type_detected": suffix.upper().replace(".", ""),
        "rows_processed": len(df_result),
        "output_file": consolidated_path,
        "total_amount": float(df_result["amount"].sum())
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        res = process_input_file(sys.argv[1])
        print(f"Procesado con éxito: {res}")