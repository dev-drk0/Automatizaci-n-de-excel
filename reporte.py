from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Any
import pandas as pd
from pypdf import PdfReader


def extract_invoice_data(pdf_path: Path) -> Dict[str, Any]:
    """
    Analiza el texto de una factura en PDF y extrae de forma inteligente:
    Fecha, Proveedor y Monto Total.
    """
    reader = PdfReader(pdf_path)
    full_text = ""
    
    # Extraer todo el texto del PDF unificado
    for page in reader.pages:
        text = page.extract_text()
        if text:
            full_text += text + "\n"
            
    # Valores por defecto por si no se encuentra algo
    vendor = "No detectado"
    date = "No detectada"
    total_amount = 0.0

    # 1. EXTRAER PROVEEDOR (Usualmente las primeras líneas o antes de "Factura" / "RFC")
    lines = [line.strip() for line in full_text.split("\n") if line.strip()]
    if lines:
        # Intentamos tomar la primera línea significativa como el nombre del Proveedor
        for line in lines[:5]:
            if not any(keyword in line.lower() for keyword in ["factura", "folio", "fecha", "receptor", "cliente", "rfc"]):
                vendor = line
                break

    # 2. EXTRAER FECHA (Formatos comunes: DD/MM/AAAA, AAAA-MM-DD, etc.)
    date_match = re.search(r"(\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b)|(\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b)", full_text)
    if date_match:
        date = date_match.group(0)

    # 3. EXTRAER MONTO TOTAL 
    # Buscamos palabras clave de facturación y capturamos el número decimal más cercano
    # Evitamos capturar IDs largos buscando números que tengan estructura de moneda (ej: 1,250.00 o 450.00)
    keywords_total = [r"total", r"importe total", r"total a pagar", r"neto", r"monto total"]
    
    possible_amounts = []
    
    for line in full_text.split("\n"):
        line_lower = line.lower()
        if any(kw in line_lower for kw in keywords_total):
            # Buscar montos con formato decimal (ej: 1,500.50 o 300.00)
            amounts = re.findall(r"[\$-]?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2}))", line)
            for amt in amounts:
                try:
                    val = float(amt.replace(",", ""))
                    # Filtro inteligente: Un ID de pago o calle suele no tener decimales o ser gigantesco.
                    if 0.1 <= val < 99999999.0: 
                        possible_amounts.append(val)
                except ValueError:
                    continue

    if possible_amounts:
        # Normalmente el total real es el valor más alto encontrado cerca de la palabra "Total"
        total_amount = max(possible_amounts)
    else:
        # Búsqueda de emergencia en todo el texto si no se encontró con palabras clave
        all_decimals = re.findall(r"[\$-]?\s*([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2}))", full_text)
        valid_decimals = []
        for amt in all_decimals:
            try:
                val = float(amt.replace(",", ""))
                valid_decimals.append(val)
            except ValueError:
                pass
        if valid_decimals:
            total_amount = max(valid_decimals) # El monto más grande suele ser el Total

    return {
        "Archivo": pdf_path.name,
        "Proveedor": vendor,
        "Fecha": date,
        "Monto Total": total_amount
    }


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
    
    # Lista donde acumularemos las filas del reporte
    records = []

    # --- PROCESAMIENTO EXCLUSIVO EXCEL O PDF ---
    if suffix in [".xlsx", ".xls"]:
        # Si suben un Excel, asumimos que es una lista previa que quieren validar/acomodar
        df = pd.read_excel(input_path)
        df.columns = [c.strip().title() for c in df.columns]
        
        # Estandarizar columnas requeridas
        required = ["Proveedor", "Fecha", "Monto Total"]
        for col in required:
            if col not in df.columns:
                df[col] = "No provisto" if col != "Monto Total" else 0.0
                
        # Limpiar montos
        df["Monto Total"] = pd.to_numeric(df["Monto Total"], errors="coerce").fillna(0.0)
        df["Archivo"] = input_path.name
        records = df[["Archivo", "Proveedor", "Fecha", "Monto Total"]].to_dict(orient="records")
            
    elif suffix == ".pdf":
        # Extraer los datos de la factura con las reglas inteligentes
        invoice_data = extract_invoice_data(input_path)
        records.append(invoice_data)
        
    else:
        raise ValueError(f"Formato no soportado ({suffix}). Por favor sube solo archivos PDF o Excel.")

    if not records:
        raise ValueError("No se pudieron extraer datos válidos del archivo.")

    # Crear DataFrame final consolidado
    df_result = pd.DataFrame(records)

    # Exportar al Excel Unificado de Salida
    consolidated_path = output_dir / "control_facturas.xlsx"
    
    # Si el archivo final ya existe, opcionalmente podríamos hacer un append, 
    # pero para cumplir el requerimiento de "un solo clic", guardamos limpio el resultado actual:
    df_result.to_excel(consolidated_path, index=False, sheet_name="Facturas Procesadas")

    return {
        "status": "success",
        "file_type_detected": suffix.upper().replace(".", ""),
        "rows_processed": len(df_result),
        "output_file": consolidated_path,
        "total_amount": float(df_result["Monto Total"].sum())
    }