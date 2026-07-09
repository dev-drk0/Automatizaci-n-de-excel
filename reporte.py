from __future__ import annotations

from pathlib import Path
from typing import Dict, Any

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet


def build_report(input_path: str | Path, output_dir: str | Path | None = None) -> Dict[str, Any]:
    base_dir = Path(__file__).resolve().parent
    input_path = Path(input_path)
    if not input_path.is_absolute():
        input_path = base_dir / input_path

    output_dir = Path(output_dir or "output")
    if not output_dir.is_absolute():
        output_dir = base_dir / output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    if input_path.suffix.lower() == ".xlsx":
        df = pd.read_excel(input_path)
    else:
        df = pd.read_csv(input_path)

    if "amount" not in df.columns:
        raise ValueError("El archivo debe incluir una columna 'amount'.")

    df["amount"] = pd.to_numeric(df["amount"], errors="coerce")
    df = df.dropna(subset=["amount"])

    summary = {
        "total_sales": float(df["amount"].sum()),
        "average_sale": float(df["amount"].mean()),
        "transactions": int(len(df)),
    }

    excel_path = output_dir / "reporte_ventas.xlsx"
    chart_path = output_dir / "ventas_por_producto.png"
    pdf_path = output_dir / "reporte_ventas.pdf"
    email_path = output_dir / "correo.txt"

    sales_by_product = df.groupby("product", as_index=False)["amount"].sum().sort_values("amount", ascending=False)
    sales_by_product.to_excel(excel_path, index=False, sheet_name="Ventas")

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(sales_by_product["product"], sales_by_product["amount"], color="#4C78A8")
    ax.set_title("Ventas por producto")
    ax.set_ylabel("Monto")
    ax.tick_params(axis="x", rotation=20)
    fig.tight_layout()
    fig.savefig(chart_path, dpi=150)
    plt.close(fig)

    doc = SimpleDocTemplate(str(pdf_path), pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph("Reporte de ventas", styles["Title"]))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Ventas totales: ${summary['total_sales']:,.2f}", styles["Heading2"]))
    story.append(Paragraph(f"Promedio por venta: ${summary['average_sale']:,.2f}", styles["BodyText"]))
    story.append(Paragraph(f"Transacciones: {summary['transactions']}", styles["BodyText"]))
    story.append(Spacer(1, 12))

    data = [["Producto", "Monto"]]
    for _, row in sales_by_product.iterrows():
        data.append([row["product"], f"${row['amount']:,.2f}"])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4C78A8")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(table)
    story.append(Spacer(1, 12))
    story.append(Paragraph("Adjunto gráfico resumen del desempeño.", styles["BodyText"]))
    doc.build(story)

    email_body = (
        "Asunto: Reporte de ventas listo\n\n"
        "Hola,\n\n"
        "Adjunto el reporte generado automáticamente con Excel, gráfico y PDF.\n"
        f"Ventas totales: ${summary['total_sales']:,.2f}\n"
        f"Transacciones: {summary['transactions']}\n\n"
        "Saludos,"
    )
    email_path.write_text(email_body, encoding="utf-8")

    return {
        "excel_path": excel_path,
        "chart_path": chart_path,
        "pdf_path": pdf_path,
        "email_path": email_path,
        "summary": summary,
    }


if __name__ == "__main__":
    import sys

    input_file = sys.argv[1] if len(sys.argv) > 1 else "sample_sales.csv"
    build_report(input_file)
    print("Reporte generado correctamente.")
