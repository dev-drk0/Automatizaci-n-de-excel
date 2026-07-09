from pathlib import Path

import pandas as pd

from reporte import build_report


def test_build_report_creates_expected_outputs(tmp_path):
    input_path = Path(__file__).resolve().parents[1] / "sample_sales.csv"
    output_dir = tmp_path / "output"

    result = build_report(input_path, output_dir)

    assert result["excel_path"].exists()
    assert result["chart_path"].exists()
    assert result["pdf_path"].exists()
    assert result["email_path"].exists()
    assert result["summary"]["total_sales"] > 0


def test_build_report_accepts_excel_input(tmp_path):
    input_path = tmp_path / "ventas.xlsx"
    output_dir = tmp_path / "output"

    pd.DataFrame(
        [
            {"product": "Producto A", "amount": 1200},
            {"product": "Producto B", "amount": 800},
        ]
    ).to_excel(input_path, index=False)

    result = build_report(input_path, output_dir)

    assert result["excel_path"].exists()
    assert result["summary"]["transactions"] == 2
