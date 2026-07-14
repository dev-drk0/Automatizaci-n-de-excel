from __future__ import annotations

import threading
from pathlib import Path
import flet as ft
import tkinter as tk
from tkinter import filedialog

from reporte import process_multiple_files_in_memory


def main(page: ft.Page):
    # --- CONFIGURACIÓN DE PÁGINA ---
    page.title = "Extractor Automático de Facturas"
    page.vertical_alignment = ft.MainAxisAlignment.CENTER
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.padding = 30
    page.bgcolor = "#F5F7FB"
    page.theme = ft.Theme(color_scheme_seed="#4C78A8")

    # Ocultar la ventana raíz de Tkinter para que solo salgan los diálogos
    root = tk.Tk()
    root.withdraw()

    # --- ESTADOS Y VARIABLES ---
    selected_files_paths: list[str] = []

    # --- COMPONENTES DE INTERFAZ ---
    file_path = ft.Text(
        value="Ningún archivo seleccionado",
        color="#4C78A8",
        italic=True,
    )

    status = ft.Text(
        value="Listo para procesar lotes de facturas",
        color=ft.Colors.GREY_700,
    )

    progress_bar = ft.ProgressBar(value=0, visible=False, width=500, height=8, border_radius=10)
    progress_ring = ft.ProgressRing(visible=False, width=28, height=28)
    preview = ft.Column(spacing=8, horizontal_alignment=ft.CrossAxisAlignment.START)

    # --- MANEJADORES DE EVENTOS ---
    def pick_files_click(e):
        nonlocal selected_files_paths
        # Diálogo nativo usando Tkinter (Adiós a los bugs de FilePicker)
        files = filedialog.askopenfilenames(
            title="Selecciona facturas en PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if files:
            selected_files_paths = list(files)
            file_path.value = f"{len(files)} factura(s) seleccionada(s)."
            file_path.italic = False
            status.value = "Lote cargado. Listo para extraer."
        else:
            status.value = "Selección cancelada."
        page.update()

    def start_processing_click(e):
        if not selected_files_paths:
            status.value = "Por favor, selecciona primero archivos PDF."
            status.color = ft.Colors.RED_600
            page.update()
            return

        # Guardar archivo nativo usando Tkinter
        save_path = filedialog.asksaveasfilename(
            title="¿Dónde deseas guardar tu reporte de Excel?",
            initialfile="control_facturas.xlsx",
            filetypes=[("Excel", "*.xlsx")],
            defaultextension=".xlsx"
        )

        if not save_path:
            status.value = "Exportación cancelada por el usuario."
            page.update()
            return

        # Si el usuario eligió ruta, disparamos el hilo pesado
        progress_bar.visible = True
        progress_ring.visible = True
        progress_bar.value = 0.30
        status.value = "Extrayendo textos y montos..."
        status.color = ft.Colors.BLUE_600
        preview.controls.clear()
        page.update()

        def run_extraction():
            try:
                result = process_multiple_files_in_memory(selected_files_paths, save_path)
                
                progress_bar.value = 1.0
                status.value = "¡Archivo guardado exitosamente!"
                status.color = ft.Colors.GREEN_600
                preview.controls = [
                    ft.ListTile(
                        leading=ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_800, size=32),
                        title=ft.Text(f"Guardado en: {Path(save_path).name}", weight=ft.FontWeight.BOLD),
                        subtitle=ft.Text(f"Facturas procesadas: {result['rows_processed']}\nMonto Total Absoluto: ${result['grand_total']:,.2f}"),
                    )
                ]
            except Exception as exc:
                status.value = f"Error: {exc}"
                status.color = ft.Colors.RED_600
            finally:
                progress_bar.visible = False
                progress_ring.visible = False
                page.update()

        threading.Thread(target=run_extraction, daemon=True).start()

    # --- INTERFAZ GRÁFICA ---
    header = ft.Column(
        [
            ft.Icon(ft.Icons.RECEIPT_LONG, size=56, color=ft.Colors.BLUE_ACCENT),
            ft.Text("Extractor Automático de Facturas", size=30, weight=ft.FontWeight.BOLD),
            ft.Text("Procesa lotes de PDFs nativos y exporta un Excel personalizado al instante.", color=ft.Colors.GREY_700),
        ],
        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        spacing=10,
    )

    dashboard_card = ft.Container(
        content=ft.Column(
            [
                ft.Text("📂 Facturas en PDF", weight=ft.FontWeight.BOLD, size=16),
                ft.Container(
                    content=file_path,
                    bgcolor=ft.Colors.GREY_50,
                    padding=15,
                    border_radius=8,
                    border=ft.Border.all(1, ft.Colors.GREY_300),
                ),
                ft.FilledButton(
                    "Seleccionar PDFs (Soporta Múltiples)",
                    icon=ft.Icons.FILE_UPLOAD,
                    on_click=pick_files_click,
                ),
                
                ft.Divider(height=30, color=ft.Colors.GREY_200),

                ft.Text("⚙️ Procesamiento", weight=ft.FontWeight.BOLD, size=16),
                status,
                ft.Row([progress_ring, progress_bar], alignment=ft.MainAxisAlignment.CENTER, spacing=15),
                ft.FilledButton(
                    "Procesar y Exportar Excel",
                    icon=ft.Icons.BOLT,
                    on_click=start_processing_click,
                ),

                ft.Divider(height=30, color=ft.Colors.GREY_200),

                ft.Text("📊 Resumen de Operación", weight=ft.FontWeight.BOLD, size=16),
                preview,
            ],
            spacing=15,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
        ),
        bgcolor=ft.Colors.WHITE,
        padding=30,
        border_radius=15,
        width=600,
        shadow=ft.BoxShadow(
            blur_radius=15, 
            color=ft.Colors.with_opacity(0.1, ft.Colors.BLACK), 
            offset=ft.Offset(0, 4)
        ),
    )

    page.add(ft.Column([header, ft.Divider(height=10, color=ft.Colors.TRANSPARENT), dashboard_card], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=20))

if __name__ == "__main__":
    # Volvemos de forma segura al modo escritorio clásico
    ft.run(main)