from __future__ import annotations

import asyncio
from pathlib import Path
import flet as ft

# Importamos tu función de reporte.py
from reporte import process_multiple_files_in_memory


class ReportApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Extractor Automático de Facturas"
        self.page.vertical_alignment = ft.MainAxisAlignment.CENTER
        self.page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.page.padding = 30
        
        # Color de fondo moderno
        self.page.bgcolor = "#F5F7FB"
        self.page.theme = ft.Theme(color_scheme_seed="#4C78A8")

        # --- REGISTRAR LOS PICKERS COMO SERVICIOS (Solución al Timeout) ---
        self.file_picker_open = ft.FilePicker()
        self.file_picker_save = ft.FilePicker()
        self.page.services.append(self.file_picker_open)
        self.page.services.append(self.file_picker_save)

        # Asignar los eventos
        self.file_picker_open.on_result = self.on_files_selected
        self.file_picker_save.on_result = self.on_save_location_selected

        # Componentes de UI
        self.file_path = ft.Text(
            value="Ningún archivo seleccionado",
            color="#4C78A8",
            italic=True,
        )

        self.status = ft.Text(
            value="Listo para procesar lotes de facturas",
            color=ft.Colors.GREY_700,
        )

        self.progress_bar = ft.ProgressBar(
            value=0,
            visible=False,
            width=500,
            height=8,
            border_radius=10,
        )

        self.progress_ring = ft.ProgressRing(
            visible=False,
            width=28,
            height=28,
        )

        self.preview = ft.Column(
            spacing=8,
            horizontal_alignment=ft.CrossAxisAlignment.START,
        )

        self.selected_files_paths: list[str] = []

        # --- CONSTRUCCIÓN DE LA INTERFAZ ---
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
                        content=self.file_path,
                        bgcolor=ft.Colors.GREY_50,
                        padding=15,
                        border_radius=8,
                        border=ft.Border.all(1, ft.Colors.GREY_300),
                    ),
                    ft.FilledButton(
                        "Seleccionar PDFs (Soporta Múltiples)",
                        icon=ft.Icons.FILE_UPLOAD,
                        on_click=self.pick_files_click,
                    ),
                    
                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    ft.Text("⚙️ Procesamiento", weight=ft.FontWeight.BOLD, size=16),
                    self.status,
                    ft.Row([self.progress_ring, self.progress_bar], alignment=ft.MainAxisAlignment.CENTER, spacing=15),
                    ft.FilledButton(
                        "Procesar y Exportar Excel",
                        icon=ft.Icons.BOLT,
                        on_click=self.start_processing_click,
                    ),

                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    ft.Text("📊 Resumen de Operación", weight=ft.FontWeight.BOLD, size=16),
                    self.preview,
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

        self.page.add(
            ft.Column(
                [header, ft.Divider(height=10, color=ft.Colors.TRANSPARENT), dashboard_card],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=20,
            )
        )

    async def pick_files_click(self, e):
        try:
            await self.file_picker_open.pick_files(
                dialog_title="Selecciona facturas en PDF",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["pdf"],
                allow_multiple=True
            )
        except Exception as exc:
            self.status.value = f"Error al abrir el selector: {exc}"
            self.page.update()

    def on_files_selected(self, e: ft.FilePickerResultEvent):
        if e.files:
            self.selected_files_paths = [f.path for f in e.files]
            self.file_path.value = f"{len(e.files)} factura(s) seleccionada(s)."
            self.file_path.italic = False
            self.status.value = "Lote cargado. Listo para extraer."
        else:
            self.status.value = "Selección cancelada."
        self.page.update()

    async def start_processing_click(self, e):
        if not self.selected_files_paths:
            self.status.value = "Por favor, selecciona primero archivos PDF."
            self.status.color = ft.Colors.RED_600
            self.page.update()
            return

        try:
            await self.file_picker_save.save_file(
                dialog_title="¿Dónde deseas guardar tu reporte de Excel?",
                file_name="control_facturas.xlsx",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx"]
            )
        except Exception as exc:
            self.status.value = f"Error al abrir guardado: {exc}"
            self.page.update()

    def on_save_location_selected(self, e: ft.FilePickerResultEvent):
        if not e.path:
            self.status.value = "Exportación cancelada por el usuario."
            self.page.update()
            return

        save_path = e.path
        
        # Usamos asyncio para manejar la ejecución del hilo de forma limpia como en tu ejemplo
        asyncio.create_task(self.run_extraction_async(save_path))

    async def run_extraction_async(self, save_path: str):
        self.progress_bar.visible = True
        self.progress_ring.visible = True
        self.progress_bar.value = 0.20
        self.status.value = "Leyendo PDFs y extrayendo textos..."
        self.status.color = ft.Colors.BLUE_600
        self.preview.controls.clear()
        self.page.update()

        try:
            await asyncio.sleep(0.2)
            self.progress_bar.value = 0.50
            self.status.value = "Procesando montos y estructurando Excel..."
            self.page.update()

            # Ejecutamos tu función pesada de reporte.py en un hilo secundario
            result = await asyncio.to_thread(
                process_multiple_files_in_memory,
                self.selected_files_paths,
                save_path
            )

            # Éxito
            self.progress_bar.value = 1.0
            self.status.value = "¡Archivo guardado exitosamente!"
            self.status.color = ft.Colors.GREEN_600

            self.preview.controls = [
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.CHECK_CIRCLE, color=ft.Colors.GREEN_800, size=32),
                    title=ft.Text(f"Guardado en: {Path(save_path).name}", weight=ft.FontWeight.BOLD),
                    subtitle=ft.Text(f"Facturas procesadas: {result['rows_processed']}\nMonto Total Absoluto: ${result['grand_total']:,.2f}"),
                )
            ]
            self.page.update()

            await asyncio.sleep(0.8)
            self.progress_bar.visible = False
            self.progress_ring.visible = False
            self.page.update()

        except Exception as exc:
            self.progress_bar.visible = False
            self.progress_ring.visible = False
            self.status.value = f"Error: {exc}"
            self.status.color = ft.Colors.RED_600
            self.page.update()


async def main(page: ft.Page):
    ReportApp(page)

if __name__ == "__main__":
    ft.run(main)