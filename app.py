from __future__ import annotations

import asyncio
from pathlib import Path
import flet as ft

from reporte import process_multiple_files


class ReportApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Extractor Automático de Facturas"
        self.page.vertical_alignment = ft.MainAxisAlignment.CENTER
        self.page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.page.padding = 30
        
        self.page.bgcolor = "#F5F7FB"
        self.page.theme = ft.Theme(color_scheme_seed="#4C78A8")

        self.picker = ft.FilePicker()
        self.page.services.append(self.picker)

        self.file_path = ft.Text(
            value="Ningún archivo seleccionado",
            color="#4C78A8",
            italic=True,
        )

        self.status = ft.Text(
            value="Listo para procesar lotes de archivos",
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

        # Aquí guardaremos la lista de rutas de archivos seleccionados
        self.selected_files_paths: list[str] = []

        # --- CONSTRUCCIÓN DE LA INTERFAZ ---
        header = ft.Column(
            [
                ft.Icon(
                    ft.Icons.RECEIPT_LONG,
                    size=56,
                    color=ft.Colors.BLUE_ACCENT,
                ),
                ft.Text(
                    "Extractor Automático de Facturas",
                    size=30,
                    weight=ft.FontWeight.BOLD,
                ),
                ft.Text(
                    "Sube uno o múltiples PDFs a la vez para consolidar montos, proveedores y fechas en segundos.",
                    color=ft.Colors.GREY_700,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10,
        )

        dashboard_card = ft.Container(
            content=ft.Column(
                [
                    ft.Text("📂 Archivos en Cola", weight=ft.FontWeight.BOLD, size=16),
                    ft.Container(
                        content=self.file_path,
                        bgcolor=ft.Colors.GREY_50,
                        padding=15,
                        border_radius=8,
                        border=ft.Border(
                            top=ft.BorderSide(1, ft.Colors.GREY_300),
                            bottom=ft.BorderSide(1, ft.Colors.GREY_300),
                            left=ft.BorderSide(1, ft.Colors.GREY_300),
                            right=ft.BorderSide(1, ft.Colors.GREY_300),
                        ),
                    ),
                    ft.FilledButton(
                        "Seleccionar archivos (Soporta Múltiples)",
                        icon=ft.Icons.FILE_UPLOAD,
                        on_click=self.pick_file,
                    ),
                    
                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    ft.Text("⚙️ Acciones en lote", weight=ft.FontWeight.BOLD, size=16),
                    self.status,
                    ft.Row(
                        [self.progress_ring, self.progress_bar],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=15,
                    ),
                    ft.FilledButton(
                        "Procesar Todo y Consolidar",
                        icon=ft.Icons.BOLT,
                        on_click=self.process_file_data,
                    ),

                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    ft.Text("📊 Resumen del Lote", weight=ft.FontWeight.BOLD, size=16),
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
                offset=ft.Offset(0, 4),
            ),
        )

        self.page.add(
            ft.Column(
                [header, ft.Divider(height=10, color=ft.Colors.TRANSPARENT), dashboard_card],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=20,
            )
        )

    async def pick_file(self, e):
        try:
            # ACTIVAMOS allow_multiple=True
            files = await self.picker.pick_files(
                dialog_title="Selecciona una o más facturas",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["pdf", "xlsx", "xls"],
                allow_multiple=True,
            )

            if files:
                self.selected_files_paths = [f.path for f in files]
                cant = len(files)
                self.file_path.value = f"{cant} archivo(s) seleccionados para procesar."
                self.file_path.italic = False
                self.status.value = "Lote cargado. Listo para extraer datos masivos."
            else:
                self.status.value = "No se seleccionó ningún archivo."

            self.page.update()
        except Exception as exc:
            self.status.value = f"Error al abrir el selector: {exc}"
            self.page.update()

    async def process_file_data(self, e):
        if not self.selected_files_paths:
            self.status.value = "Por favor, selecciona al menos un archivo."
            self.status.color = ft.Colors.RED_600
            self.page.update()
            return

        self.progress_bar.visible = True
        self.progress_ring.visible = True
        self.progress_bar.value = 0.15
        self.status.value = "Iniciando lectura de documentos masivos..."
        self.status.color = ft.Colors.BLUE_600
        self.preview.controls.clear() 
        self.page.update()

        try:
            await asyncio.sleep(0.4)
            self.progress_bar.value = 0.60
            self.status.value = "Mapeando estructuras y sumando montos internos..."
            self.page.update()

            # Enviamos TODA la lista de archivos a procesar juntos
            result = await asyncio.to_thread(
                process_multiple_files,
                self.selected_files_paths,
            )

            self.progress_bar.value = 0.90
            self.status.value = "Generando libro maestro de Excel..."
            self.page.update()
            await asyncio.sleep(0.3)

            self.progress_bar.value = 1.0
            self.status.value = "¡Consolidación masiva completada!"
            self.status.color = ft.Colors.GREEN_600

            output_file_name = result["output_file"].name
            self.preview.controls = [
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.DATA_EXPLORATION, color=ft.Colors.GREEN_800, size=32),
                    title=ft.Text(f"Archivo de Control: {output_file_name}", weight=ft.FontWeight.BOLD),
                    subtitle=ft.Text(
                        f"Archivos leídos: {result['rows_processed']} de {result['total_files']}\n"
                        f"Suma Total de Facturación: ${result['grand_total']:,.2f}"
                    ),
                )
            ]
            self.page.update()

            await asyncio.sleep(1.0)
            self.progress_bar.visible = False
            self.progress_ring.visible = False
            self.page.update()

        except Exception as exc:
            self.progress_bar.visible = False
            self.progress_ring.visible = False
            self.status.value = f"Error en procesamiento: {exc}"
            self.status.color = ft.Colors.RED_600
            self.page.update()


async def main(page: ft.Page):
    ReportApp(page)


if __name__ == "__main__":
    ft.app(target=main)