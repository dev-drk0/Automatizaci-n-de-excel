from __future__ import annotations

import asyncio
from pathlib import Path
import flet as ft

# Importamos la nueva función lógica
from reporte import process_input_file


class ReportApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Procesador de Archivos (Excel/PDF)"
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
            value="Listo para procesar el archivo de entrada",
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

        self.selected_file_path: str | None = None

        # --- CONSTRUCCIÓN DE LA INTERFAZ ---
        header = ft.Column(
            [
                ft.Icon(
                    ft.Icons.FILE_PRESENT,
                    size=56,
                    color=ft.Colors.BLUE_ACCENT,
                ),
                ft.Text(
                    "Extractor & Consolidador de Datos",
                    size=30,
                    weight=ft.FontWeight.BOLD,
                ),
                ft.Text(
                    "Sube un archivo Excel, CSV o PDF para estructurar y acomodar sus datos automáticamente.",
                    color=ft.Colors.GREY_700,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10,
        )

        dashboard_card = ft.Container(
            content=ft.Column(
                [
                    ft.Text("📂 Archivo de Origen", weight=ft.FontWeight.BOLD, size=16),
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
                        "Seleccionar Excel o PDF",
                        icon=ft.Icons.FILE_UPLOAD,
                        on_click=self.pick_file,
                    ),
                    
                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    ft.Text("⚙️ Acciones", weight=ft.FontWeight.BOLD, size=16),
                    self.status,
                    ft.Row(
                        [self.progress_ring, self.progress_bar],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=15,
                    ),
                    ft.FilledButton(
                        "Procesar y Acomodar",
                        icon=ft.Icons.BOLT,
                        on_click=self.process_file_data,
                    ),

                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    ft.Text("📊 Salida Estructurada", weight=ft.FontWeight.BOLD, size=16),
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
            files = await self.picker.pick_files(
                dialog_title="Selecciona un archivo origen",
                file_type=ft.FilePickerFileType.CUSTOM,
                # Agregamos 'pdf' a las extensiones permitidas
                allowed_extensions=["xlsx", "xls", "csv", "pdf"],
                allow_multiple=False,
            )

            if files:
                selected = Path(files[0].path)
                self.selected_file_path = str(selected)
                self.file_path.value = f"Archivo detectado: {selected.name}"
                self.file_path.italic = False
                self.status.value = "Archivo cargado. Listo para extraer datos."
            else:
                self.status.value = "No se seleccionó ningún archivo."

            self.page.update()
        except Exception as exc:
            self.status.value = f"Error al abrir el selector: {exc}"
            self.page.update()

    async def process_file_data(self, e):
        if not self.selected_file_path:
            self.status.value = "Por favor, selecciona primero un archivo válido."
            self.status.color = ft.Colors.RED_600
            self.page.update()
            return

        self.progress_bar.visible = True
        self.progress_ring.visible = True
        self.progress_bar.value = 0.20
        self.status.value = "Analizando el tipo de archivo..."
        self.status.color = ft.Colors.BLUE_600
        self.preview.controls.clear() 
        self.page.update()

        try:
            await asyncio.sleep(0.3)
            self.progress_bar.value = 0.50
            self.status.value = "Extrayendo y mapeando estructuras..."
            self.page.update()

            # Ejecutamos la nueva función del backend
            result = await asyncio.to_thread(
                process_input_file,
                self.selected_file_path,
            )

            self.progress_bar.value = 0.90
            self.status.value = "Escribiendo archivo de salida unificado..."
            self.page.update()
            await asyncio.sleep(0.2)

            self.progress_bar.value = 1.0
            self.status.value = "¡Datos acomodados con éxito!"
            self.status.color = ft.Colors.GREEN_600

            # Ahora mostramos solo el archivo consolidado final
            output_file_name = result["output_file"].name
           # Modifica esta parte en app.py:
            self.preview.controls = [
                ft.ListTile(
                leading=ft.Icon(ft.Icons.FEED, color=ft.Colors.GREEN_800), # <-- Cambiado a ft.Icons.FEED
             title=ft.Text(f"Resultado: {output_file_name}"),
            subtitle=ft.Text(
            f"Tipo de origen: {result['file_type_detected']} | "
            f"Registros: {result['rows_processed']} | "
            f"Monto total: ${result['total_amount']:,.2f}"
                ),
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
            self.status.value = f"Error en procesamiento: {exc}"
            self.status.color = ft.Colors.RED_600
            self.page.update()


async def main(page: ft.Page):
    ReportApp(page)


if __name__ == "__main__":
    ft.run(main)