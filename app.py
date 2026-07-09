from __future__ import annotations

import asyncio
from pathlib import Path

import flet as ft

from reporte import build_report


class ReportApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Automatización de Reportes"
        self.page.vertical_alignment = ft.MainAxisAlignment.CENTER
        self.page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
        self.page.padding = 30
        self.page.theme = ft.Theme(color_scheme_seed="#4C78A8")

        # Registrar FilePicker como servicio
        self.picker = ft.FilePicker()
        self.page.services.append(self.picker)

        self.file_path = ft.Text(
            value="Ningún archivo seleccionado",
            color="#4C78A8",
        )

        self.status = ft.Text(
            value="Listo para generar el reporte",
            color="#666",
        )

        self.progress = ft.ProgressRing(
            visible=False,
            width=32,
            height=32,
        )

        self.result_text = ft.Text("")
        self.selected_file_path: str | None = None

        self.page.add(
            ft.Column(
                [
                    ft.Text(
                        "Automatización de Reportes de Ventas",
                        size=28,
                        weight=ft.FontWeight.BOLD,
                    ),
                    ft.Text(
                        "Selecciona un archivo Excel o CSV de ventas para generar "
                        "Excel, gráfica, PDF y un correo listo para enviar.",
                        width=650,
                    ),
                    ft.FilledButton(
                        "Seleccionar archivo de ventas",
                        on_click=self.pick_file,
                    ),
                    self.file_path,
                    ft.FilledButton(
                        "Generar reporte",
                        on_click=self.generate_report,
                    ),
                    self.progress,
                    self.status,
                    self.result_text,
                ],
                spacing=12,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            )
        )

    async def pick_file(self, e):
        try:
            files = await self.picker.pick_files(
                dialog_title="Selecciona un archivo de ventas",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx", "xls", "csv"],
                allow_multiple=False,
            )

            if files:
                selected = Path(files[0].path)
                self.selected_file_path = str(selected)

                self.file_path.value = f"Archivo: {selected.name}"
                self.status.value = "Archivo cargado. Puedes generar el reporte."
            else:
                self.status.value = "No se seleccionó ningún archivo."

            self.page.update()

        except Exception as exc:
            self.status.value = f"No se pudo abrir el selector: {exc}"
            self.page.update()

    async def generate_report(self, e):
        if not self.selected_file_path:
            self.status.value = "Selecciona primero un archivo de ventas."
            self.page.update()
            return

        self.progress.visible = True
        self.status.value = "Generando reporte..."
        self.result_text.value = ""
        self.page.update()

        try:
            # Ejecuta build_report() en un hilo sin bloquear la UI
            result = await asyncio.to_thread(
                build_report,
                self.selected_file_path,
            )

            self.progress.visible = False
            self.status.value = "Reporte generado correctamente."

            self.result_text.value = (
                f"Excel: {result['excel_path'].name}\n"
                f"Gráfica: {result['chart_path'].name}\n"
                f"PDF: {result['pdf_path'].name}\n"
                f"Correo: {result['email_path'].name}"
            )

            self.page.update()

        except Exception as exc:
            self.progress.visible = False
            self.status.value = f"Error: {exc}"
            self.page.update()


async def main(page: ft.Page):
    ReportApp(page)


if __name__ == "__main__":
    ft.run(main)