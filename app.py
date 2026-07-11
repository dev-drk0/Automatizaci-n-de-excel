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
        
        # Color de fondo moderno (Gris claro)
        self.page.bgcolor = "#F5F7FB"
        self.page.theme = ft.Theme(color_scheme_seed="#4C78A8")

        # Registrar FilePicker como servicio
        self.picker = ft.FilePicker()
        self.page.services.append(self.picker)

        # Componentes de UI
        self.file_path = ft.Text(
            value="Ningún archivo seleccionado",
            color="#4C78A8",
            italic=True,
        )

        self.status = ft.Text(
            value="Listo para generar el reporte",
            color=ft.Colors.GREY_700,
        )

        # 1. ProgressBar configurada correctamente
        self.progress_bar = ft.ProgressBar(
            value=0,
            visible=False,
            width=500,
            height=8,
            border_radius=10,
        )

        # 2. ProgressRing para carga simultánea
        self.progress_ring = ft.ProgressRing(
            visible=False,
            width=28,
            height=28,
        )

        # 3. Contenedor dinámico para los resultados
        self.preview = ft.Column(
            spacing=8,
            horizontal_alignment=ft.CrossAxisAlignment.START,
        )

        self.selected_file_path: str | None = None

        # --- CONSTRUCCIÓN DE LA INTERFAZ ---
        
        # Encabezado del Dashboard
        header = ft.Column(
            [
                ft.Icon(
                    ft.Icons.ANALYTICS,
                    size=56,
                    color=ft.Colors.BLUE,
                ),
                ft.Text(
                    "Automatización de Reportes",
                    size=30,
                    weight=ft.FontWeight.BOLD,
                ),
                ft.Text(
                    "Convierte un Excel en múltiples entregables automáticamente.",
                    color=ft.Colors.GREY_700,
                ),
            ],
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=10,
        )

        # Tarjeta Principal (Dashboard Blanco con Sombras)
        dashboard_card = ft.Container(
            content=ft.Column(
                [
                    # Sección Selección de Archivo
                    ft.Text("📂 Archivo", weight=ft.FontWeight.BOLD, size=16),
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
                        "Seleccionar archivo de ventas",
                        icon=ft.Icons.FOLDER_OPEN,
                        on_click=self.pick_file,
                    ),
                    
                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    # Sección Generación y Progreso
                    ft.Text("⚙️ Generación", weight=ft.FontWeight.BOLD, size=16),
                    self.status,
                    ft.Row(
                        [self.progress_ring, self.progress_bar],
                        alignment=ft.MainAxisAlignment.CENTER,
                        spacing=15,
                    ),
                    ft.FilledButton(
                        "Generar reporte",
                        icon=ft.Icons.PLAY_ARROW,
                        on_click=self.generate_report,
                    ),

                    ft.Divider(height=30, color=ft.Colors.GREY_200),

                    # Sección Resultados Finales
                    ft.Text("📁 Archivos generados", weight=ft.FontWeight.BOLD, size=16),
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

        # Renderizado final en la página
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
                dialog_title="Selecciona un archivo de ventas",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx", "xls", "csv"],
                allow_multiple=False,
            )

            if files:
                selected = Path(files[0].path)
                self.selected_file_path = str(selected)

                self.file_path.value = f"Archivo: {selected.name}"
                self.file_path.italic = False
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
            self.status.color = ft.Colors.RED_600
            self.page.update()
            return

        # Mostrar e iniciar las barras de progreso
        self.progress_bar.visible = True
        self.progress_ring.visible = True
        self.progress_bar.value = 0.10
        self.status.value = "Iniciando procesamiento..."
        self.status.color = ft.Colors.BLUE_600
        self.preview.controls.clear() 
        self.page.update()

        try:
            # Actualizaciones visuales ordenadas de progreso durante el proceso
            await asyncio.sleep(0.3)
            self.status.value = "Leyendo Excel y procesando datos..."
            self.progress_bar.value = 0.30
            self.page.update()

            # Ejecuta build_report() en un hilo en segundo plano
            result = await asyncio.to_thread(
                build_report,
                self.selected_file_path,
            )

            self.status.value = "Generando gráfica y PDF..."
            self.progress_bar.value = 0.55
            self.page.update()
            await asyncio.sleep(0.2)

            self.status.value = "Redactando cuerpo de correo electrónico..."
            self.progress_bar.value = 0.80
            self.page.update()
            await asyncio.sleep(0.2)

            # Carga exitosa de resultados
            self.progress_bar.value = 1.0
            self.status.value = "¡Reporte generado con éxito!"
            self.status.color = ft.Colors.GREEN_600

            # Inyección de la lista limpia de archivos generados con iconos a color
            self.preview.controls = [
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.TABLE_VIEW, color=ft.Colors.GREEN_700),
                    title=ft.Text(result["excel_path"].name),
                ),
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.INSERT_CHART, color=ft.Colors.BLUE_700),
                    title=ft.Text(result["chart_path"].name),
                ),
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.PICTURE_AS_PDF, color=ft.Colors.RED_700),
                    title=ft.Text(result["pdf_path"].name),
                ),
                ft.ListTile(
                    leading=ft.Icon(ft.Icons.MAIL, color=ft.Colors.AMBER_700),
                    title=ft.Text(result["email_path"].name),
                ),
            ]
            self.page.update()

            # Desvanecer los cargadores después de terminar
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