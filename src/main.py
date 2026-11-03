import flet as ft
import os

def main(page: ft.Page):
    page.title = "Control de Gasoil"
    
    # 1. Identificación del usuario actual (esto debe venir de tu sistema de login o estado)
    # Por ejemplo, supongamos que guardas las credenciales del usuario logueado:
    dni_actual = "48862156j"          # Cambia esto por la variable real donde guardas el DNI del usuario
    nombre_actual = "jesus martin gordillo" # Cambia esto por la variable real del nombre

    # 2. Función que se ejecuta al pulsar el botón para enviar el Excel
    def compartir_excel_chof(e):
        # Asegúrate de que esta ruta coincide exactamente con donde 'excel_estorage.py' guarda el archivo
        ruta_excel = "gasoil_registros.xlsx" # (Modifica por el nombre/ruta real de tu archivo excel)
        
        if os.path.exists(ruta_excel):
            # Abre el menú nativo de Android (WhatsApp, Bluetooth, Drive, etc.)
            page.share_files(
                files=[ruta_excel],
                subject="Registros Excel - Control Gasoil",
                text="Aquí tienes el archivo Excel con los registros actualizados."
            )
        else:
            page.snack_bar = ft.SnackBar(ft.Text("⚠️ El archivo Excel aún no se ha generado."))
            page.snack_bar.open = True
            page.update()

    # 3. Comprobación estricta del chófer (DNI y Nombre exactos)
    es_chofer_autorizado = (
        dni_actual.strip().lower() == "48862156j" and 
        "jesus" in nombre_actual.strip().lower()
    )

    # Creamos el botón pequeño superior izquierdo SÓLO si es el chófer autorizado
    boton_compartir = None
    if es_chofer_autorizado:
        boton_compartir = ft.IconButton(
            icon=ft.Icons.SHARE,
            icon_size=18,
            tooltip="Enviar Excel (WhatsApp / Bluetooth)",
            bgcolor=ft.Colors.BLUE_50,
            on_click=compartir_excel_chof
        )

    # 4. Cabecera del PASO 1 (colocando el botón a la izquierda si existe)
    header_paso_1 = ft.Row(
        controls=[
            *( [boton_compartir] if boton_compartir else [] ),
            ft.Text("Paso 1: Registro Inicial", size=20, weight=ft.FontWeight.BOLD)
        ],
        alignment=ft.MainAxisAlignment.START,
    )

    # Contenido de la vista del Paso 1
    vista_paso_1 = ft.Column(
        controls=[
            header_paso_1,
            ft.Divider(),
            ft.Text("Bienvenido, Jesús. Introduce los datos correspondientes al paso 1."),
            # ... el resto de tus controles del paso 1 ...
        ]
    )

    # Añadir vista inicial a la página
    page.add(vista_paso_1)

ft.app(target=main)