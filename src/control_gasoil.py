import asyncio
import base64
import datetime
import os

import flet as ft

from excel_storage import (
    append_record,
    last_vehicle_kilometer,
    read_reference_sheets,
    read_vehicle_registrations,
)


TYPE_LABELS = {"coche": "Coche", "camion": "Camión"}
BLUE_PALE = ft.Colors.LIGHT_BLUE_50
BLUE_SOFT = ft.Colors.LIGHT_BLUE_100
BLUE_MAIN = ft.Colors.BLUE_700
BLUE_DARK = ft.Colors.BLUE_800
RED = ft.Colors.RED_700
GREEN = ft.Colors.GREEN_700


def find_column(dataframe, candidates):
    normalized_columns = {
        " ".join(str(column).strip().casefold().split()): column
        for column in dataframe.columns
    }
    for candidate in candidates:
        column = normalized_columns.get(" ".join(candidate.casefold().split()))
        if column is not None:
            return column
    return None


def main(page: ft.Page):
    page.title = "Control de Gasoil - Flota"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = BLUE_PALE
    page.padding = 8
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER
    page.vertical_alignment = ft.MainAxisAlignment.CENTER

    picker = ft.FilePicker()
    preferences = ft.SharedPreferences()
    storage_paths = None if page.web else ft.StoragePaths()
    persisted_workbook_name = "fleet_workbook.xlsx"
    session = {
        "workbook": None,
        "filename": "projecto1 ia.xlsx",
        "drivers": None,
        "dni_column": None,
        "conductor_column": None,
        "has_new_records": False,
    }
    state = {"step": 0, "previous_km": None}

    # ==========================================
    # 🔐 CONFIGURACIÓN DEL CHOFER AUTORIZADO
    # ==========================================
    DNI_AUTORIZADO = "48862156j"
    NOMBRE_AUTORIZADO = "jesus"

    # Función robusta para compartir o exportar el Excel sin errores de atributos
    async def compartir_excel_chof(e):
        if session["workbook"] is None:
            set_status("⚠️ No hay ningún Excel cargado para compartir.", RED)
            return

        try:
            # Si estamos en web o da problemas el puente nativo, usamos exportar/guardar archivo
            if page.web:
                await exportar_excel(sufijo="compartir")
                return

            directory = await storage_paths.get_application_documents_directory()
            os.makedirs(directory, exist_ok=True)
            path_excel_fisico = os.path.join(directory, persisted_workbook_name)

            with open(path_excel_fisico, "wb") as f:
                f.write(session["workbook"])

            if os.path.exists(path_excel_fisico):
                try:
                    page.share_files(
                        files=[path_excel_fisico],
                        subject="Registros Excel - Control Gasoil",
                        text="Aquí tienes el archivo Excel actualizado con los registros de la flota."
                    )
                except Exception:
                    # Alternativa segura si el sistema operativo bloquea la acción directa
                    await exportar_excel(sufijo="compartir")
            else:
                set_status("⚠️ El archivo Excel físico no se pudo generar.", RED)

        except Exception as ex:
            # Fallback seguro para abrir el diálogo de guardado/descarga si la API nativa falla
            try:
                await exportar_excel(sufijo="compartir")
            except Exception:
                set_status(f"⚠️ Error al compartir: {ex}", RED)

    # Creamos el botón de compartir (solo aparece si coincide tu DNI)
    def crear_boton_compartir():
        dni_actual = str(txt_dni.value or "").strip().lower()
        nombre_actual = str(txt_conductor.value or "").strip().lower()
        
        es_autorizado = (
            dni_actual == DNI_AUTORIZADO and 
            NOMBRE_AUTORIZADO in nombre_actual
        )
        
        if es_autorizado:
            return ft.IconButton(
                icon=ft.Icons.SHARE,
                icon_size=22,
                tooltip="Compartir / Guardar Excel",
                icon_color=BLUE_MAIN,
                on_click=compartir_excel_chof
            )
        return ft.Container()

    # Contenedor dinámico para el icono superior derecho
    contenedor_icono_compartir = ft.Container(content=ft.Container())

    dd_tipo_vehiculo = ft.Dropdown(
        label="Tipo de vehículo",
        width=420,
        options=[
            ft.DropdownOption(key="coche", text="Coche"),
            ft.DropdownOption(key="camion", text="Camión"),
        ],
    )
    dd_matricula = ft.Dropdown(
        label="Matrícula del vehículo",
        width=420,
        options=[],
        disabled=True,
    )
    txt_dni = ft.TextField(
        label="Escribe el DNI de la persona",
        width=420,
        hint_text="Introduce el DNI manualmente",
    )
    txt_conductor = ft.TextField(
        label="Conductor", width=420, read_only=True, value=""
    )
    txt_km_anterior = ft.TextField(
        label="Últimos kilómetros registrados para este vehículo",
        width=420,
        read_only=True,
        value="Selecciona primero la matrícula",
    )
    txt_km_final = ft.TextField(
        label="Kilómetros finales",
        width=420,
        keyboard_type=ft.KeyboardType.NUMBER,
    )
    txt_km_dif = ft.TextField(
        label="Kilómetros de trabajo (calculados)", width=420, read_only=True, value="0"
    )
    txt_gas_inicio = ft.TextField(
        label="Contador inicial de gasoil",
        width=420,
        keyboard_type=ft.KeyboardType.NUMBER,
    )
    txt_gas_final = ft.TextField(
        label="Contador final de gasoil",
        width=420,
        keyboard_type=ft.KeyboardType.NUMBER,
    )
    txt_gas_dif = ft.TextField(
        label="Litros consumidos", width=420, read_only=True, value="0"
    )

    lbl_archivo = ft.Text("Aún no se ha cargado ningún Excel.")
    lbl_validacion = ft.Text("", color=RED)
    lbl_status = ft.Text("", weight=ft.FontWeight.BOLD)
    lbl_error_km = ft.Text("", color=RED, size=12)
    lbl_error_gas = ft.Text("", color=RED, size=12)
    lbl_resumen = ft.Text("")

    def set_status(message: str, color=BLUE_DARK):
        lbl_status.value = message
        lbl_status.color = color
        page.update()

    async def guardar_copia_local(workbook_bytes, filename):
        if page.web:
            encoded = base64.b64encode(workbook_bytes).decode("ascii")
            if not await preferences.set("fleet_workbook_bytes", encoded):
                raise OSError("No se pudo guardar el Excel en el almacenamiento del navegador.")
        else:
            directory = await storage_paths.get_application_documents_directory()
            os.makedirs(directory, exist_ok=True)
            path = os.path.join(directory, persisted_workbook_name)
            with open(path, "wb") as workbook_file:
                workbook_file.write(workbook_bytes)
        await preferences.set("fleet_workbook_filename", filename)

    async def leer_copia_local():
        filename = await preferences.get("fleet_workbook_filename")
        if page.web:
            encoded = await preferences.get("fleet_workbook_bytes")
            if not encoded:
                return None, filename
            return base64.b64decode(encoded), filename

        directory = await storage_paths.get_application_documents_directory()
        path = os.path.join(directory, persisted_workbook_name)
        if not os.path.isfile(path):
            return None, filename
        with open(path, "rb") as workbook_file:
            return workbook_file.read(), filename

    def activar_excel(workbook_bytes, filename, label="Archivo cargado"):
        _, conductores = read_reference_sheets(workbook_bytes)
        dni_column = find_column(conductores, ("dni",))
        conductor_column = find_column(
            conductores,
            ("conductor", "nombre conductor", "nombre del conductor"),
        )
        if dni_column is None or conductor_column is None:
            raise ValueError(
                "La hoja de conductores debe tener columnas DNI y Conductor. "
                f"Columnas encontradas: {list(conductores.columns)}"
            )
        session["workbook"] = workbook_bytes
        session["filename"] = filename or "projecto1 ia.xlsx"
        session["drivers"] = conductores
        session["dni_column"] = dni_column
        session["conductor_column"] = conductor_column
        session["has_new_records"] = False
        state["previous_km"] = None
        dd_tipo_vehiculo.value = None
        dd_matricula.value = None
        dd_matricula.options = []
        dd_matricula.disabled = True
        for field in (txt_dni, txt_km_final, txt_gas_inicio, txt_gas_final):
            field.value = ""
        txt_conductor.value = ""
        txt_km_dif.value = "0"
        txt_gas_dif.value = "0"
        lbl_error_km.value = ""
        lbl_error_gas.value = ""
        lbl_archivo.value = f"{label}: {session['filename']}"
        lbl_status.value = ""
        state["step"] = 1
        update_screen()

    async def cargar_excel():
        seleccion = await picker.pick_files(
            dialog_title="Selecciona el libro de control de gasoil",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
            allow_multiple=False,
            with_data=True,
        )
        if not seleccion:
            return

        archivo = seleccion[0]
        if not archivo.bytes:
            set_status("No se pudo leer el archivo seleccionado.", RED)
            return

        try:
            activar_excel(archivo.bytes, archivo.name)
            try:
                await guardar_copia_local(archivo.bytes, archivo.name)
                lbl_status.value = "Excel cargado y guardado para abrirlo automáticamente la próxima vez."
                lbl_status.color = GREEN
            except Exception as storage_ex:
                lbl_status.value = f"Excel cargado para esta sesión, pero no se pudo guardar localmente: {storage_ex}"
                lbl_status.color = RED
            page.update()
        except Exception as ex:
            session["workbook"] = None
            session["drivers"] = None
            session["dni_column"] = None
            session["conductor_column"] = None
            set_status(f"No se pudo leer el Excel: {ex}", RED)

    async def restaurar_excel_guardado():
        try:
            workbook_bytes, filename = await leer_copia_local()
            if not workbook_bytes:
                return
            activar_excel(
                workbook_bytes,
                filename or "projecto1 ia.xlsx",
                label="Excel recuperado automáticamente",
            )
            lbl_status.value = "Se recuperó la última copia guardada en este dispositivo."
            lbl_status.color = GREEN
            page.update()
        except Exception as ex:
            lbl_status.value = f"No se pudo recuperar el Excel guardado: {ex}"
            lbl_status.color = RED
            page.update()

    async def exportar_excel(sufijo="actualizado"):
        if session["workbook"] is None:
            set_status("Primero carga un archivo Excel.", RED)
            return False
        try:
            nombre = session["filename"]
            if nombre.lower().endswith(".xlsx"):
                nombre = nombre[:-5]
            await picker.save_file(
                dialog_title="Guardar Excel actualizado",
                file_name=f"{nombre}_{sufijo}.xlsx",
                file_type=ft.FilePickerFileType.CUSTOM,
                allowed_extensions=["xlsx"],
                src_bytes=session["workbook"],
            )
            set_status("Descarga del Excel actualizado iniciada.", GREEN)
            return True
        except Exception as ex:
            set_status(f"No se pudo descargar el Excel: {ex}", RED)
            return False

    async def descargar_excel():
        await exportar_excel()

    def on_tipo_select():
        tipo = dd_tipo_vehiculo.value
        dd_matricula.options = []
        dd_matricula.value = None
        dd_matricula.disabled = not bool(tipo and session["workbook"])
        state["previous_km"] = None
        txt_km_anterior.value = "Selecciona primero la matrícula"
        txt_km_final.value = ""
        txt_km_dif.value = "0"
        if tipo and session["workbook"]:
            try:
                matriculas = read_vehicle_registrations(session["workbook"], tipo)
                dd_matricula.options = [
                    ft.DropdownOption(key=matricula, text=matricula)
                    for matricula in matriculas
                ]
                if not matriculas:
                    lbl_status.value = (
                        "No se encontraron matrículas para este tipo en B2:B25 y E2:E25."
                    )
                    lbl_status.color = RED
            except Exception as ex:
                lbl_status.value = f"No se pudieron cargar las matrículas: {ex}"
                lbl_status.color = RED
        calculate_differences()

    def on_matricula_select():
        registration = dd_matricula.value
        state["previous_km"] = None
        txt_km_final.value = ""
        txt_km_dif.value = "0"
        if not registration or session["workbook"] is None:
            txt_km_anterior.value = "Selecciona primero la matrícula"
        else:
            try:
                previous_km = last_vehicle_kilometer(
                    session["workbook"], registration
                )
                state["previous_km"] = previous_km
                if previous_km is None:
                    txt_km_anterior.value = "Sin registro anterior; será el kilometraje inicial"
                else:
                    txt_km_anterior.value = f"{previous_km:g} km"
            except Exception as ex:
                txt_km_anterior.value = "No se pudo consultar el historial"
                lbl_status.value = f"No se pudo consultar el kilometraje anterior: {ex}"
                lbl_status.color = RED
        calculate_differences()

    def normalize_dni(value):
        normalized = str(value or "").strip().upper()
        if normalized.endswith(".0") and normalized[:-2].isdigit():
            normalized = normalized[:-2]
        return normalized

    def on_dni_change():
        conductores = session["drivers"]
        dni = normalize_dni(txt_dni.value)
        txt_conductor.value = ""
        dni_column = session["dni_column"]
        conductor_column = session["conductor_column"]
        if conductores is not None and dni and dni_column and conductor_column:
            matches = conductores[
                conductores[dni_column].map(normalize_dni) == dni
            ]
            if not matches.empty:
                txt_conductor.value = str(matches.iloc[0][conductor_column])
        
        contenedor_icono_compartir.content = crear_boton_compartir()
        update_navigation()
        page.update()

    def numeric_value(field):
        value = (field.value or "").strip()
        if not value:
            return None
        try:
            number = float(value.replace(",", "."))
        except ValueError:
            return None
        return number if number >= 0 else None

    def calculate_differences():
        km_end = numeric_value(txt_km_final)
        km_start = state["previous_km"]
        gas_start = numeric_value(txt_gas_inicio)
        gas_end = numeric_value(txt_gas_final)

        if txt_km_final.value:
            if km_end is None:
                lbl_error_km.value = "Introduce un kilometraje numérico igual o mayor que cero."
                txt_km_dif.value = "0"
            elif km_start is None:
                lbl_error_km.value = ""
                txt_km_dif.value = "0"
            elif km_end < km_start:
                lbl_error_km.value = "El kilometraje actual no puede ser menor que el último registrado."
                txt_km_dif.value = "0"
            elif km_end - km_start > 1000:
                lbl_error_km.value = "El trabajo no puede superar los 1.000 km por registro."
                txt_km_dif.value = f"{km_end - km_start:g}"
            else:
                lbl_error_km.value = ""
                txt_km_dif.value = f"{km_end - km_start:g}"
        else:
            lbl_error_km.value = ""
            txt_km_dif.value = "0"

        if txt_gas_inicio.value or txt_gas_final.value:
            if gas_start is None or gas_end is None:
                lbl_error_gas.value = "Introduce contadores numéricos iguales o mayores que cero."
                txt_gas_dif.value = "0"
            elif gas_end < gas_start:
                lbl_error_gas.value = "El contador final no puede ser menor que el inicial."
                txt_gas_dif.value = "0"
            elif gas_end - gas_start > 100:
                lbl_error_gas.value = "El repostaje no puede superar los 100 litros por registro."
                txt_gas_dif.value = f"{gas_end - gas_start:g}"
            else:
                lbl_error_gas.value = ""
                txt_gas_dif.value = f"{gas_end - gas_start:g}"
        else:
            lbl_error_gas.value = ""
            txt_gas_dif.value = "0"

        update_navigation()
        page.update()

    def validate_step(step):
        if step == 0:
            return False, "Carga primero el archivo Excel."
        if step == 1:
            if dd_tipo_vehiculo.value not in TYPE_LABELS:
                return False, "Selecciona Coche o Camión para continuar."
            return True, "Tipo de vehículo seleccionado."
        if step == 2:
            if not dd_matricula.value:
                return False, "Selecciona una matrícula para continuar."
            return True, "Matrícula seleccionada."
        if step == 3:
            if not normalize_dni(txt_dni.value):
                return False, "Escribe el DNI de la persona para continuar."
            if not txt_conductor.value:
                return False, "El DNI no coincide con un conductor del Excel."
            return True, "DNI validado."
        if step == 4:
            end = numeric_value(txt_km_final)
            if end is None:
                return False, "Escribe el kilometraje actual con un valor numérico no negativo."
            start = state["previous_km"]
            if start is None:
                return True, "Primer registro del vehículo; el kilometraje actual quedará como referencia."
            if end < start:
                return False, "El kilometraje actual no puede ser menor que el último registrado."
            if end - start > 1000:
                return False, "El trabajo no puede superar los 1.000 km por registro."
            return True, "Kilometraje correcto; los kilómetros de trabajo se calcularon automáticamente."
        if step == 5:
            start = numeric_value(txt_gas_inicio)
            end = numeric_value(txt_gas_final)
            if start is None or end is None:
                return False, "Completa ambos contadores con valores numéricos no negativos."
            if end < start:
                return False, "El contador final no puede ser menor que el inicial."
            if end - start > 100:
                return False, "El repostaje no puede superar los 100 litros por registro."
            return True, "Contadores correctos."
        return True, "Todos los datos son correctos."

    def update_navigation():
        step = state["step"]
        valid, message = validate_step(step)
        lbl_validacion.value = message
        lbl_validacion.color = GREEN if valid else RED
        btn_adelante.bgcolor = GREEN if valid else RED
        btn_adelante.color = ft.Colors.WHITE
        if step == 6:
            btn_adelante.visible = False
            btn_anadir.visible = True
            btn_anadir.bgcolor = GREEN if valid else RED
            btn_anadir.color = ft.Colors.WHITE
            btn_anadir.disabled = not valid
        else:
            btn_adelante.visible = 1 <= step <= 5
            btn_anadir.visible = False
            btn_adelante.content = "Adelante"
        btn_atras.visible = step > 1
        progress.value = step / 6 if step else 0
        lbl_paso.value = f"Paso {step} de 6" if step else "Preparación"
        for index, panel in enumerate(step_panels):
            panel.visible = index == step
        btn_descargar.visible = step == 6 and session["has_new_records"]

    def update_screen():
        update_navigation()
        page.update()

    def go_next():
        valid, message = validate_step(state["step"])
        if not valid:
            lbl_validacion.value = message
            lbl_validacion.color = RED
            page.update()
            return
        state["step"] += 1
        if state["step"] == 6:
            refresh_summary()
        update_screen()

    def go_back():
        if state["step"] > 1:
            state["step"] -= 1
            update_screen()

    def refresh_summary():
        lbl_resumen.value = (
            "Fecha y hora: se registrarán automáticamente al añadir el registro.\n"
            f"Tipo de vehículo: {TYPE_LABELS[dd_tipo_vehiculo.value]}\n"
            f"Matrícula: {dd_matricula.value}\n"
            f"Conductor: {txt_conductor.value}\n"
            f"Kilómetros anteriores: {txt_km_anterior.value}\n"
            f"Kilometraje actual: {txt_km_final.value} "
            f"(trabajo: {txt_km_dif.value})\n"
            f"Gasoil: {txt_gas_inicio.value} → {txt_gas_final.value} "
            f"(litros: {txt_gas_dif.value})"
        )

    async def guardar_datos():
        km_valid, km_message = validate_step(4)
        gas_valid, gas_message = validate_step(5)
        if not km_valid or not gas_valid or session["workbook"] is None:
            if not km_valid:
                message = km_message
            elif not gas_valid:
                message = gas_message
            else:
                message = "Carga primero el Excel."
            lbl_validacion.value = message
            lbl_validacion.color = RED
            page.update()
            return

        momento_registro = datetime.datetime.now()
        km_final = numeric_value(txt_km_final)
        km_inicio = state["previous_km"] if state["previous_km"] is not None else km_final
        registro = {
            "Hora": momento_registro.strftime("%H:%M:%S"),
            "fecha repostaje": momento_registro.strftime("%Y-%m-%d"),
            "tipo vehiculo": TYPE_LABELS[dd_tipo_vehiculo.value],
            "vehículo": dd_matricula.value,
            "DNI": normalize_dni(txt_dni.value),
            "conductor": txt_conductor.value,
            "Inicio km": km_inicio,
            "km final": km_final,
            "km Trabajo": km_final - km_inicio,
            "Cont. Inicio": numeric_value(txt_gas_inicio),
            "Cont. Final": numeric_value(txt_gas_final),
            "Litros": numeric_value(txt_gas_final) - numeric_value(txt_gas_inicio),
        }
        try:
            session["workbook"] = append_record(session["workbook"], registro)
            session["has_new_records"] = True
            persistence_message = ""
            try:
                await guardar_copia_local(session["workbook"], session["filename"])
            except Exception as storage_ex:
                persistence_message = (
                    f" No se pudo guardar en el dispositivo: {storage_ex}. "
                    "Descarga el Excel antes de cerrar la app."
                )
            for field in (
                txt_km_final, txt_gas_inicio, txt_gas_final,
            ):
                field.value = ""
            dd_tipo_vehiculo.value = None
            dd_matricula.value = None
            dd_matricula.options = []
            dd_matricula.disabled = True
            txt_dni.value = ""
            txt_conductor.value = ""
            state["previous_km"] = None
            txt_km_anterior.value = "Selecciona primero la matrícula"
            txt_km_dif.value = "0"
            txt_gas_dif.value = "0"
            lbl_error_km.value = ""
            lbl_error_gas.value = ""
            lbl_status.value = (
                "Registro añadido y copia local actualizada."
                " El Excel actualizado estará disponible en el último paso."
                + persistence_message
            )
            lbl_status.color = RED if persistence_message else GREEN
            state["step"] = 1
            update_screen()
        except Exception as ex:
            lbl_status.value = f"No se pudo añadir el registro: {ex}"
            lbl_status.color = RED
            page.update()

    def on_any_input():
        calculate_differences()

    dd_tipo_vehiculo.on_select = on_tipo_select
    dd_matricula.on_select = on_matricula_select
    txt_dni.on_change = on_dni_change
    txt_km_final.on_change = on_any_input
    txt_gas_inicio.on_change = on_any_input
    txt_gas_final.on_change = on_any_input

    btn_cargar = ft.Button(
        content="Cargar archivo Excel",
        on_click=cargar_excel,
        bgcolor=BLUE_MAIN,
        color=ft.Colors.WHITE,
    )
    btn_atras = ft.Button(
        content="Atrás",
        on_click=go_back,
        bgcolor=BLUE_SOFT,
        color=BLUE_DARK,
        visible=False,
    )
    btn_adelante = ft.Button(
        content="Adelante",
        on_click=go_next,
        bgcolor=RED,
        color=ft.Colors.WHITE,
    )
    btn_anadir = ft.Button(
        content="Añadir registro",
        on_click=guardar_datos,
        bgcolor=GREEN,
        color=ft.Colors.WHITE,
        visible=False,
    )
    btn_descargar = ft.Button(
        content="Descargar Excel actualizado",
        on_click=descargar_excel,
        bgcolor=BLUE_MAIN,
        color=ft.Colors.WHITE,
        visible=False,
    )

    step_panels = [
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.Text("Para empezar", size=20, weight=ft.FontWeight.BOLD),
                ft.Text("Carga el libro Excel de la flota y los conductores."),
                btn_cargar,
                lbl_archivo,
            ],
        ),
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.Text("Elige el tipo que corresponde al vehículo."),
                dd_tipo_vehiculo,
            ],
        ),
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                ft.Text("Solo se mostrarán las matrículas del tipo seleccionado."),
                dd_matricula,
            ],
        ),
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[txt_dni, txt_conductor],
        ),
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[txt_km_final, lbl_error_km],
        ),
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[txt_gas_inicio, txt_gas_final, txt_gas_dif, lbl_error_gas],
        ),
        ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[lbl_resumen],
        ),
    ]

    card_width = min(560, max(240, (page.width or 900) - 24))
    card_content_width = card_width - 40
    for field in (
        dd_tipo_vehiculo, dd_matricula, txt_dni, txt_conductor,
        txt_km_anterior, txt_km_final, txt_km_dif,
        txt_gas_inicio, txt_gas_final, txt_gas_dif,
    ):
        field.width = min(420, card_content_width)

    progress = ft.ProgressBar(value=0, color=BLUE_MAIN, bgcolor=BLUE_SOFT)
    lbl_paso = ft.Text("Preparación", color=BLUE_DARK)
    navigation = ft.Row(
        tight=True,
        alignment=ft.MainAxisAlignment.CENTER,
        controls=[btn_atras, btn_adelante, btn_anadir],
    )

    header_cabecera = ft.Row(
        controls=[
            ft.Column(
                tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.START,
                controls=[
                    ft.Text("Control de Gasoil", size=24, weight=ft.FontWeight.BOLD, color=BLUE_DARK),
                    ft.Text("Registro guiado de la flota", color=BLUE_DARK),
                ],
            ),
            contenedor_icono_compartir,
        ],
        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        vertical_alignment=ft.CrossAxisAlignment.START,
    )

    card = ft.Container(
        width=card_width,
        padding=24,
        bgcolor=ft.Colors.WHITE,
        border=ft.Border.all(1, BLUE_SOFT),
        border_radius=16,
        content=ft.Column(
            tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            controls=[
                header_cabecera,
                lbl_paso,
                progress,
                ft.Container(height=10),
                *step_panels,
                lbl_validacion,
                navigation,
                btn_descargar,
                lbl_status,
            ],
        ),
    )
    page.add(
        ft.Container(
            expand=True,
            alignment=ft.Alignment.CENTER,
            content=ft.Column(
                tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                controls=[card],
            ),
        )
    )
    update_navigation()
    page.update()
    asyncio.create_task(restaurar_excel_guardado())
