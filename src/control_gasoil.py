import datetime

import flet as ft

from excel_storage import (
    append_record,
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
    session = {
        "workbook": None,
        "filename": "projecto1 ia.xlsx",
        "drivers": None,
        "dni_column": None,
        "conductor_column": None,
        "new_records": 0,
        "records_since_auto_download": 0,
    }
    state = {"step": 0}

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
    txt_km_inicio = ft.TextField(
        label="Kilómetros iniciales", width=420, keyboard_type=ft.KeyboardType.NUMBER
    )
    txt_km_final = ft.TextField(
        label="Kilómetros finales", width=420, keyboard_type=ft.KeyboardType.NUMBER
    )
    txt_km_dif = ft.TextField(
        label="Kilómetros de trabajo", width=420, read_only=True, value="0"
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
            _, conductores = read_reference_sheets(archivo.bytes)
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
            session["workbook"] = archivo.bytes
            session["filename"] = archivo.name
            session["drivers"] = conductores
            session["dni_column"] = dni_column
            session["conductor_column"] = conductor_column
            session["new_records"] = 0
            session["records_since_auto_download"] = 0
            txt_dni.value = ""
            txt_conductor.value = ""
            lbl_archivo.value = f"Archivo cargado: {archivo.name}"
            btn_descargar.visible = False
            lbl_status.value = ""
            state["step"] = 1
            update_screen()
        except Exception as ex:
            session["workbook"] = None
            session["drivers"] = None
            session["dni_column"] = None
            session["conductor_column"] = None
            set_status(f"No se pudo leer el Excel: {ex}", RED)

    async def exportar_excel(sufijo="actualizado", automatico=False):
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
            if automatico:
                set_status(
                    f"Excel descargado automáticamente tras {session['new_records']} registros nuevos.",
                    GREEN,
                )
            else:
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
        update_navigation()
        page.update()

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
        km_start = numeric_value(txt_km_inicio)
        km_end = numeric_value(txt_km_final)
        gas_start = numeric_value(txt_gas_inicio)
        gas_end = numeric_value(txt_gas_final)

        if txt_km_inicio.value or txt_km_final.value:
            if km_start is None or km_end is None:
                lbl_error_km.value = "Introduce kilómetros numéricos iguales o mayores que cero."
                txt_km_dif.value = "0"
            elif km_end < km_start:
                lbl_error_km.value = "Los kilómetros finales no pueden ser menores que los iniciales."
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
            start = numeric_value(txt_km_inicio)
            end = numeric_value(txt_km_final)
            if start is None or end is None:
                return False, "Completa ambos kilómetros con valores numéricos no negativos."
            if end < start:
                return False, "Los kilómetros finales no pueden ser menores que los iniciales."
            if end - start > 1000:
                return False, "El trabajo no puede superar los 1.000 km por registro."
            return True, "Kilometraje correcto."
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
            f"Kilómetros: {txt_km_inicio.value} → {txt_km_final.value} "
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
        registro = {
            "Hora": momento_registro.strftime("%H:%M:%S"),
            "fecha repostaje": momento_registro.strftime("%Y-%m-%d"),
            "tipo vehiculo": TYPE_LABELS[dd_tipo_vehiculo.value],
            "vehículo": dd_matricula.value,
            "DNI": normalize_dni(txt_dni.value),
            "conductor": txt_conductor.value,
            "Inicio km": numeric_value(txt_km_inicio),
            "km final": numeric_value(txt_km_final),
            "km Trabajo": numeric_value(txt_km_final) - numeric_value(txt_km_inicio),
            "Cont. Inicio": numeric_value(txt_gas_inicio),
            "Cont. Final": numeric_value(txt_gas_final),
            "Litros": numeric_value(txt_gas_final) - numeric_value(txt_gas_inicio),
        }
        try:
            session["workbook"] = append_record(session["workbook"], registro)
            session["new_records"] += 1
            session["records_since_auto_download"] += 1
            should_auto_download = session["records_since_auto_download"] >= 10
            for field in (
                txt_km_inicio, txt_km_final, txt_gas_inicio, txt_gas_final,
            ):
                field.value = ""
            dd_tipo_vehiculo.value = None
            dd_matricula.value = None
            dd_matricula.options = []
            dd_matricula.disabled = True
            txt_dni.value = ""
            txt_conductor.value = ""
            txt_km_dif.value = "0"
            txt_gas_dif.value = "0"
            lbl_error_km.value = ""
            lbl_error_gas.value = ""
            lbl_status.value = "Registro añadido correctamente. Descarga el Excel actualizado para guardarlo."
            lbl_status.color = GREEN
            btn_descargar.visible = True
            state["step"] = 1
            update_screen()
            if should_auto_download:
                total = session["new_records"]
                downloaded = await exportar_excel(
                    f"actualizado_{total}_registros", automatico=True
                )
                if downloaded:
                    session["records_since_auto_download"] = 0
        except Exception as ex:
            lbl_status.value = f"No se pudo añadir el registro: {ex}"
            lbl_status.color = RED
            page.update()

    def on_any_input():
        calculate_differences()

    dd_tipo_vehiculo.on_select = on_tipo_select
    dd_matricula.on_select = lambda: update_navigation()
    txt_dni.on_change = on_dni_change
    txt_km_inicio.on_change = on_any_input
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
            controls=[txt_km_inicio, txt_km_final, txt_km_dif, lbl_error_km],
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
        txt_km_inicio, txt_km_final, txt_km_dif,
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
                ft.Text("Control de Gasoil", size=26, weight=ft.FontWeight.BOLD, color=BLUE_DARK),
                ft.Text("Registro guiado de la flota", color=BLUE_DARK),
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
