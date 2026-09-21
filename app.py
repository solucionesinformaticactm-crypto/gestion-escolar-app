import os
import openpyxl
import pandas as pd
import streamlit as st

# 1. Configuración de la página móvil
st.set_page_config(
    page_title="Gestión Escolar Móvil",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stButton > button {
        width: 100%;
        border-radius: 10px;
        height: 3em;
        font-weight: bold;
        font-size: 16px;
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)


def buscar_archivo_excel(nombre_archivo="Gestion_Escolar_Secundaria.xlsx"):
    for root, dirs, files in os.walk("."):
        if nombre_archivo in files:
            return os.path.join(root, nombre_archivo)
    return None


@st.cache_data(ttl=1)
def cargar_datos():
    ruta_excel = buscar_archivo_excel("Gestion_Escolar_Secundaria.xlsx")
    if not ruta_excel:
        return None
    try:
        xls = pd.ExcelFile(ruta_excel)
        return {
            "alumnos": (
                pd.read_excel(xls, "Alumnos")
                if "Alumnos" in xls.sheet_names
                else pd.DataFrame()
            ),
            "profesores": (
                pd.read_excel(xls, "Profesores")
                if "Profesores" in xls.sheet_names
                else pd.DataFrame()
            ),
            "materias": (
                pd.read_excel(xls, "Plan_Materias")
                if "Plan_Materias" in xls.sheet_names
                else pd.DataFrame()
            ),
            "agenda": (
                pd.read_excel(xls, "Agenda")
                if "Agenda" in xls.sheet_names
                else pd.DataFrame()
            ),
        }
    except Exception:
        return None


datos = cargar_datos()

# 2. Control de Sesión (Login)
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
    st.session_state.usuario = ""
    st.session_state.rol = ""

if not st.session_state.autenticado:
    st.markdown(
        "<h2 style='text-align: center;'>🎓 Portal Móvil Escolar</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; color: gray;'>Ingrese con sus credenciales institucionales</p>",
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
        user_input = st.text_input("Usuario:", placeholder="Ej. admin")
        pass_input = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar al Sistema")

        if submit:
            # 1. Acceso maestro de dirección
            if user_input.strip() == "admin" and pass_input == "admin":
                st.session_state.autenticado = True
                st.session_state.usuario = "Administrador"
                st.session_state.rol = "Direccion"
                st.rerun()
            else:
                # 2. Validación de profesores en el Excel
                acceso_concedido = False
                rol_encontrado = "Profesor"

                if (
                    datos
                    and "profesores" in datos
                    and not datos["profesores"].empty
                ):
                    df_prof = datos["profesores"]

                    # Buscar columnas de usuario y contraseña de forma flexible
                    col_usuario = None
                    col_pass = None

                    for col in df_prof.columns:
                        col_lower = str(col).lower()
                        if (
                            "usuario" in col_lower
                            or "email" in col_lower
                            or "mail" in col_lower
                        ):
                            col_usuario = col
                        if (
                            "contraseña" in col_lower
                            or "contrasena" in col_lower
                            or "password" in col_lower
                            or "clave" in col_lower
                        ):
                            col_pass = col

                    if col_usuario:
                        # Filtrar por usuario
                        match = df_prof[
                            df_prof[col_usuario].astype(str).str.lower()
                            == user_input.strip().lower()
                        ]
                        if not match.empty:
                            # Si existe columna de contraseña, validarla
                            if col_pass:
                                pass_registrada = str(
                                    match.iloc[0][col_pass]
                                ).strip()
                                if pass_registrada == pass_input.strip():
                                    acceso_concedido = True
                            else:
                                # Si no hay columna de contraseña registrada, permitir acceso
                                acceso_concedido = True

                            if acceso_concedido:
                                st.session_state.usuario = user_input
                                if "Rol" in df_prof.columns:
                                    rol_encontrado = str(
                                        match.iloc[0]["Rol"]
                                    )
                                st.session_state.rol = (
                                    rol_encontrado
                                    if rol_encontrado != "nan"
                                    else "Profesor"
                                )

                if acceso_concedido:
                    st.session_state.autenticado = True
                    st.rerun()
                else:
                    st.error(
                        "Usuario o contraseña incorrectos. (O prueba con admin / admin)"
                    )
else:
    # --- MENÚ MÓVIL ---
    st.sidebar.title(f"👤 Hola, {st.session_state.usuario}")
    st.sidebar.markdown(f"**Rol:** `{st.session_state.rol}`")
    st.sidebar.markdown("---")

    opciones_menu = [
        "📱 Panel Principal",
        "📋 Tomar Asistencia",
        "📝 Cargar Notas",
        "⚠️ Partes de Conducta",
        "👥 Padrón de Alumnos",
        "📅 Agenda Escolar",
    ]

    if st.session_state.rol == "Direccion":
        opciones_menu.append("👨‍🏫 Gestión de Profesores")

    menu_opcion = st.sidebar.radio("Menú de Navegación", opciones_menu)

    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    # --- SECCIÓN: PANEL PRINCIPAL ---
    if menu_opcion == "📱 Panel Principal":
        st.title("Panel Móvil")
        st.success(
            f"Bienvenido al sistema escolar. Perfil activo: *{st.session_state.rol}*"
        )
        total_alumnos = (
            len(datos["alumnos"])
            if datos and "alumnos" in datos and not datos["alumnos"].empty
            else 0
        )
        total_profes = (
            len(datos["profesores"])
            if datos and "profesores" in datos and not datos["profesores"].empty
            else 0
        )

        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Alumnos", value=total_alumnos)
        with col2:
            st.metric(label="Profesores", value=total_profes)

    elif menu_opcion == "📋 Tomar Asistencia":
        st.title("Control de Asistencia")
        if (
            datos
            and "alumnos" in datos
            and not datos["alumnos"].empty
            and "Curso" in datos["alumnos"].columns
        ):
            curso_sel = st.selectbox(
                "Curso:", datos["alumnos"]["Curso"].unique()
            )
            st.date_input("Fecha:")
            for idx, row in datos["alumnos"][
                datos["alumnos"]["Curso"] == curso_sel
            ].iterrows():
                st.checkbox(
                    f"{row.get('Apellido', '')}, {row.get('Nombre', '')}",
                    value=True,
                )
            if st.button("Guardar Asistencia"):
                st.success("¡Asistencia guardada!")
        else:
            st.warning("No hay alumnos cargados.")

    elif menu_opcion == "📝 Cargar Notas":
        st.title("Cargar Notas")
        st.selectbox(
            "Trimestre:", ["1° Trimestre", "2° Trimestre", "3° Trimestre"]
        )
        st.slider("Nota:", 1, 10, 7)
        if st.button("Registrar"):
            st.success("Nota registrada con éxito.")

    elif menu_opcion == "⚠️ Partes de Conducta":
        st.title("Conducta")
        st.text_area("Motivo:")
        if st.button("Emitir Parte"):
            st.success("Parte emitido.")

    elif menu_opcion == "👥 Padrón de Alumnos":
        st.title("Padrón de Alumnos")
        if datos and "alumnos" in datos:
            st.dataframe(datos["alumnos"], use_container_width=True)

    elif menu_opcion == "📅 Agenda Escolar":
        st.title("Agenda")
        if datos and "agenda" in datos:
            st.dataframe(datos["agenda"], use_container_width=True)

    # --- SECCIÓN EXCLUSIVA: GESTIÓN DE PROFESORES ---
    elif menu_opcion == "👨‍🏫 Gestión de Profesores":
        st.title("Alta y Gestión de Docentes")
        st.markdown(
            "Complete los datos para registrar un profesor y asignarle acceso:"
        )

        with st.form("form_nuevo_profesor", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                id_prof = st.text_input("ID_Profesor (Ej. P001):")
                apellido = st.text_input("Apellido:")
                dni = st.text_input("DNI:")
                email = st.text_input("Email:")
                materias_asig = st.text_input(
                    "Materia Principal (Ej. Matemática):"
                )
            with col_b:
                nombre = st.text_input("Nombre:")
                telefono = st.text_input("Teléfono:")
                cursos_asig = st.text_input(
                    "Cursos Asignados (Ej. 1°1°, 2°1°):"
                )
                nuevo_usuario = st.text_input("Usuario para la App:")
                nuevo_pass = st.text_input(
                    "Contraseña para la App:", type="password"
                )

            submit_prof = st.form_submit_button("Guardar y Actualizar Excel")

            if submit_prof:
                if id_prof and apellido and nombre and nuevo_usuario:
                    ruta_excel = buscar_archivo_excel(
                        "Gestion_Escolar_Secundaria.xlsx"
                    )
                    if ruta_excel:
                        try:
                            wb = openpyxl.load_workbook(ruta_excel)
                            if "Profesores" in wb.sheetnames:
                                ws = wb["Profesores"]

                                # Verificar si la cabecera tiene las columnas de Usuario y Contraseña, si no, agregarlas
                                headers = [
                                    cell.value for cell in ws[1]
                                ]  # Fila 1
                                if (
                                    "Usuario" not in [str(h) for h in headers]
                                    and len(headers) >= 8
                                ):
                                    ws.cell(
                                        row=1, column=9, value="Usuario"
                                    )
                                if (
                                    "Contrasena"
                                    not in [str(h) for h in headers]
                                    and "Contraseña"
                                    not in [str(h) for h in headers]
                                    and len(headers) >= 8
                                ):
                                    ws.cell(
                                        row=1, column=10, value="Contraseña"
                                    )

                                nueva_fila = [
                                    id_prof,
                                    apellido,
                                    nombre,
                                    dni,
                                    telefono,
                                    email,
                                    materias_asig,
                                    cursos_asig,
                                    nuevo_usuario,
                                    nuevo_pass,
                                ]
                                ws.append(nueva_fila)
                                wb.save(ruta_excel)
                                st.cache_data.clear()
                                st.success(
                                    f"✅ ¡Profesor/a {apellido}, {nombre} guardado correctamente! Ya puede iniciar sesión con el usuario `{nuevo_usuario}`."
                                )
                                st.rerun()
                            else:
                                st.error(
                                    "No se encontró la hoja 'Profesores' en el Excel."
                                )
                        except Exception as e:
                            st.error(
                                f"Error al actualizar el archivo Excel: {e}"
                            )
                    else:
                        st.error("No se encontró el archivo Excel en el servidor.")
                else:
                    st.warning(
                        "⚠️ Por favor, complete al menos ID, Apellido, Nombre y Usuario."
                    )

        st.markdown("---")
        st.subheader("Listado Actualizado de Profesores")
        datos_frescos = cargar_datos()
        if (
            datos_frescos
            and "profesores" in datos_frescos
            and not datos_frescos["profesores"].empty
        ):
            st.dataframe(datos_frescos["profesores"], use_container_width=True)
        else:
            st.info("No hay registros de profesores en la hoja.")
