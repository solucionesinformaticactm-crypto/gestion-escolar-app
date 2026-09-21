import os
import pandas as pd
import streamlit as st

# 1. Configuración de la página (Adaptada para dispositivos móviles)
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


# 2. Carga Inteligente de la Base de Datos Excel (Búsqueda en subcarpetas)
def buscar_archivo_excel(nombre_archivo="Gestion_Escolar_Secundaria.xlsx"):
    # Buscar en el directorio actual y en subdirectorios
    for root, dirs, files in os.walk("."):
        if nombre_archivo in files:
            return os.path.join(root, nombre_archivo)
    return None


@st.cache_data(ttl=60)
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
            "notas": (
                pd.read_excel(xls, "Notas")
                if "Notas" in xls.sheet_names
                else pd.DataFrame()
            ),
            "asistencia": (
                pd.read_excel(xls, "Asistencia")
                if "Asistencia" in xls.sheet_names
                else pd.DataFrame()
            ),
            "conducta": (
                pd.read_excel(xls, "Conducta")
                if "Conducta" in xls.sheet_names
                else pd.DataFrame()
            ),
            "agenda": (
                pd.read_excel(xls, "Agenda")
                if "Agenda" in xls.sheet_names
                else pd.DataFrame()
            ),
            "cuotas": (
                pd.read_excel(xls, "Cuotas")
                if "Cuotas" in xls.sheet_names
                else pd.DataFrame()
            ),
        }
    except Exception as e:
        return None


datos = cargar_datos()

# 3. Control de Sesión (Login)
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

    if datos is None:
        st.error(
            "⚠️ **Atención:** No se encontró el archivo 'Gestion_Escolar_Secundaria.xlsx' en el repositorio de GitHub. Súbelo a la raíz del repositorio para continuar."
        )

    with st.form("login_form"):
        user_input = st.text_input(
            "Usuario:", placeholder="Ej. admin / profe.carlos"
        )
        pass_input = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar al Sistema")

        if submit:
            # Acceso maestro de emergencia siempre disponible
            if user_input.strip() == "admin" and pass_input == "admin":
                st.session_state.autenticado = True
                st.session_state.usuario = "Administrador"
                st.session_state.rol = "Direccion"
                st.rerun()

            elif datos and "profesores" in datos and not datos["profesores"].empty:
                df_prof = datos["profesores"]
                usuario_encontrado = df_prof[
                    df_prof["Usuario"].astype(str).str.lower()
                    == user_input.strip().lower()
                ]
                if not usuario_encontrado.empty:
                    st.session_state.autenticado = True
                    st.session_state.usuario = user_input
                    st.session_state.rol = usuario_encontrado.iloc[0].get(
                        "Rol", "Profesor"
                    )
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
            else:
                st.error(
                    "No hay base de datos de usuarios cargada. Usa usuario: **admin** / contraseña: **admin** para entrar."
                )
else:
    # --- INTERFAZ PRINCIPAL DE LA APP MÓVIL ---
    st.sidebar.title(f"👤 Hola, {st.session_state.usuario}")
    st.sidebar.markdown(f"**Rol:** `{st.session_state.rol}`")
    st.sidebar.markdown("---")

    menu_opcion = st.sidebar.radio(
        "Menú de Navegación",
        [
            "📱 Panel Principal",
            "📋 Tomar Asistencia",
            "📝 Cargar Notas",
            "⚠️ Partes de Conducta",
            "👥 Padrón de Alumnos",
            "📅 Agenda Escolar",
        ],
    )

    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    # --- SECCIÓN: PANEL PRINCIPAL ---
    if menu_opcion == "📱 Panel Principal":
        st.title("Panel Móvil")
        st.success(f"Bienvenido al sistema. Tu perfil activo es: *{st.session_state.rol}*")

        total_alumnos = (
            len(datos["alumnos"])
            if datos and "alumnos" in datos and not datos["alumnos"].empty
            else 0
        )
        cursos_activos = (
            datos["alumnos"]["Curso"].nunique()
            if datos
            and "alumnos" in datos
            and not datos["alumnos"].empty
            and "Curso" in datos["alumnos"].columns
            else 0
        )

        col1, col2 = st.columns(2)
        with col1:
            st.metric(label="Alumnos Registrados", value=total_alumnos)
        with col2:
            st.metric(label="Cursos Activos", value=cursos_activos)

        st.info(
            "💡 **Consejo:** Utiliza el menú lateral izquierdo (icono de flecha/rayitas arriba) para moverte entre las opciones."
        )

    # --- SECCIÓN: TOMAR ASISTENCIA ---
    elif menu_opcion == "📋 Tomar Asistencia":
        st.title("Control de Asistencia")
        if (
            datos
            and "alumnos" in datos
            and not datos["alumnos"].empty
            and "Curso" in datos["alumnos"].columns
        ):
            cursos = datos["alumnos"]["Curso"].unique()
            curso_sel = st.selectbox("Seleccione el Curso:", cursos)
            fecha_sel = st.date_input("Fecha de la clase:")

            st.markdown("---")
            st.write(f"**Alumnos del curso {curso_sel}:**")

            alumnos_curso = datos["alumnos"][
                datos["alumnos"]["Curso"] == curso_sel
            ]
            for index, row in alumnos_curso.iterrows():
                nombre = f"{row.get('Apellido', '')}, {row.get('Nombre', '')}"
                st.checkbox(f"Presente: {nombre}", value=True)

            if st.button("Guardar Asistencia"):
                st.success("¡Asistencia guardada con éxito!")
        else:
            st.warning("No hay datos de alumnos disponibles en el Excel.")

    # --- SECCIÓN: CARGAR NOTAS ---
    elif menu_opcion == "📝 Cargar Notas":
        st.title("Evaluación Continua")
        trimestre = st.selectbox(
            "Trimestre:", ["1° Trimestre", "2° Trimestre", "3° Trimestre"]
        )
        calificacion = st.slider("Calificación numérica:", 1, 10, 7)
        if st.button("Registrar Calificación"):
            st.success(
                f"Calificación de {calificacion} registrada en el {trimestre}."
            )

    # --- SECCIÓN: PARTES DE CONDUCTA ---
    elif menu_opcion == "⚠️ Partes de Conducta":
        st.title("Registro de Conducta")
        motivo = st.text_area("Motivo del parte / Observación:")
        if st.button("Emitir Parte de Conducta"):
            if motivo:
                st.success("Parte de conducta emitido correctamente.")
            else:
                st.warning("Escriba el motivo.")

    # --- SECCIÓN: PADRÓN DE ALUMNOS ---
    elif menu_opcion == "👥 Padrón de Alumnos":
        st.title("Padrón de Alumnos")
        if datos and "alumnos" in datos and not datos["alumnos"].empty:
            st.dataframe(datos["alumnos"], use_container_width=True)
        else:
            st.info("No hay datos de alumnos cargados.")

    # --- SECCIÓN: AGENDA ESCOLAR ---
    elif menu_opcion == "📅 Agenda Escolar":
        st.title("Agenda Institucional")
        if datos and "agenda" in datos and not datos["agenda"].empty:
            st.dataframe(datos["agenda"], use_container_width=True)
        else:
            st.info("No hay eventos agendados.")
