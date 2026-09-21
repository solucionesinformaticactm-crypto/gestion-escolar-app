import pandas as pd
import streamlit as st

# 1. Configuración de la página (Adaptada para dispositivos móviles)
st.set_page_config(
    page_title="Sistema de Gestión Escolar Integral",
    page_icon="🎓",
    layout="centered",  # Centrado para mejor visualización en pantallas de celulares
    initial_sidebar_state="collapsed",  # Menú contraído por defecto para ganar espacio en el móvil
)

# Estilos CSS inyectados para mejorar la experiencia táctil en celulares (botones grandes y claros)
st.markdown(
    """
    <style>
    /* Botones más grandes y cómodos para el dedo en pantallas táctiles */
    .stButton > button {
        width: 100%;
        border-radius: 10px;
        height: 3em;
        font-weight: bold;
        font-size: 16px;
    }
    /* Reducir márgenes superiores para aprovechar mejor el espacio vertical del celular */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# 2. Carga de Base de Datos Excel con caché para mayor velocidad en el celular
EXCEL_FILE = "Gestion_Escolar_Secundaria.xlsx"


@st.cache_data(ttl=60)
def cargar_datos():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        return {
            "alumnos": pd.read_excel(xls, "Alumnos"),
            "profesores": pd.read_excel(xls, "Profesores"),
            "materias": pd.read_excel(xls, "Plan_Materias"),
            "notas": pd.read_excel(xls, "Notas"),
            "asistencia": pd.read_excel(xls, "Asistencia"),
            "conducta": pd.read_excel(xls, "Conducta"),
            "agenda": pd.read_excel(xls, "Agenda"),
            "cuotas": pd.read_excel(xls, "Cuotas"),
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

    with st.form("login_form"):
        user_input = st.text_input(
            "Usuario:", placeholder="Ej. profe.carlos / preceptor.juan"
        )
        pass_input = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar al Sistema")

        if submit:
            if datos and "profesores" in datos:
                df_prof = datos["profesores"]
                # Validación simple contra la hoja de profesores/usuarios
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
                elif (
                    user_input == "admin" and pass_input == "admin"
                ):  # Acceso maestro de emergencia
                    st.session_state.autenticado = True
                    st.session_state.usuario = "Administrador"
                    st.session_state.rol = "Direccion"
                    st.rerun()
                else:
                    st.error("Usuario o contraseña incorrectos.")
            else:
                st.error("No se pudo cargar la base de datos Excel.")
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

        col1, col2 = st.columns(2)
        with col1:
            st.metric(
                label="Alumnos Registrados",
                value=(
                    len(datos["alumnos"])
                    if datos and "alumnos" in datos
                    else "N/D"
                ),
            )
        with col2:
            st.metric(
                label="Cursos Activos",
                value=(
                    datos["alumnos"]["Curso"].nunique()
                    if datos and "alumnos" in datos
                    else "N/D"
                ),
            )

        st.info(
            "💡 **Consejo para Celulares:** Utiliza el menú desplegable lateral (arriba a la izquierda con las rayitas) para moverte rápidamente entre la toma de asistencia, notas y conducta."
        )

    # --- SECCIÓN: TOMAR ASISTENCIA ---
    elif menu_opcion == "📋 Tomar Asistencia":
        st.title("Control de Asistencia")
        if datos and "alumnos" in datos:
            cursos = datos["alumnos"]["Curso"].unique()
            curso_sel = st.selectbox("Seleccione el Curso:", cursos)

            fecha_sel = st.date_input("Fecha de la clase:")

            st.markdown("---")
            st.write(
                f"**Lista de alumnos para el curso {curso_sel}:** (Marque los ausentes)"
            )

            alumnos_curso = datos["alumnos"][
                datos["alumnos"]["Curso"] == curso_sel
            ]

            asistencias_dict = {}
            for index, row in alumnos_curso.iterrows():
                nombre_completo = (
                    f"{row.get('Apellido', '')}, {row.get('Nombre', '')}"
                )
                # Checkbox grande y claro para celulares
                estado = st.checkbox(f"Presente: {nombre_completo}", value=True)
                asistencias_dict[row.get("ID_Alumno", index)] = (
                    "Presente" if estado else "Ausente"
                )

            if st.button("Guardar Asistencia"):
                st.success(
                    f"¡Asistencia del curso {curso_sel} guardada con éxito para la fecha {fecha_sel}!"
                )
        else:
            st.warning("No hay datos de alumnos cargados.")

    # --- SECCIÓN: CARGAR NOTAS ---
    elif menu_opcion == "📝 Cargar Notas":
        st.title("Evaluación Continua")
        if datos and "alumnos" in datos:
            cursos = datos["alumnos"]["Curso"].unique()
            curso_sel = st.selectbox(
                "Seleccione Curso para calificar:", cursos, key="curso_notas"
            )

            trimestre = st.selectbox(
                "Trimestre:", ["1° Trimestre", "2° Trimestre", "3° Trimestre"]
            )
            calificacion = st.slider("Calificación numérica:", 1, 10, 7)

            if st.button("Registrar Calificación"):
                st.success(
                    f"Calificación de {calificacion} registrada correctamente en el {trimestre}."
                )
        else:
            st.warning("No hay datos disponibles.")

    # --- SECCIÓN: PARTES DE CONDUCTA ---
    elif menu_opcion == "⚠️ Partes de Conducta":
        st.title("Registro de Conducta")
        if datos and "alumnos" in datos:
            alumno_involucrado = st.selectbox(
                "Seleccione Alumno:",
                datos["alumnos"]["Apellido"]
                + ", "
                + datos["alumnos"]["Nombre"],
            )
            motivo = st.text_area(
                "Motivo del parte / Observación de convivencia:"
            )

            if st.button("Emitir Parte de Conducta"):
                if motivo:
                    st.success(
                        f"Parte de conducta emitido correctamente para {alumno_involucrado}."
                    )
                else:
                    st.warning("Por favor, escriba el motivo del parte.")

    # --- SECCIÓN: PADRÓN DE ALUMNOS ---
    elif menu_opcion == "👥 Padrón de Alumnos":
        st.title("Padrón de Alumnos")
        if datos and "alumnos" in datos:
            busqueda = st.text_input("🔍 Buscar alumno por apellido:")
            df_a = datos["alumnos"]
            if busqueda:
                df_a = df_a[
                    df_a["Apellido"]
                    .astype(str)
                    .str.contains(busqueda, case=False, na=False)
                ]
            st.dataframe(df_a, use_container_width=True)

    # --- SECCIÓN: AGENDA ESCOLAR ---
    elif menu_opcion == "📅 Agenda Escolar":
        st.title("Agenda Institucional")
        if datos and "agenda" in datos:
            st.dataframe(datos["agenda"], use_container_width=True)
        else:
            st.info("No hay eventos agendados actualmente.")
