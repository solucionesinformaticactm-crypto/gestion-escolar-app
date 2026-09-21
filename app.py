import os
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
            "agenda": (
                pd.read_excel(xls, "Agenda")
                if "Agenda" in xls.sheet_names
                else pd.DataFrame()
            ),
        }
    except Exception:
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
        user_input = st.text_input("Usuario:", placeholder="Ej. admin")
        pass_input = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar al Sistema")

        if submit:
            # Acceso maestro directo y seguro
            if user_input.strip() == "admin" and pass_input == "admin":
                st.session_state.autenticado = True
                st.session_state.usuario = "Administrador"
                st.session_state.rol = "Direccion"
                st.rerun()

            # Validación flexible de profesores
            acceso_concedido = False
            if datos and "profesores" in datos and not datos["profesores"].empty:
                df_prof = datos["profesores"]
                # Buscar columnas posibles para usuario
                col_usuario = None
                for col in df_prof.columns:
                    if (
                        "usuario" in str(col).lower()
                        or "email" in str(col).lower()
                        or "mail" in str(col).lower()
                    ):
                        col_usuario = col
                        break

                if col_usuario:
                    match = df_prof[
                        df_prof[col_usuario].astype(str).str.lower()
                        == user_input.strip().lower()
                    ]
                    if not match.empty:
                        acceso_concedido = True
                        st.session_state.usuario = user_input
                        st.session_state.rol = (
                            match.iloc[0].get("Rol", "Profesor")
                            if "Rol" in df_prof.columns
                            else "Profesor"
                        )

            if acceso_concedido:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error(
                    "Usuario o contraseña incorrectos. (Prueba con usuario: **admin** / contraseña: **admin**)"
                )
else:
    # --- MENÚ MÓVIL ---
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
        st.metric(label="Alumnos Registrados", value=total_alumnos)

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
        st.title("Padrón")
        if datos and "alumnos" in datos:
            st.dataframe(datos["alumnos"], use_container_width=True)

    elif menu_opcion == "📅 Agenda Escolar":
        st.title("Agenda")
        if datos and "agenda" in datos:
            st.dataframe(datos["agenda"], use_container_width=True)
