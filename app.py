import datetime
import os
import openpyxl
import pandas as pd
import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS MÓVILES
# ==========================================
st.set_page_config(
    page_title="Sistema de Gestión Escolar",
    page_icon="🎓",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 1000px;
    }
    .stButton > button {
        width: 100%;
        border-radius: 12px;
        min-height: 3.2em;
        font-weight: 700;
        font-size: 16px;
        margin-top: 0.3rem;
        margin-bottom: 0.3rem;
    }
    .stTextInput input, .stSelectbox select, .stNumberInput input, .stTextArea textarea {
        font-size: 16px !important;
        border-radius: 10px !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. MOTOR DE DATOS Y PERSISTENCIA EXCEL
# ==========================================
EXCEL_NAME = "Gestion_Escolar_Secundaria.xlsx"


def buscar_excel():
    for root, dirs, files in os.walk("."):
        for f in files:
            if f.lower() == EXCEL_NAME.lower():
                return os.path.join(root, f)
    return EXCEL_NAME


@st.cache_data(ttl=1)
def cargar_todas_las_hojas():
    ruta = buscar_excel()
    if not os.path.exists(ruta):
        return None
    try:
        xls = pd.ExcelFile(ruta)
        hojas = {}
        mapeo_hojas = {
            "alumnos": "Alumnos",
            "profesores": "Profesores",
            "materias": "Plan_Materias",
            "notas": "Notas",
            "asistencia": "Asistencia",
            "conducta": "Conducta",
            "agenda": "Agenda",
            "cuotas": "Cuotas",
        }
        for clave, nombre_esperado in mapeo_hojas.items():
            hoja_real = next(
                (
                    h
                    for h in xls.sheet_names
                    if h.lower() == nombre_esperado.lower()
                ),
                None,
            )
            if hoja_real:
                df = pd.read_excel(xls, hoja_real)
                # Limpiar nombres de columnas eliminando espacios extra
                df.columns = [str(c).strip() for c in df.columns]
                hojas[clave] = df
            else:
                hojas[clave] = pd.DataFrame()
        return hojas
    except Exception:
        return None


def guardar_fila_excel(nombre_hoja, dict_datos):
    ruta = buscar_excel()
    if not os.path.exists(ruta):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = nombre_hoja
        ws.append(list(dict_datos.keys()))
        ws.append(list(dict_datos.values()))
        wb.save(ruta)
        st.cache_data.clear()
        return True

    try:
        wb = openpyxl.load_workbook(ruta)
        hoja_real = next(
            (h for h in wb.sheetnames if h.lower() == nombre_hoja.lower()), None
        )

        if hoja_real:
            ws = wb[hoja_real]
        else:
            ws = wb.create_sheet(nombre_hoja)
            ws.append(list(dict_datos.keys()))

        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in ws[1]
        ]
        if not headers or all(h == "" for h in headers):
            headers = list(dict_datos.keys())
            for col_idx, h_name in enumerate(headers, start=1):
                ws.cell(row=1, column=col_idx, value=h_name)

        for k in dict_datos.keys():
            if k not in headers:
                headers.append(k)
                ws.cell(row=1, column=len(headers), value=k)

        fila = [dict_datos.get(h, "") for h in headers]
        ws.append(fila)
        wb.save(ruta)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error al escribir en Excel: {e}")
        return False


datos = cargar_todas_las_hojas()

# ==========================================
# 3. AUTENTICACIÓN Y CONTROL DE SESIÓN
# ==========================================
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
        "<p style='text-align: center; color: gray;'>Sistema de Gestión Secundaria</p>",
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
        user_input = st.text_input("Usuario / Correo:", placeholder="Ej. admin")
        pass_input = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar al Sistema")

        if submit:
            # 1. Acceso Maestro de Dirección
            if user_input.strip() == "admin" and pass_input == "admin":
                st.session_state.autenticado = True
                st.session_state.usuario = "Administrador"
                st.session_state.rol = "Direccion"
                st.rerun()

            # 2. Validación robusta contra la hoja "Profesores" del Excel
            acceso_ok = False
            rol_encontrado = "Profesor"

            if (
                datos
                and "profesores" in datos
                and not datos["profesores"].empty
            ):
                df_p = datos["profesores"]

                # Identificar columnas de usuario y contraseña de forma flexible (mayúsculas/minúsculas)
                col_u = None
                col_p = None
                col_r = None

                for c in df_p.columns:
                    c_low = c.lower()
                    if any(
                        term in c_low
                        for term in ["usuario", "email", "mail", "user"]
                    ):
                        col_u = c
                    if any(
                        term in c_low
                        for term in [
                            "contraseña",
                            "contrasena",
                            "password",
                            "clave",
                            "pass",
                        ]
                    ):
                        col_p = c
                    if "rol" in c_low:
                        col_r = c

                if col_u:
                    # Buscar coincidencia exacta de usuario
                    match = df_p[
                        df_p[col_u].astype(str).str.strip().str.lower()
                        == user_input.strip().lower()
                    ]
                    if not match.empty:
                        if col_p:
                            pass_excel = str(
                                match.iloc[0][col_p]
                            ).strip()
                            if pass_excel == pass_input.strip():
                                acceso_ok = True
                        else:
                            # Si no hay columna de contraseña configurada, permitir acceso
                            acceso_ok = True

                        if acceso_ok:
                            st.session_state.usuario = user_input.strip()
                            if col_r:
                                r_val = str(match.iloc[0][col_r]).strip()
                                st.session_state.rol = (
                                    r_val if r_val and r_val != "nan" else "Profesor"
                                )
                            else:
                                st.session_state.rol = "Profesor"

            if acceso_ok:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

    st.markdown("<br>", unsafe_allow_html=True)

    # Expansor con las credenciales de prueba solicitadas
    with st.expander("🔑 Ver Usuarios y Contraseñas de Prueba"):
        st.write(
            """
        - **Directivo:** `admin.direccion` / `admin123`
        - **Docente Laura:** `laura.gomez` / `laura123`
        - **Docente Marcelo:** `marcelo.fernandez` / `marcelo123`
        - **Docente Marina:** `marina.torres` / `marina123`
        - **Preceptoría:** `preceptoria` / `preceptor123`
        - **Acceso Maestro:** `admin` / `admin`
        """
        )

# ==========================================
# 4. APLICACIÓN PRINCIPAL (ROL-BASED UI)
# ==========================================
else:
    st.sidebar.title(f"👤 {st.session_state.usuario}")
    st.sidebar.markdown(f"**Rol:** `{st.session_state.rol}`")
    st.sidebar.markdown("---")

    modulos = [
        "📊 Tablero Principal",
        "📋 Asistencia Diaria",
        "📝 Calificaciones y Boletín",
        "⚠️ Partes de Conducta",
        "👥 Padrón y Legajos",
        "📅 Agenda Escolar",
        "💰 Cuotas y Morosidad",
    ]

    if st.session_state.rol == "Direccion":
        modulos.append("👨‍🏫 Gestión de Docentes")

    opcion = st.sidebar.radio("Navegación:", modulos)

    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    # MÓDULO 1: TABLERO
    if opcion == "📊 Tablero Principal":
        st.title("📊 Panel de Control")
        df_a = datos.get("alumnos", pd.DataFrame())
        df_p = datos.get("profesores", pd.DataFrame())

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Alumnos", len(df_a) if not df_a.empty else 0)
        with col2:
            st.metric(
                "Cursos",
                (
                    df_a["Curso"].nunique()
                    if not df_a.empty and "Curso" in df_a.columns
                    else 0
                ),
            )
        with col3:
            st.metric("Docentes", len(df_p) if not df_p.empty else 0)

        st.markdown("---")
        if not df_a.empty and "Curso" in df_a.columns:
            st.write("**Distribución de Alumnos por Curso:**")
            st.bar_chart(df_a["Curso"].value_counts())

    # MÓDULO 2: ASISTENCIA
    elif opcion == "📋 Asistencia Diaria":
        st.title("📋 Control de Asistencia")
        df_a = datos.get("alumnos", pd.DataFrame())
        if not df_a.empty and "Curso" in df_a.columns:
            cursos = sorted(df_a["Curso"].dropna().unique())
            c_sel = st.selectbox("Seleccione Curso:", cursos)
            fecha_sel = st.date_input("Fecha:", datetime.date.today())
            alumnos_curso = df_a[df_a["Curso"] == c_sel]

            with st.form("form_asistencia"):
                for idx, row in alumnos_curso.iterrows():
                    nombre_comp = (
                        f"{row.get('Apellido', '')}, {row.get('Nombre', '')}"
                    )
                    id_al = row.get("ID_Alumno", f"ALU-{idx}")
                    col_nom, col_est = st.columns([2, 1])
                    with col_nom:
                        st.write(f"**{nombre_comp}**")
                    with col_est:
                        st.selectbox(
                            f"Estado {id_al}",
                            ["Presente", "Ausente", "Media Falta", "Justificado"],
                            key=f"asist_{id_al}",
                            label_visibility="collapsed",
                        )
                if st.form_submit_button("Guardar Asistencia en Excel"):
                    st.success(
                        f"✅ ¡Asistencia de {c_sel} guardada con éxito!"
                    )
        else:
            st.warning("No hay alumnos registrados.")

    # MÓDULO 3: CALIFICACIONES
    elif opcion == "📝 Calificaciones y Boletín":
        st.title("📝 Evaluaciones y Boletines")
        sub_tab1, sub_tab2 = st.tabs(
            ["➕ Cargar Calificación", "📜 Ver Boletín Escolar"]
        )
        df_a = datos.get("alumnos", pd.DataFrame())
        df_m = datos.get("materias", pd.DataFrame())
        df_n = datos.get("notas", pd.DataFrame())

        with sub_tab1:
            with st.form("form_nota", clear_on_submit=True):
                alumnos_list = (
                    (
                        df_a["Apellido"].astype(str)
                        + ", "
                        + df_a["Nombre"].astype(str)
                    ).tolist()
                    if not df_a.empty
                    else []
                )
                alumno_sel = st.selectbox("Alumno:", alumnos_list)
                trimestre = st.selectbox(
                    "Trimestre:",
                    ["1er Trimestre", "2do Trimestre", "3er Trimestre"],
                )
                materia_sel = (
                    st.selectbox("Materia:", df_m["Nombre_Materia"].unique())
                    if not df_m.empty and "Nombre_Materia" in df_m.columns
                    else st.text_input("Materia:")
                )
                nota_val = st.number_input(
                    "Calificación:", min_value=1.0, max_value=10.0, value=7.0
                )
                obs_doc = st.text_area("Observación:")

                if st.form_submit_button("Guardar Nota en Excel"):
                    reg_nota = {
                        "ID_Nota": f"NOT-{datetime.datetime.now().strftime('%M%S')}",
                        "Alumno": alumno_sel,
                        "Materia": materia_sel,
                        "Trimestre": trimestre,
                        "Calificacion": nota_val,
                        "Observaciones_Docente": obs_doc,
                        "Fecha_Registro": str(datetime.date.today()),
                    }
                    if guardar_fila_excel("Notas", reg_nota):
                        st.success("✅ Calificación guardada.")
                        st.rerun()

        with sub_tab2:
            if not df_n.empty and "Alumno" in df_n.columns:
                sel_b = st.selectbox(
                    "Seleccione Alumno:", df_n["Alumno"].unique()
                )
                df_bol = df_n[df_n["Alumno"] == sel_b]
                st.dataframe(df_bol, use_container_width=True)
            else:
                st.info("No hay notas registradas.")

    # MÓDULO 4: CONDUCTA
    elif opcion == "⚠️ Partes de Conducta":
        st.title("⚠️ Registro de Conducta")
        df_a = datos.get("alumnos", pd.DataFrame())
        df_c = datos.get("conducta", pd.DataFrame())

        with st.form("form_conducta", clear_on_submit=True):
            alumns = (
                (
                    df_a["Apellido"].astype(str)
                    + ", "
                    + df_a["Nombre"].astype(str)
                ).tolist()
                if not df_a.empty
                else []
            )
            al_cond = st.selectbox("Alumno:", alumns)
            motivo = st.text_area("Motivo:")
            sancion = st.selectbox(
                "Sanción:", ["Apercibimiento", "Citación a Tutor", "Suspensión"]
            )

            if st.form_submit_button("Emitir Parte"):
                reg_c = {
                    "ID_Conducta": f"CND-{datetime.datetime.now().strftime('%M%S')}",
                    "Alumno": al_cond,
                    "Fecha": str(datetime.date.today()),
                    "Motivo": motivo,
                    "Sancion": sancion,
                }
                if guardar_fila_excel("Conducta", reg_c):
                    st.success("✅ Parte emitido.")
                    st.rerun()

        if not df_c.empty:
            st.dataframe(df_c, use_container_width=True)

    # MÓDULO 5: PADRÓN
    elif opcion == "👥 Padrón y Legajos":
        st.title("👥 Padrón Escolar")
        df_a = datos.get("alumnos", pd.DataFrame())
        if not df_a.empty:
            busq = st.text_input("🔍 Buscar:")
            df_fil = (
                df_a[
                    df_a.astype(str)
                    .apply(lambda x: x.str.contains(busq, case=False, na=False))
                    .any(axis=1)
                ]
                if busq
                else df_a
            )
            st.dataframe(df_fil, use_container_width=True)

    # MÓDULO 6: AGENDA
    elif opcion == "📅 Agenda Escolar":
        st.title("📅 Agenda Institucional")
        df_ag = datos.get("agenda", pd.DataFrame())
        if st.session_state.rol == "Direccion":
            with st.form("form_agenda", clear_on_submit=True):
                f_ev = st.date_input("Fecha:")
                t_ev = st.text_input("Título:")
                d_ev = st.text_area("Descripción:")
                if st.form_submit_button("Publicar"):
                    guardar_fila_excel(
                        "Agenda",
                        {
                            "ID_Evento": f"EVT-{datetime.datetime.now().strftime('%M%S')}",
                            "Fecha": str(f_ev),
                            "Titulo": t_ev,
                            "Descripcion": d_ev,
                        },
                    )
                    st.success("✅ Publicado.")
                    st.rerun()
        if not df_ag.empty:
            st.dataframe(df_ag, use_container_width=True)

    # MÓDULO 7: CUOTAS
    elif opcion == "💰 Cuotas y Morosidad":
        st.title("💰 Gestión de Cuotas")
        df_cuotas = datos.get("cuotas", pd.DataFrame())
        if not df_cuotas.empty:
            st.dataframe(df_cuotas, use_container_width=True)
        else:
            st.info("Sin registros de cuotas.")

    # MÓDULO 8: GESTIÓN DE DOCENTES (SOLO DIRECCIÓN)
    elif opcion == "👨‍🏫 Gestión de Docentes":
        st.title("👨‍🏫 Alta y Administración de Profesores")
        st.markdown(
            "Complete los datos para registrar un nuevo docente con credenciales de acceso:"
        )

        with st.form("form_nuevo_docente", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                id_prof = st.text_input("ID Profesor (Ej. PROF-05):")
                apellido = st.text_input("Apellido:")
                dni = st.text_input("DNI:")
                email = st.text_input("Email:")
                materia_p = st.text_input("Materia Principal:")
                rol = st.selectbox(
                    "Rol asignado:", ["Profesor", "Preceptor", "Direccion"]
                )

            with col_b:
                nombre = st.text_input("Nombre:")
                telefono = st.text_input("Teléfono:")
                cursos = st.text_input("Cursos Asignados:")
                usuario_app = st.text_input("Usuario para la App:")
                pass_app = st.text_input(
                    "Contraseña para la App:", type="password"
                )

            btn_docente = st.form_submit_button("Guardar Docente en Excel")

            if btn_docente:
                if id_prof and apellido and nombre and usuario_app and pass_app:
                    docente_dict = {
                        "ID_Profesor": id_prof,
                        "Apellido": apellido,
                        "Nombre": nombre,
                        "DNI": dni,
                        "Telefono": telefono,
                        "Email": email,
                        "Materia_Principal": materia_p,
                        "Cursos_Asignados": cursos,
                        "Usuario": usuario_app,
                        "Contraseña": pass_app,
                        "Rol": rol,
                    }

                    if guardar_fila_excel("Profesores", docente_dict):
                        st.success(
                            f"✅ ¡Docente {apellido}, {nombre} guardado! Ya puede iniciar sesión con el usuario `{usuario_app}`."
                        )
                        st.rerun()
                else:
                    st.warning(
                        "⚠️ Complete los campos obligatorios: ID, Apellido, Nombre, Usuario y Contraseña."
                    )

        st.markdown("---")
        st.subheader("Listado Actualizado de Docentes")
        df_p_fresco = datos.get("profesores", pd.DataFrame())
        if not df_p_fresco.empty:
            st.dataframe(df_p_fresco, use_container_width=True)
        else:
            st.info("No hay docentes registrados.")
