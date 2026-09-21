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
        max-width: 1100px;
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
        for hoja in xls.sheet_names:
            df = pd.read_excel(xls, hoja)
            # Normalizar nombres de columnas eliminando espacios
            df.columns = [str(c).strip() for c in df.columns]
            hojas[hoja.lower()] = df
        return hojas
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
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
        # Buscar la hoja case-insensitive
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
        "<p style='text-align: center; color: gray;'>Sistema Integral de Gestión Secundaria</p>",
        unsafe_allow_html=True,
    )

    with st.form("login_form"):
        user_input = st.text_input("Usuario / Correo:", placeholder="Ej. admin.direccion")
        pass_input = st.text_input("Contraseña:", type="password")
        submit = st.form_submit_button("Ingresar al Sistema")

        if submit:
            u_limpio = user_input.strip().lower()
            p_limpio = pass_input.strip()

            # Credenciales maestras por defecto
            credenciales_maestras = {
                "admin.direccion": ("admin123", "Direccion"),
                "admin": ("admin", "Direccion"),
                "laura.gomez": ("laura123", "Profesor"),
                "marcelo.fernandez": ("marcelo123", "Profesor"),
                "marina.torres": ("marina123", "Profesor"),
                "preceptoria": ("preceptor123", "Preceptor"),
            }

            acceso_ok = False
            rol_encontrado = "Profesor"

            # 1. Validar contra credenciales maestras
            if u_limpio in credenciales_maestras:
                pass_valida, rol_valido = credenciales_maestras[u_limpio]
                if p_limpio == pass_valida:
                    acceso_ok = True
                    rol_encontrado = rol_valido

            # 2. Validar dinámicamente contra la hoja "profesores" del Excel
            if not acceso_ok and datos and "profesores" in datos:
                df_p = datos["profesores"]
                for idx, row in df_p.iterrows():
                    fila_str = [str(val).strip().lower() for val in row.values]
                    if u_limpio in fila_str:
                        acceso_ok = True
                        # Buscar si hay una columna de rol explícita
                        for col in df_p.columns:
                            if "rol" in col.lower():
                                val_rol = str(row[col]).strip()
                                if val_rol and val_rol != "nan":
                                    rol_encontrado = val_rol
                        break

            if acceso_ok:
                st.session_state.autenticado = True
                st.session_state.usuario = user_input.strip()
                st.session_state.rol = rol_encontrado
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")

    st.markdown("<br>", unsafe_allow_html=True)
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
# 4. APLICACIÓN PRINCIPAL (INTEGRAL POR ROL)
# ==========================================
else:
    st.sidebar.title(f"👤 {st.session_state.usuario}")
    st.sidebar.markdown(f"**Rol:** `{st.session_state.rol}`")
    st.sidebar.markdown("---")

    # Menú completo y estructurado
    modulos = [
        "📊 Tablero Estadístico",
        "👥 Padrón de Alumnos",
        "👨‍👩‍👧 Padrón de Tutores",
        "📋 Control de Asistencia",
        "📝 Calificaciones y Boletín",
        "⚠️ Partes de Conducta",
        "📅 Agenda Institucional",
        "💰 Cuotas y Morosidad",
    ]

    if st.session_state.rol in ["Direccion", "Preceptor"]:
        modulos.append("👨‍🏫 Gestión de Docentes")

    opcion = st.sidebar.radio("Navegación:", modulos)

    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    # ------------------------------------------
    # 1. TABLERO ESTADÍSTICO
    # ------------------------------------------
    if opcion == "📊 Tablero Estadístico":
        st.title("📊 Panel de Control Institucional")
        df_a = datos.get("alumnos", pd.DataFrame())
        df_p = datos.get("profesores", pd.DataFrame())
        df_t = datos.get("tutores", pd.DataFrame())

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Alumnos", len(df_a) if not df_a.empty else 0)
        with col2:
            st.metric("Total Docentes", len(df_p) if not df_p.empty else 0)
        with col3:
            st.metric("Total Tutores", len(df_t) if not df_t.empty else 0)

        st.markdown("---")
        if not df_a.empty:
            col_curso = next((c for c in df_a.columns if "curso" in c.lower() or "año" in c.lower()), None)
            if col_curso:
                st.write("**Distribución por Curso / Año:**")
                st.bar_chart(df_a[col_curso].value_counts())

    # ------------------------------------------
    # 2. PADRÓN DE ALUMNOS
    # ------------------------------------------
    elif opcion == "👥 Padrón de Alumnos":
        st.title("👥 Padrón General de Alumnos")
        df_a = datos.get("alumnos", pd.DataFrame())
        if not df_a.empty:
            busq = st.text_input("🔍 Buscar alumno por nombre, apellido o DNI:")
            if busq:
                mask = df_a.astype(str).apply(lambda x: x.str.contains(busq, case=False, na=False)).any(axis=1)
                df_a = df_a[mask]
            st.dataframe(df_a, use_container_width=True)
        else:
            st.info("No hay datos cargados en la hoja de Alumnos.")

    # ------------------------------------------
    # 3. PADRÓN DE TUTORES
    # ------------------------------------------
    elif opcion == "👨‍👩‍👧 Padrón de Tutores":
        st.title("👨‍👩‍👧 Padrón de Tutores y Contactos")
        df_t = datos.get("tutores", pd.DataFrame())
        if not df_t.empty:
            st.dataframe(df_t, use_container_width=True)
        else:
            st.info("No hay registros en la hoja de Tutores.")

    # ------------------------------------------
    # 4. CONTROL DE ASISTENCIA
    # ------------------------------------------
    elif opcion == "📋 Control de Asistencia":
        st.title("📋 Registro y Control de Asistencia")
        df_a = datos.get("alumnos", pd.DataFrame())
        col_curso = next((c for c in df_a.columns if "curso" in c.lower() or "año" in c.lower()), None)

        if not df_a.empty and col_curso:
            cursos = sorted(df_a[col_curso].dropna().unique())
            c_sel = st.selectbox("Seleccione el Curso:", cursos)
            fecha_sel = st.date_input("Fecha de Asistencia:", datetime.date.today())
            alumnos_curso = df_a[df_a[col_curso] == c_sel]

            st.write(f"**Alumnos del curso {c_sel} ({len(alumnos_curso)}):**")
            with st.form("form_asistencia"):
                asistencias_reg = []
                for idx, row in alumnos_curso.iterrows():
                    nom = row.get("Nombre_Alumno", row.get("Nombre", ""))
                    ape = row.get("Apellido_Alumno", row.get("Apellido", ""))
                    id_al = row.get("ID_Alumno", f"ALU-{idx}")

                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.write(f"**{ape}, {nom}**")
                    with col2:
                        estado = st.selectbox(
                            f"est_{id_al}",
                            ["Presente", "Ausente", "Media Falta", "Justificado"],
                            key=f"asist_{id_al}",
                            label_visibility="collapsed",
                        )
                    asistencias_reg.append({"Alumno": f"{ape}, {nom}", "Estado": estado})

                if st.form_submit_button("Guardar Asistencia"):
                    for reg in asistencias_reg:
                        guardar_fila_excel(
                            "Asistencia",
                            {
                                "ID_Asistencia": f"ASI-{datetime.datetime.now().strftime('%M%S')}",
                                "Fecha": str(fecha_sel),
                                "Curso": c_sel,
                                "Alumno": reg["Alumno"],
                                "Estado_Asistencia": reg["Estado"],
                            },
                        )
                    st.success("✅ ¡Asistencia guardada correctamente en Excel!")
                    st.rerun()
        else:
            st.warning("No se encontró la columna de cursos o datos en Alumnos.")

    # ------------------------------------------
    # 5. CALIFICACIONES Y BOLETÍN
    # ------------------------------------------
    elif opcion == "📝 Calificaciones y Boletín":
        st.title("📝 Evaluaciones y Boletines")
        tab1, tab2 = st.tabs(["➕ Registrar Calificación", "📜 Ver Boletín Escolar"])
        df_a = datos.get("alumnos", pd.DataFrame())
        df_m = datos.get("materias", pd.DataFrame())
        df_n = datos.get("notas", pd.DataFrame())

        with tab1:
            with st.form("form_nota", clear_on_submit=True):
                alumnos_list = (
                    (df_a["Apellido_Alumno"].astype(str) + ", " + df_a["Nombre_Alumno"].astype(str)).tolist()
                    if not df_a.empty and "Apellido_Alumno" in df_a.columns
                    else []
                )
                alumno_sel = st.selectbox("Alumno:", alumnos_list if alumnos_list else ["Sin alumnos"])
                trimestre = st.selectbox("Trimestre:", ["1er Trimestre", "2do Trimestre", "3er Trimestre"])
                materia_sel = (
                    st.selectbox("Materia:", df_m["Nombre_Materia"].unique())
                    if not df_m.empty and "Nombre_Materia" in df_m.columns
                    else st.text_input("Materia:")
                )
                nota_val = st.number_input("Calificación (1 a 10):", min_value=1.0, max_value=10.0, value=7.0)
                obs_doc = st.text_area("Observaciones del Docente:")

                if st.form_submit_button("Guardar Calificación"):
                    guardar_fila_excel(
                        "Notas",
                        {
                            "ID_Nota": f"NOT-{datetime.datetime.now().strftime('%M%S')}",
                            "Alumno": alumno_sel,
                            "Materia": materia_sel,
                            "Trimestre": trimestre,
                            "Calificacion": nota_val,
                            "Observaciones_Docente": obs_doc,
                            "Fecha_Registro": str(datetime.date.today()),
                        },
                    )
                    st.success("✅ Calificación registrada con éxito.")
                    st.rerun()

        with tab2:
            if not df_n.empty and "Alumno" in df_n.columns:
                sel_b = st.selectbox("Seleccione Alumno:", df_n["Alumno"].unique())
                df_bol = df_n[df_n["Alumno"] == sel_b]
                st.dataframe(df_bol, use_container_width=True)
                prom = pd.to_numeric(df_bol["Calificacion"], errors="coerce").mean()
                st.info(f"📈 **Promedio General:** `{prom:.2f}`")
            else:
                st.info("No hay calificaciones registradas.")

    # ------------------------------------------
    # 6. PARTES DE CONDUCTA
    # ------------------------------------------
    elif opcion == "⚠️ Partes de Conducta":
        st.title("⚠️ Registro de Conducta y Convivencia")
        df_a = datos.get("alumnos", pd.DataFrame())
        df_c = datos.get("conducta", pd.DataFrame())

        with st.form("form_conducta", clear_on_submit=True):
            alumns = (
                (df_a["Apellido_Alumno"].astype(str) + ", " + df_a["Nombre_Alumno"].astype(str)).tolist()
                if not df_a.empty and "Apellido_Alumno" in df_a.columns
                else []
            )
            al_cond = st.selectbox("Alumno:", alumns if alumns else ["Sin alumnos"])
            motivo = st.text_area("Motivo del parte:")
            sancion = st.selectbox("Medida Aplicada:", ["Apercibimiento", "Citación a Tutor", "Suspensión", "Acta de Compromiso"])

            if st.form_submit_button("Emitir Parte"):
                guardar_fila_excel(
                    "Conducta",
                    {
                        "ID_Conducta": f"CND-{datetime.datetime.now().strftime('%M%S')}",
                        "Alumno": al_cond,
                        "Fecha": str(datetime.date.today()),
                        "Motivo": motivo,
                        "Sancion": sancion,
                        "Registrado_Por": st.session_state.usuario,
                    },
                )
                st.success("✅ Parte emitido correctamente.")
                st.rerun()

        if not df_c.empty:
            st.dataframe(df_c, use_container_width=True)

    # ------------------------------------------
    # 7. AGENDA INSTITUCIONAL
    # ------------------------------------------
    elif opcion == "📅 Agenda Institucional":
        st.title("📅 Agenda Institucional")
        df_ag = datos.get("agenda", pd.DataFrame())

        if st.session_state.rol in ["Direccion", "Preceptor"]:
            with st.form("form_agenda", clear_on_submit=True):
                f_ev = st.date_input("Fecha del Evento:")
                t_ev = st.text_input("Título:")
                d_ev = st.text_area("Descripción:")
                if st.form_submit_button("Publicar en Agenda"):
                    guardar_fila_excel(
                        "Agenda",
                        {
                            "ID_Evento": f"EVT-{datetime.datetime.now().strftime('%M%S')}",
                            "Fecha": str(f_ev),
                            "Titulo": t_ev,
                            "Descripcion": d_ev,
                        },
                    )
                    st.success("✅ Evento publicado.")
                    st.rerun()

        if not df_ag.empty:
            st.dataframe(df_ag, use_container_width=True)
        else:
            st.info("No hay eventos agendados.")

    # ------------------------------------------
    # 8. CUOTAS Y MOROSIDAD
    # ------------------------------------------
    elif opcion == "💰 Cuotas y Morosidad":
        st.title("💰 Gestión de Cuotas y Pagos")
        df_cuotas = datos.get("cuotas", pd.DataFrame())
        if not df_cuotas.empty:
            st.dataframe(df_cuotas, use_container_width=True)
        else:
            st.info("No hay registros de cuotas.")

    # ------------------------------------------
    # 9. GESTIÓN DE DOCENTES (DIRECCIÓN / PRECEPTORÍA)
    # ------------------------------------------
    elif opcion == "👨‍🏫 Gestión de Docentes":
        st.title("👨‍🏫 Alta y Administración de Personal")
        st.markdown("Complete los datos para registrar un docente o preceptor:")

        with st.form("form_nuevo_docente", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                id_prof = st.text_input("ID Profesor (Ej. PROF-08):")
                apellido = st.text_input("Apellido:")
                dni = st.text_input("DNI:")
                email = st.text_input("Email:")
                materia_p = st.text_input("Materia Principal:")
                rol = st.selectbox("Rol Asignado:", ["Profesor", "Preceptor", "Direccion"])
            with col_b:
                nombre = st.text_input("Nombre:")
                telefono = st.text_input("Teléfono:")
                cursos = st.text_input("Cursos Asignados:")
                usuario_app = st.text_input("Usuario para la App:")
                pass_app = st.text_input("Contraseña para la App:", type="password")

            if st.form_submit_button("Guardar Docente en Excel"):
                if id_prof and apellido and nombre and usuario_app and pass_app:
                    docente_dict = {
                        "ID_Profesor": id_prof,
                        "Apellido": apellido,
                        "Nombre": nombre,
                        "DNI": dni,
                        "Telefono": telefono,
                        "Email": email,
                        "Materia_Principal": materia_p,
                        "Cursos_Asigned": cursos,
                        "Usuario": usuario_app,
                        "Contraseña": pass_app,
                        "Rol": rol,
                    }
                    if guardar_fila_excel("Profesores", docente_dict):
                        st.success(f"✅ ¡Docente {apellido}, {nombre} guardado! Ya puede iniciar sesión con el usuario `{usuario_app}`.")
                        st.rerun()
                else:
                    st.warning("⚠️ Complete ID, Apellido, Nombre, Usuario y Contraseña.")

        st.markdown("---")
        st.subheader("Listado Actualizado de Docentes")
        df_p_fresco = datos.get("profesores", pd.DataFrame())
        if not df_p_fresco.empty:
            st.dataframe(df_p_fresco, use_container_width=True)
        else:
            st.info("No hay docentes registrados.")
