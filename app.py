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

# Estilos CSS Responsivos Mobile-First
st.markdown(
    """
    <style>
    /* Contenedor principal ajustado para celulares */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
        max-width: 1000px;
    }
    
    /* Botones táctiles de tamaño cómodo */
    .stButton > button {
        width: 100%;
        border-radius: 12px;
        min-height: 3.2em;
        font-weight: 700;
        font-size: 16px;
        margin-top: 0.3rem;
        margin-bottom: 0.3rem;
        transition: all 0.2s ease-in-out;
    }

    /* Evitar Zoom molesto en iOS Safari al tocar un campo */
    .stTextInput input, .stSelectbox select, .stNumberInput input, .stTextArea textarea {
        font-size: 16px !important;
        border-radius: 10px !important;
    }

    /* Tarjetas de información estilo App Nativa */
    .mobile-card {
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 12px;
    }

    /* Adaptación de tablas para pantallas pequeñas */
    .stDataFrame {
        width: 100% !important;
        overflow-x: auto !important;
    }

    @media (max-width: 768px) {
        h1 { font-size: 1.8rem !important; }
        h2 { font-size: 1.4rem !important; }
        h3 { font-size: 1.2rem !important; }
        div[data-testid="stSidebar"] {
            width: 85% !important;
        }
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
                df.columns = [str(c).strip() for c in df.columns]
                hojas[clave] = df
            else:
                hojas[clave] = pd.DataFrame()
        return hojas
    except Exception:
        return None


def guardar_fila_excel(nombre_hoja, dict_datos):
    ruta = buscar_excel()

    # Si el archivo no existe, lo crea
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

        # Leer o inicializar cabeceras
        headers = [
            str(cell.value).strip() if cell.value is not None else ""
            for cell in ws[1]
        ]
        if not headers or all(h == "" for h in headers):
            headers = list(dict_datos.keys())
            for col_idx, h_name in enumerate(headers, start=1):
                ws.cell(row=1, column=col_idx, value=h_name)

        # Asegurar columnas faltantes
        for k in dict_datos.keys():
            if k not in headers:
                headers.append(k)
                ws.cell(row=1, column=len(headers), value=k)

        # Armar fila respetando orden de cabecera
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
            # 1. Login Maestro Emergencia
            if user_input.strip() == "admin" and pass_input == "admin":
                st.session_state.autenticado = True
                st.session_state.usuario = "Administrador"
                st.session_state.rol = "Direccion"
                st.rerun()

            # 2. Validación contra Profesores / Personal
            acceso_ok = False
            rol_user = "Profesor"

            if (
                datos
                and "profesores" in datos
                and not datos["profesores"].empty
            ):
                df_p = datos["profesores"]

                col_u = next(
                    (
                        c
                        for c in df_p.columns
                        if any(
                            k in c.lower() for k in ["usuario", "email", "mail"]
                        )
                    ),
                    None,
                )
                col_p = next(
                    (
                        c
                        for c in df_p.columns
                        if any(
                            k in c.lower()
                            for k in [
                                "contraseña",
                                "contrasena",
                                "password",
                                "clave",
                            ]
                        )
                    ),
                    None,
                )

                if col_u:
                    match = df_p[
                        df_p[col_u].astype(str).str.lower()
                        == user_input.strip().lower()
                    ]
                    if not match.empty:
                        if col_p:
                            pass_excel = str(match.iloc[0][col_p]).strip()
                            if pass_excel == pass_input.strip():
                                acceso_ok = True
                        else:
                            acceso_ok = True

                        if acceso_ok:
                            st.session_state.usuario = user_input.strip()
                            if "Rol" in df_p.columns:
                                r = str(match.iloc[0]["Rol"])
                                st.session_state.rol = (
                                    r if r != "nan" else "Profesor"
                                )
                            else:
                                st.session_state.rol = "Profesor"

            if acceso_ok:
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.error(
                    "Credenciales incorrectas. Para acceso directo use: admin / admin"
                )

# ==========================================
# 4. APLICACIÓN PRINCIPAL (ROL-BASED UI)
# ==========================================
else:
    # Sidebar de perfil
    st.sidebar.title(f"👤 {st.session_state.usuario}")
    st.sidebar.markdown(f"**Rol:** `{st.session_state.rol}`")
    st.sidebar.markdown("---")

    # Módulos navegables
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

    # ------------------------------------------
    # MÓDULO 1: TABLERO PRINCIPAL
    # ------------------------------------------
    if opcion == "📊 Tablero Principal":
        st.title("📊 Panel de Control")

        df_a = datos.get("alumnos", pd.DataFrame())
        df_p = datos.get("profesores", pd.DataFrame())
        df_asist = datos.get("asistencia", pd.DataFrame())

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric(
                "Alumnos",
                len(df_a) if not df_a.empty else 0,
            )
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
            st.metric(
                "Docentes",
                len(df_p) if not df_p.empty else 0,
            )

        st.markdown("---")
        st.subheader("📌 Resumen Institucional")

        if not df_a.empty and "Curso" in df_a.columns:
            st.write("**Distribución de Alumnos por Curso:**")
            st.bar_chart(df_a["Curso"].value_counts())
        else:
            st.info("Cargue la hoja 'Alumnos' para visualizar estadísticas.")

    # ------------------------------------------
    # MÓDULO 2: ASISTENCIA DIARIA
    # ------------------------------------------
    elif opcion == "📋 Asistencia Diaria":
        st.title("📋 Control de Asistencia")

        df_a = datos.get("alumnos", pd.DataFrame())
        if not df_a.empty and "Curso" in df_a.columns:
            cursos = sorted(df_a["Curso"].dropna().unique())
            c_sel = st.selectbox("Seleccione Curso:", cursos)
            fecha_sel = st.date_input("Fecha:", datetime.date.today())

            alumnos_curso = df_a[df_a["Curso"] == c_sel]

            st.write(f"**Nómina de Alumnos ({len(alumnos_curso)}):**")

            asistencias_guardar = []
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
                        estado = st.selectbox(
                            f"Estado {id_al}",
                            ["Presente", "Ausente", "Media Falta", "Justificado"],
                            key=f"asist_{id_al}",
                            label_visibility="collapsed",
                        )

                    asistencias_guardar.append(
                        {
                            "ID_Asistencia": f"ASI-{datetime.datetime.now().strftime('%m%d%H%M%S')}-{id_al}",
                            "ID_Alumno": id_al,
                            "Curso": c_sel,
                            "Fecha": str(fecha_sel),
                            "Estado_Asistencia": estado,
                            "Observaciones": "",
                        }
                    )

                btn_asist = st.form_submit_button(
                    "Guardar Asistencia en Excel"
                )

                if btn_asist:
                    exito = True
                    for reg in asistencias_guardar:
                        if not guardar_fila_excel("Asistencia", reg):
                            exito = False
                    if exito:
                        st.success(
                            f"✅ ¡Asistencia de {c_sel} guardada con éxito!"
                        )
                        st.rerun()
        else:
            st.warning("No se encontraron alumnos registrados en la base.")

    # ------------------------------------------
    # MÓDULO 3: CALIFICACIONES Y BOLETÍN
    # ------------------------------------------
    elif opcion == "📝 Calificaciones y Boletín":
        st.title("📝 Evaluaciones y Boletines")

        sub_tab1, sub_tab2 = st.tabs(
            ["➕ Cargar Calificación", "📜 Ver Boletín Escolar"]
        )

        df_a = datos.get("alumnos", pd.DataFrame())
        df_m = datos.get("materias", pd.DataFrame())
        df_n = datos.get("notas", pd.DataFrame())

        with sub_tab1:
            st.subheader("Registrar Nueva Calificación")
            with st.form("form_nota", clear_on_submit=True):
                col_n1, col_n2 = st.columns(2)
                with col_n1:
                    if not df_a.empty:
                        alumnos_list = (
                            df_a["Apellido"].astype(str)
                            + ", "
                            + df_a["Nombre"].astype(str)
                        ).tolist()
                        alumno_sel = st.selectbox("Alumno:", alumnos_list)
                    else:
                        alumno_sel = st.text_input("ID / Nombre Alumno:")

                    trimestre = st.selectbox(
                        "Trimestre:",
                        ["1er Trimestre", "2do Trimestre", "3er Trimestre"],
                    )

                with col_n2:
                    if not df_m.empty and "Nombre_Materia" in df_m.columns:
                        materia_sel = st.selectbox(
                            "Materia:", df_m["Nombre_Materia"].unique()
                        )
                    else:
                        materia_sel = st.text_input(
                            "Materia:", placeholder="Ej. Matemática"
                        )

                    nota_val = st.number_input(
                        "Calificación (1 al 10):",
                        min_value=1.0,
                        max_value=10.0,
                        value=7.0,
                        step=0.5,
                    )

                obs_doc = st.text_area("Observación del Docente:")
                submit_nota = st.form_submit_button(
                    "Guardar Nota en Excel"
                )

                if submit_nota:
                    # Extraer ID Alumno si está disponible
                    id_al = "ALU-001"
                    if not df_a.empty and alumno_sel:
                        ap = alumno_sel.split(",")[0].strip()
                        match = df_a[
                            df_a["Apellido"].astype(str).str.strip() == ap
                        ]
                        if not match.empty:
                            id_al = match.iloc[0].get("ID_Alumno", "ALU-001")

                    reg_nota = {
                        "ID_Nota": f"NOT-{datetime.datetime.now().strftime('%M%S')}",
                        "ID_Alumno": id_al,
                        "Alumno": alumno_sel,
                        "Materia": materia_sel,
                        "Trimestre": trimestre,
                        "Calificacion": nota_val,
                        "Observaciones_Docente": obs_doc,
                        "Fecha_Registro": str(datetime.date.today()),
                    }

                    if guardar_fila_excel("Notas", reg_nota):
                        st.success("✅ Calificación registrada correctamente.")
                        st.rerun()

        with sub_tab2:
            st.subheader("Boletín de Calificaciones")
            if not df_n.empty:
                alumns = (
                    df_n["Alumno"].unique()
                    if "Alumno" in df_n.columns
                    else df_n["ID_Alumno"].unique()
                )
                sel_b = st.selectbox("Seleccione Alumno para Boletín:", alumns)

                df_bol = (
                    df_n[df_n["Alumno"] == sel_b]
                    if "Alumno" in df_n.columns
                    else df_n[df_n["ID_Alumno"] == sel_b]
                )

                if not df_bol.empty:
                    st.dataframe(df_bol, use_container_width=True)
                    if "Calificacion" in df_bol.columns:
                        prom = pd.to_numeric(
                            df_bol["Calificacion"], errors="coerce"
                        ).mean()
                        st.info(f"📈 **Promedio General:** `{prom:.2f}`")
                else:
                    st.info("No hay calificaciones registradas para este alumno.")
            else:
                st.info("No hay registros en la hoja de Notas.")

    # ------------------------------------------
    # MÓDULO 4: PARTES DE CONDUCTA
    # ------------------------------------------
    elif opcion == "⚠️ Partes de Conducta":
        st.title("⚠️ Registro de Conducta")

        df_a = datos.get("alumnos", pd.DataFrame())
        df_c = datos.get("conducta", pd.DataFrame())

        with st.form("form_conducta", clear_on_submit=True):
            if not df_a.empty:
                alumns = (
                    df_a["Apellido"].astype(str)
                    + ", "
                    + df_a["Nombre"].astype(str)
                ).tolist()
                al_cond = st.selectbox("Alumno Involucrado:", alumns)
            else:
                al_cond = st.text_input("Nombre de Alumno:")

            fecha_c = st.date_input("Fecha:", datetime.date.today())
            motivo = st.text_area("Motivo del parte / Observación de Convivencia:")
            sancion = st.selectbox(
                "Medida Aplicada:",
                [
                    "Apercibimiento Leve",
                    "Apercibimiento Grave",
                    "Citación a Tutor",
                    "Suspensión",
                    "Acta de Compromiso",
                ],
            )

            btn_cond = st.form_submit_button("Emitir Parte de Conducta")

            if btn_cond:
                if motivo:
                    reg_c = {
                        "ID_Conducta": f"CND-{datetime.datetime.now().strftime('%M%S')}",
                        "Alumno": al_cond,
                        "Fecha": str(fecha_c),
                        "Motivo": motivo,
                        "Sancion": sancion,
                        "Registrado_Por": st.session_state.usuario,
                    }
                    if guardar_fila_excel("Conducta", reg_c):
                        st.success("✅ Parte de conducta guardado con éxito.")
                        st.rerun()
                else:
                    st.warning("Escriba el motivo de la observación.")

        st.markdown("---")
        st.subheader("Historial de Partes")
        if not df_c.empty:
            st.dataframe(df_c, use_container_width=True)
        else:
            st.info("Sin registros de conducta.")

    # ------------------------------------------
    # MÓDULO 5: PADRÓN Y LEGAJOS DE ALUMNOS
    # ------------------------------------------
    elif opcion == "👥 Padrón y Legajos":
        st.title("👥 Padrón Escolar")

        df_a = datos.get("alumnos", pd.DataFrame())

        if not df_a.empty:
            busqueda = st.text_input(
                "🔍 Buscar alumno por Apellido, Nombre o DNI:"
            )

            df_fil = df_a.copy()
            if busqueda:
                mask = df_fil.astype(str).apply(
                    lambda x: x.str.contains(busqueda, case=False, na=False)
                ).any(axis=1)
                df_fil = df_fil[mask]

            st.dataframe(df_fil, use_container_width=True)
        else:
            st.info("No hay alumnos cargados en la hoja 'Alumnos'.")

    # ------------------------------------------
    # MÓDULO 6: AGENDA ESCOLAR
    # ------------------------------------------
    elif opcion == "📅 Agenda Escolar":
        st.title("📅 Agenda Institucional")

        df_ag = datos.get("agenda", pd.DataFrame())

        if st.session_state.rol in ["Direccion", "Preceptor"]:
            with st.expander("➕ Agregar Nuevo Evento a la Agenda"):
                with st.form("form_agenda", clear_on_submit=True):
                    f_ev = st.date_input("Fecha de Evento:")
                    t_ev = st.text_input("Título del Evento:")
                    d_ev = st.text_area("Descripción:")
                    dest = st.selectbox(
                        "Destinatarios:",
                        [
                            "Toda la Comunidad",
                            "Docentes",
                            "Padres/Tutores",
                            "Alumnos",
                        ],
                    )

                    if st.form_submit_button("Publicar Evento"):
                        reg_ev = {
                            "ID_Evento": f"EVT-{datetime.datetime.now().strftime('%M%S')}",
                            "Fecha": str(f_ev),
                            "Titulo": t_ev,
                            "Descripcion": d_ev,
                            "Destinatarios": dest,
                        }
                        if guardar_fila_excel("Agenda", reg_ev):
                            st.success("✅ Evento publicado en la agenda.")
                            st.rerun()

        st.markdown("---")
        if not df_ag.empty:
            st.dataframe(df_ag, use_container_width=True)
        else:
            st.info("No hay eventos agendados.")

    # ------------------------------------------
    # MÓDULO 7: CUOTAS Y MOROSIDAD
    # ------------------------------------------
    elif opcion == "💰 Cuotas y Morosidad":
        st.title("💰 Gestión de Cuotas")

        df_cuotas = datos.get("cuotas", pd.DataFrame())
        df_a = datos.get("alumnos", pd.DataFrame())

        sub_c1, sub_c2 = st.tabs(["💵 Registrar Pago", "📋 Estado de Cuentas"])

        with sub_c1:
            with st.form("form_pago", clear_on_submit=True):
                if not df_a.empty:
                    alumns = (
                        df_a["Apellido"].astype(str)
                        + ", "
                        + df_a["Nombre"].astype(str)
                    ).tolist()
                    al_pago = st.selectbox("Alumno:", alumns)
                else:
                    al_pago = st.text_input("Alumno:")

                mes_pago = st.selectbox(
                    "Mes de Cuota:",
                    [
                        "Marzo",
                        "Abril",
                        "Mayo",
                        "Junio",
                        "Julio",
                        "Agosto",
                        "Septiembre",
                        "Octubre",
                        "Noviembre",
                        "Diciembre",
                    ],
                )
                monto = st.number_input("Monto Pagado ($):", min_value=0, value=15000, step=1000)
                metodo = st.selectbox(
                    "Método de Pago:",
                    ["Efectivo", "Transferencia", "Tarjeta de Débito/Crédito"],
                )

                if st.form_submit_button("Registrar Pago en Excel"):
                    reg_pago = {
                        "ID_Pago": f"PAG-{datetime.datetime.now().strftime('%M%S')}",
                        "Alumno": al_pago,
                        "Mes_Cuota": mes_pago,
                        "Monto_Pagado": monto,
                        "Fecha_Pago": str(datetime.date.today()),
                        "Estado_Pago": "Pagado",
                        "Metodo_Pago": metodo,
                    }
                    if guardar_fila_excel("Cuotas", reg_pago):
                        st.success("✅ Pago registrado con éxito.")
                        st.rerun()

        with sub_c2:
            if not df_cuotas.empty:
                st.dataframe(df_cuotas, use_container_width=True)
            else:
                st.info("No hay registros de pagos en la hoja Cuotas.")

    # ------------------------------------------
    # MÓDULO 8: GESTIÓN DE DOCENTES (SOLO DIRECCIÓN)
    # ------------------------------------------
    elif opcion == "👨‍🏫 Gestión de Docentes":
        st.title("👨‍🏫 Alta y Administración de Profesores")

        st.markdown(
            "Complete los datos para registrar un nuevo docente con credenciales de acceso:"
        )

        with st.form("form_nuevo_docente", clear_on_submit=True):
            col_a, col_b = st.columns(2)
            with col_a:
                id_prof = st.text_input("ID Profesor (Ej. PROF-01):")
                apellido = st.text_input("Apellido:")
                dni = st.text_input("DNI:")
                email = st.text_input("Email institucional:")
                materia_p = st.text_input("Materia Principal:")
                rol = st.selectbox("Rol asignado:", ["Profesor", "Preceptor", "Direccion"])

            with col_b:
                nombre = st.text_input("Nombre:")
                telefono = st.text_input("Teléfono / WhatsApp:")
                cursos = st.text_input("Cursos Asignados (Ej. 1°A, 2°B):")
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
                            f"✅ ¡Docente {apellido}, {nombre} guardado correctamente! Ya puede ingresar con el usuario `{usuario_app}`."
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
