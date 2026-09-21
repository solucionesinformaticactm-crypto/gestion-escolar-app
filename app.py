import streamlit as st
import pandas as pd
from datetime import datetime, date
import urllib.parse
import io

# Configuración de página y estética Gris / Violeta
st.set_page_config(
    page_title="Sistema de Gestión Escolar",
    page_icon="🎓",
    layout="wide"
)

# Estilos personalizados en CSS
st.markdown("""
    <style>
    .main { background-color: #F8F9F9; }
    .stButton>button {
        background-color: #6C3483;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #512E5F;
        color: white;
    }
    .metric-card {
        background-color: #FFFFFF;
        border-left: 5px solid #6C3483;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .login-box {
        background-color: #FFFFFF;
        padding: 30px;
        border-radius: 12px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        border-top: 6px solid #6C3483;
        max-width: 450px;
        margin: auto;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = 'Gestion_Escolar_Secundaria.xlsx'

# Base de datos de Usuarios y Contraseñas
USUARIOS_DB = {
    "admin.direccion": {"password": "admin123", "nombre": "Equipo Directivo", "rol": "Directivo", "id": "DIR"},
    "laura.gomez": {"password": "laura123", "nombre": "Laura Gómez", "rol": "Docente", "id": "PR-01"},
    "marcelo.fernandez": {"password": "marcelo123", "nombre": "Marcelo Fernández", "rol": "Docente", "id": "PR-02"},
    "marina.torres": {"password": "marina123", "nombre": "Marina Torres", "rol": "Docente", "id": "PR-08"},
    "preceptoria": {"password": "preceptor123", "nombre": "Preceptoría General", "rol": "Preceptor", "id": "PRE"}
}

# Manejo de Estado de Sesión (Session State)
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_actual" not in st.session_state:
    st.session_state["usuario_actual"] = None

# Función auxiliar para generar enlace de WhatsApp
def generar_link_whatsapp(numero, mensaje):
    num_limpio = ''.join(filter(str.isdigit, str(numero)))
    mensaje_codificado = urllib.parse.quote(mensaje)
    return f"https://wa.me/{num_limpio}?text={mensaje_codificado}"

# Carga e inicialización de datos desde la planilla Excel
@st.cache_data
def cargar_datos_iniciales():
    alumnos = pd.read_excel(EXCEL_FILE, sheet_name='Alumnos')
    profesores = pd.read_excel(EXCEL_FILE, sheet_name='Profesores')
    plan = pd.read_excel(EXCEL_FILE, sheet_name='Plan_Materias')
    notas = pd.read_excel(EXCEL_FILE, sheet_name='Notas')
    asistencia = pd.read_excel(EXCEL_FILE, sheet_name='Asistencia')
    return alumnos, profesores, plan, notas, asistencia

alumnos_init, profesores_df, plan_df, notas_init, asistencia_init = cargar_datos_iniciales()

# Mantener dataframes en Session State para permitir modificaciones dinámicas
if "alumnos_df" not in st.session_state:
    st.session_state["alumnos_df"] = alumnos_init.copy()
if "notas_df" not in st.session_state:
    st.session_state["notas_df"] = notas_init.copy()
if "asistencia_df" not in st.session_state:
    st.session_state["asistencia_df"] = asistencia_init.copy()

alumnos_df = st.session_state["alumnos_df"]
notas_df = st.session_state["notas_df"]
asistencia_df = st.session_state["asistencia_df"]

# Garantizar que las columnas de evaluación continua existan en el DataFrame de Notas
cols_eval = [
    '1T_Prueba1', '1T_Oral1', '1T_Prueba2', '1T_Oral2', '1T_Participacion', '1T_Carpeta',
    '2T_Prueba1', '2T_Oral1', '2T_Prueba2', '2T_Oral2', '2T_Participacion', '2T_Carpeta',
    '3T_Prueba1', '3T_Oral1', '3T_Prueba2', '3T_Oral2', '3T_Participacion', '3T_Carpeta'
]
for col in cols_eval:
    if col not in st.session_state["notas_df"].columns:
        st.session_state["notas_df"][col] = 0.0

# -----------------------------------------------------------------------------
# 🔐 PANTALLA DE INICIO DE SESIÓN
# -----------------------------------------------------------------------------
if not st.session_state["autenticado"]:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown('<div class="login-box">', unsafe_allow_html=True)
        st.image("https://img.icons8.com/isometric/100/graduation-cap.png", width=80)
        st.title("Portal de Gestión Escolar")
        st.subheader("Iniciar Sesión")
        
        user_input = st.text_input("Usuario:", placeholder="Ej. laura.gomez, admin.direccion")
        pass_input = st.text_input("Contraseña:", type="password", placeholder="••••••••")
        
        if st.button("Ingresar al Sistema", use_container_width=True):
            if user_input in USUARIOS_DB and USUARIOS_DB[user_input]["password"] == pass_input:
                st.session_state["autenticado"] = True
                st.session_state["usuario_actual"] = USUARIOS_DB[user_input]
                st.rerun()
            else:
                st.error("⚠️ Usuario o contraseña incorrectos. Por favor verifique sus datos.")
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander("🔑 Ver Usuarios y Contraseñas de Prueba"):
            st.write("""
            - **Directivo:** `admin.direccion` / `admin123`
            - **Docente Laura:** `laura.gomez` / `laura123`
            - **Docente Marcelo:** `marcelo.fernandez` / `marcelo123`
            - **Docente Marina:** `marina.torres` / `marina123`
            - **Preceptoría:** `preceptoria` / `preceptor123`
            """)

# -----------------------------------------------------------------------------
# 🏫 SISTEMA PRINCIPAL (UNA VEZ AUTENTICADO)
# -----------------------------------------------------------------------------
else:
    user_info = st.session_state["usuario_actual"]
    
    st.sidebar.image("https://img.icons8.com/isometric/100/graduation-cap.png", width=70)
    st.sidebar.title("Portal Escolar")
    st.sidebar.success(f"👤 **{user_info['nombre']}**\n\nRol: *{user_info['rol']}*")
    
    if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state["autenticado"] = False
        st.session_state["usuario_actual"] = None
        st.rerun()

    # --- VISTA DIRECTIVO ---
    if user_info['rol'] == 'Directivo':
        st.title("🏛️ Panel de Control e Indicadores Institucionales")
        st.write("Visión general del rendimiento escolar, padrón de alumnos y asistencia del establecimiento.")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown(f'<div class="metric-card"><h4>Total Alumnos</h4><h2>{len(st.session_state["alumnos_df"])}</h2></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h4>Total Profesores</h4><h2>{len(profesores_df)}</h2></div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="metric-card"><h4>Promedio General</h4><h2>7.07</h2></div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="metric-card"><h4>Asistencia General</h4><h2>56%</h2></div>', unsafe_allow_html=True)

        st.markdown("---")

        tab_alumnos, tab_notas, tab_profesores = st.tabs(["👨‍🎓 Padrón de Alumnos y Nuevo Ingreso", "📊 Calificaciones Consolidadas", "👩‍🏫 Planta Docente"])

        with tab_alumnos:
            st.subheader("👨‍🎓 Registro General de Alumnos y Legajos")

            # Módulo de alta de nuevo alumno con los 22 campos exactos del Excel
            with st.expander("➕ Registrar Nuevo Ingreso de Alumno (Formulario Completo)", expanded=False):
                with st.form("form_nuevo_alumno_excel"):
                    st.markdown("##### 👤 1. Datos del Alumno")
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        apellido = st.text_input("Apellido:")
                        fecha_nac = st.date_input("Fecha_Nac.:", value=date(2010, 1, 1))
                        direccion = st.text_input("Dirección:")
                    with c2:
                        nombre = st.text_input("Nombre:")
                        anio = st.number_input("Año:", min_value=1, max_value=6, value=1)
                        telefono_alumno = st.text_input("Teléfono_Alumno:")
                    with c3:
                        dni = st.text_input("DNI:")
                        division = st.selectbox("División:", ["A", "B", "C"])
                        email_alumno = st.text_input("Email_Alumno:")
                    with c4:
                        turno = st.selectbox("Turno:", ["Mañana", "Tarde", "Vespertino"])

                    st.markdown("---")
                    st.markdown("##### 👨‍👦 2. Datos del Padre y Madre")
                    cp1, cp2, cp3, cp4 = st.columns(4)
                    with cp1:
                        apellido_padre = st.text_input("Apellido_Padre:")
                        apellido_madre = st.text_input("Apellido_Madre:")
                    with cp2:
                        nombre_padre = st.text_input("Nombre_Padre:")
                        nombre_madre = st.text_input("Nombre_Madre:")
                    with cp3:
                        dni_padre = st.text_input("DNI_Padre:")
                        dni_madre = st.text_input("DNI_Madre:")
                    with cp4:
                        telefono_padre = st.text_input("Teléfono_Padre:")
                        telefono_madre = st.text_input("Teléfono_Madre:")

                    st.markdown("---")
                    st.markdown("##### 👩‍👦 3. Datos del Tutor Responsable y Contacto")
                    ct1, ct2, ct3, ct4 = st.columns(4)
                    with ct1:
                        apellido_tutor = st.text_input("Apellido_Tutor:")
                        tutor_responsab = st.text_input("Tutor_Responsab:", placeholder="Ej. Madre, Padre, Tío/a")
                    with ct2:
                        nombre_tutor = st.text_input("Nombre_Tutor:")
                        telefono_contacto = st.text_input("Telefono_Contacto (WhatsApp):", placeholder="Ej. +5491112345678")
                    with ct3:
                        dni_tutor = st.text_input("DNI_Tutor:")
                    with ct4:
                        telefono_tutor = st.text_input("Teléfono_Tutor:")

                    st.markdown("---")
                    st.markdown("##### 📝 4. Observaciones")
                    observaciones = st.text_area("Observaciones:", placeholder="Observaciones adicionales del alumno...")

                    btn_guardar = st.form_submit_button("💾 Guardar Alumno en el Excel")

                if btn_guardar:
                    if apellido.strip() and nombre.strip():
                        id_nuevo = f"AL-00{len(st.session_state['alumnos_df']) + 1}"

                        nuevo_registro = {
                            "ID_Alumno": id_nuevo,
                            "Apellido": apellido.strip(),
                            "Nombre": nombre.strip(),
                            "DNI": dni.strip(),
                            "Fecha_Nac.": fecha_nac.strftime('%d/%m/%Y'),
                            "Año": int(anio),
                            "División": division,
                            "Turno": turno,
                            "Dirección": direccion.strip(),
                            "Teléfono_Alumno": telefono_alumno.strip(),
                            "Email_Alumno": email_alumno.strip(),
                            "Apellido_Padre": apellido_padre.strip(),
                            "Nombre_Padre": nombre_padre.strip(),
                            "DNI_Padre": dni_padre.strip(),
                            "Teléfono_Padre": telefono_padre.strip(),
                            "Apellido_Madre": apellido_madre.strip(),
                            "Nombre_Madre": nombre_madre.strip(),
                            "DNI_Madre": dni_madre.strip(),
                            "Teléfono_Madre": telefono_madre.strip(),
                            "Apellido_Tutor": apellido_tutor.strip(),
                            "Nombre_Tutor": nombre_tutor.strip(),
                            "DNI_Tutor": dni_tutor.strip(),
                            "Teléfono_Tutor": telefono_tutor.strip(),
                            "Tutor_Responsab": tutor_responsab.strip(),
                            "Telefono_Contacto": telefono_contacto.strip(),
                            "Observaciones": observaciones.strip()
                        }

                        # 1. Agregar a alumnos_df
                        st.session_state["alumnos_df"] = pd.concat([st.session_state["alumnos_df"], pd.DataFrame([nuevo_registro])], ignore_index=True)

                        # 2. Sincronizar en Asistencia
                        nueva_asistencia = {
                            "ID_Alumno": id_nuevo,
                            "Alumno": f"{apellido.strip()}, {nombre.strip()}",
                            "Año": int(anio),
                            "División": division,
                            "Estado": "Presente",
                            "Observación_Preceptor": "Nuevo Ingreso"
                        }
                        st.session_state["asistencia_df"] = pd.concat([st.session_state["asistencia_df"], pd.DataFrame([nueva_asistencia])], ignore_index=True)

                        # 3. Guardar archivo en disco
                        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                            st.session_state["alumnos_df"].to_excel(writer, sheet_name='Alumnos', index=False)
                            st.session_state["notas_df"].to_excel(writer, sheet_name='Notas', index=False)
                            st.session_state["asistencia_df"].to_excel(writer, sheet_name='Asistencia', index=False)
                            profesores_df.to_excel(writer, sheet_name='Profesores', index=False)
                            plan_df.to_excel(writer, sheet_name='Plan_Materias', index=False)

                        st.success(f"✅ ¡Alumno **{apellido.strip()}, {nombre.strip()}** guardado permanentemente en el Excel!")
                        st.rerun()
                    else:
                        st.error("⚠️ Ingrese al menos el Apellido y Nombre del alumno para realizar el registro.")

            st.markdown("<br>", unsafe_allow_html=True)
            
            # Filtros
            col_f1, col_f2 = st.columns([1, 2])
            with col_f1:
                cursos_unicos = (
                    st.session_state["alumnos_df"]['Año'].astype(str).str.strip() + "°" + 
                    st.session_state["alumnos_df"]['División'].astype(str).str.strip()
                ).unique()
                cursos_disponibles = ["Todos"] + sorted(list(cursos_unicos))
                curso_filtro = st.selectbox("Filtrar por Curso:", cursos_disponibles)

            with col_f2:
                buscar_alumno = st.text_input("Buscar por Nombre, Apellido o DNI:", placeholder="Ej. Pérez, 45123890...")

            alumnos_vista = st.session_state["alumnos_df"].copy()
            if curso_filtro != "Todos":
                partes = curso_filtro.split("°")
                anio_f = int(partes[0])
                div_f = partes[1]
                alumnos_vista = alumnos_vista[(alumnos_vista['Año'] == anio_f) & (alumnos_vista['División'] == div_f)]
            
            if buscar_alumno:
                alumnos_vista = alumnos_vista[
                    alumnos_vista['Nombre'].astype(str).str.contains(buscar_alumno, case=False, na=False) |
                    alumnos_vista['Apellido'].astype(str).str.contains(buscar_alumno, case=False, na=False) |
                    alumnos_vista['DNI'].astype(str).str.contains(buscar_alumno, case=False, na=False)
                ]

            st.dataframe(alumnos_vista, use_container_width=True)

            # Exportación manual
            st.markdown("---")
            st.markdown("##### 📥 Exportar Registro de Alumnos en Excel")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state["alumnos_df"].to_excel(writer, sheet_name='Alumnos', index=False)
                st.session_state["notas_df"].to_excel(writer, sheet_name='Notas', index=False)
                st.session_state["asistencia_df"].to_excel(writer, sheet_name='Asistencia', index=False)
            
            st.download_button(
                label="📥 Descargar Planilla Excel Actualizada (.xlsx)",
                data=buffer.getvalue(),
                file_name="Gestion_Escolar_Secundaria_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab_notas:
            st.subheader("📊 Calificaciones de Todas las Materias y Cursos")
            st.dataframe(st.session_state["notas_df"], use_container_width=True)

        with tab_profesores:
            st.subheader("👩‍🏫 Nómina de Profesores y Materias Asignadas")
            st.dataframe(profesores_df, use_container_width=True)

    # --- VISTA DOCENTE ---
    elif user_info['rol'] == 'Docente':
        prof_data = profesores_df[profesores_df['ID_Profesor'] == user_info['id']].iloc[0]
        st.title(f"📚 Gestión Académica — Prof. {prof_data['Nombre']} {prof_data['Apellido']}")
        st.markdown(f"**Materia Asignada:** `{prof_data['Materia_Principal']}`")

        tab_cargar_notas, tab_enviar_informe = st.tabs(["📝 Cargar/Modificar Notas", "📲 Enviar Informe Trimestral WhatsApp"])

        notas_prof = st.session_state["notas_df"][st.session_state["notas_df"]['ID_Profesor'] == user_info['id']]
        cursos = notas_prof[['Año', 'División']].drop_duplicates()
        opciones_cursos = [f"{row['Año']}° {row['División']}" for _, row in cursos.iterrows()]

        with tab_cargar_notas:
            curso_seleccionado = st.selectbox("Seleccionar Curso:", opciones_cursos, key="doc_curso")
            trimestre_trabajo = st.radio("Seleccionar Trimestre a Evaluar:", ["1° Trimestre", "2° Trimestre", "3° Trimestre"], horizontal=True)

            anio_sel = int(curso_seleccionado.split("°")[0])
            div_sel = curso_seleccionado.split(" ")[1]

            notas_curso = notas_prof[(notas_prof['Año'] == anio_sel) & (notas_prof['División'] == div_sel)].copy()

            st.subheader(f"Planilla de Evaluación: {prof_data['Materia_Principal']} — {curso_seleccionado}")
            st.info("💡 Completa las 6 notas individuales. La nota del trimestre y el Promedio Final se recalcularán automáticamente.")

            # Mapeo según el trimestre seleccionado
            if trimestre_trabajo == "1° Trimestre":
                cols_sub = ['1T_Prueba1', '1T_Oral1', '1T_Prueba2', '1T_Oral2', '1T_Participacion', '1T_Carpeta']
                col_trim_res = 'Nota_1er_Trim.'
            elif trimestre_trabajo == "2° Trimestre":
                cols_sub = ['2T_Prueba1', '2T_Oral1', '2T_Prueba2', '2T_Oral2', '2T_Participacion', '2T_Carpeta']
                col_trim_res = 'Nota_2do_Trim.'
            else:
                cols_sub = ['3T_Prueba1', '3T_Oral1', '3T_Prueba2', '3T_Oral2', '3T_Participacion', '3T_Carpeta']
                col_trim_res = 'Nota_3er_Trim.'

            columnas_mostrar = ['ID_Alumno', 'Alumno'] + cols_sub + [col_trim_res, 'Promedio', 'Condición']

            edited_df = st.data_editor(
                notas_curso[columnas_mostrar],
                column_config={
                    cols_sub[0]: st.column_config.NumberColumn("Prueba 1", min_value=1, max_value=10, format="%d"),
                    cols_sub[1]: st.column_config.NumberColumn("L. Oral 1", min_value=1, max_value=10, format="%d"),
                    cols_sub[2]: st.column_config.NumberColumn("Prueba 2", min_value=1, max_value=10, format="%d"),
                    cols_sub[3]: st.column_config.NumberColumn("L. Oral 2", min_value=1, max_value=10, format="%d"),
                    cols_sub[4]: st.column_config.NumberColumn("Participación", min_value=1, max_value=10, format="%d"),
                    cols_sub[5]: st.column_config.NumberColumn("Carpeta", min_value=1, max_value=10, format="%d"),
                    col_trim_res: st.column_config.NumberColumn("Prom. Trimestre", format="%.2f"),
                    "Promedio": st.column_config.NumberColumn("Prom. Final", format="%.2f"),
                    "Condición": st.column_config.TextColumn("Condición")
                },
                disabled=['ID_Alumno', 'Alumno', col_trim_res, 'Promedio', 'Condición'],
                use_container_width=True
            )

            if st.button("💾 Recalcular y Guardar en Excel"):
                # 1. Recalcular promedio del trimestre activo
                edited_df[col_trim_res] = edited_df[cols_sub].mean(axis=1).round(2)

                # 2. Recalcular Promedio Final
                edited_df['Promedio'] = (
                    edited_df['Nota_1er_Trim.'] + edited_df['Nota_2do_Trim.'] + edited_df['Nota_3er_Trim.']
                ) / 3
                edited_df['Promedio'] = edited_df['Promedio'].round(2)
                edited_df['Condición'] = edited_df['Promedio'].apply(lambda x: 'Aprobado' if x >= 6 else 'Desaprobado')

                # 3. Actualizar memory state
                st.session_state["notas_df"].update(edited_df)

                # 4. Sobrescribir archivo Excel
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                    st.session_state["alumnos_df"].to_excel(writer, sheet_name='Alumnos', index=False)
                    st.session_state["notas_df"].to_excel(writer, sheet_name='Notas', index=False)
                    st.session_state["asistencia_df"].to_excel(writer, sheet_name='Asistencia', index=False)
                    profesores_df.to_excel(writer, sheet_name='Profesores', index=False)
                    plan_df.to_excel(writer, sheet_name='Plan_Materias', index=False)

                st.success("✅ ¡Notas y promedios actualizados exitosamente en la planilla Excel!")
                st.rerun()

        with tab_enviar_informe:
            st.subheader("📲 Notificación de Informe Trimestral al Tutor")
            
            alumno_sel = st.selectbox("Seleccionar Alumno para Notificar:", notas_prof['Alumno'].unique())
            trimestre_sel = st.selectbox("Seleccionar Trimestre a Enviar:", ["1° Trimestre", "2° Trimestre", "3° Trimestre"])
            
            row_nota = notas_prof[notas_prof['Alumno'] == alumno_sel].iloc[0]
            row_alumno_info = st.session_state["alumnos_df"][st.session_state["alumnos_df"]['ID_Alumno'] == row_nota['ID_Alumno']]

            if not row_alumno_info.empty:
                row_alumno_info = row_alumno_info.iloc[0]
                tutor_nom = f"{row_alumno_info.get('Nombre_Tutor', '')} {row_alumno_info.get('Apellido_Tutor', '')}".strip() or row_alumno_info.get('Tutor_Responsab', 'Tutor')
                tel_contacto = row_alumno_info.get('Telefono_Contacto', 'No registrado')

                col_i1, col_i2 = st.columns(2)
                with col_i1:
                    st.info(f"👤 **Alumno:** {row_nota['Alumno']}\n\n📘 **Materia:** {prof_data['Materia_Principal']}")
                with col_i2:
                    st.info(f"👩‍👦 **Tutor:** {tutor_nom}\n\n📞 **Teléfono:** {tel_contacto}")

                col_t = {"1° Trimestre": 'Nota_1er_Trim.', "2° Trimestre": 'Nota_2do_Trim.', "3° Trimestre": 'Nota_3er_Trim.'}
                nota_val = row_nota[col_t[trimestre_sel]]

                msg_docente = (
                    f"Estimado/a {tutor_nom}, le informamos desde el Portal Escolar "
                    f"la calificación correspondiente al {trimestre_sel} del estudiante {row_nota['Alumno']} "
                    f"en la asignatura {prof_data['Materia_Principal']}: *{nota_val}/10*. "
                    f"Observaciones: {row_nota.get('Observación', 'Sin observaciones')}. "
                    f"Profesor/a: {prof_data['Nombre']} {prof_data['Apellido']}."
                )

                st.text_area("Mensaje a enviar:", value=msg_docente, height=120)

                if tel_contacto and str(tel_contacto) != 'nan':
                    url_wa_doc = generar_link_whatsapp(tel_contacto, msg_docente)
                    st.markdown(f'<a href="{url_wa_doc}" target="_blank"><button style="background-color:#25D366; color:white; padding:10px 20px; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📲 Abrir WhatsApp y Enviar Informe</button></a>', unsafe_allow_html=True)
                else:
                    st.warning("⚠️ El alumno no tiene un número de teléfono registrado en el Excel.")

    # --- VISTA PRECEPTORÍA ---
    elif user_info['rol'] == 'Preceptor':
        st.title("📋 Control Diario de Asistencia y Avisos a Tutores")
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            curso_p = st.selectbox("Seleccionar Curso:", ["1°A", "2°A", "3°A"])
        with col_c2:
            fecha_asistencia = st.date_input("Fecha de Asistencia:", datetime.now())

        anio_p = int(curso_p[0])
        div_p = curso_p[2]
        
        asistencia_filtrada = st.session_state["asistencia_df"][
            (st.session_state["asistencia_df"]['Año'] == anio_p) & 
            (st.session_state["asistencia_df"]['División'] == div_p)
        ].copy()

        st.subheader(f"Registro de Asistencia - Curso {curso_p} ({fecha_asistencia.strftime('%d/%m/%Y')})")
        
        edited_asistencia = st.data_editor(
            asistencia_filtrada[['ID_Alumno', 'Alumno', 'Estado', 'Observación_Preceptor']],
            column_config={
                "Estado": st.column_config.SelectboxColumn(
                    "Estado de Asistencia",
                    options=["Presente", "Ausente", "Tarde", "Justificada"],
                    required=True
                ),
                "Observación_Preceptor": st.column_config.TextColumn("Observación Preceptor")
            },
            disabled=['ID_Alumno', 'Alumno'],
            use_container_width=True
        )

        if st.button("💾 Guardar Asistencia en Excel"):
            st.session_state["asistencia_df"].update(edited_asistencia)
            with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                st.session_state["alumnos_df"].to_excel(writer, sheet_name='Alumnos', index=False)
                st.session_state["notas_df"].to_excel(writer, sheet_name='Notas', index=False)
                st.session_state["asistencia_df"].to_excel(writer, sheet_name='Asistencia', index=False)
                profesores_df.to_excel(writer, sheet_name='Profesores', index=False)
                plan_df.to_excel(writer, sheet_name='Plan_Materias', index=False)

            st.success("✅ Asistencia guardada en el Excel.")

        st.markdown("---")
        st.subheader("📲 Envío de Notificaciones de Ausencia por WhatsApp")
        
        ausentes = edited_asistencia[edited_asistencia['Estado'] == 'Ausente']

        if not ausentes.empty:
            st.warning(f"Se registraron **{len(ausentes)}** alumno(s) ausente(s) el día {fecha_asistencia.strftime('%d/%m/%Y')}:")
            
            for _, row_ausente in ausentes.iterrows():
                info_alumno = st.session_state["alumnos_df"][st.session_state["alumnos_df"]['ID_Alumno'] == row_ausente['ID_Alumno']]
                
                if not info_alumno.empty:
                    info_alumno = info_alumno.iloc[0]
                    tutor_nom = f"{info_alumno.get('Nombre_Tutor', '')} {info_alumno.get('Apellido_Tutor', '')}".strip() or info_alumno.get('Tutor_Responsab', 'Tutor/a')
                    tel_contacto = info_alumno.get('Telefono_Contacto', '')
                    
                    msg_ausencia = (
                        f"Estimado/a {tutor_nom}, le notificamos desde la Preceptoría de la Escuela "
                        f"que el estudiante *{row_ausente['Alumno']}* registra una inasistencia (AUSENTE) "
                        f"el día de la fecha ({fecha_asistencia.strftime('%d/%m/%Y')}). "
                        f"Por favor, comuníquese con el establecimiento para justificar la falta."
                    )
                    
                    col_a1, col_a2 = st.columns([3, 1])
                    with col_a1:
                        st.write(f"👤 **{row_ausente['Alumno']}** — Tutor: *{tutor_nom}* ({tel_contacto})")
                    with col_a2:
                        if pd.notna(tel_contacto) and str(tel_contacto) != '':
                            link_wa = generar_link_whatsapp(tel_contacto, msg_ausencia)
                            st.markdown(f'<a href="{link_wa}" target="_blank"><button style="background-color:#25D366; color:white; padding:6px 12px; border:none; border-radius:6px; font-weight:bold; cursor:pointer;">📲 Enviar WhatsApp</button></a>', unsafe_allow_html=True)
                        else:
                            st.caption("⚠️ Sin teléfono")
        else:
            st.success("🎉 No hay alumnos ausentes registrados en este curso.")
