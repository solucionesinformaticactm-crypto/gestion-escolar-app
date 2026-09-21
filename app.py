import streamlit as st
import pandas as pd
from datetime import datetime, date
import urllib.parse
import io

# Configuración de página y estética Gris / Violeta
st.set_page_config(
    page_title="Sistema de Gestión Escolar Integral",
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

# Manejo de Estado de Sesión
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_actual" not in st.session_state:
    st.session_state["usuario_actual"] = None

def generar_link_whatsapp(numero, mensaje):
    num_limpio = ''.join(filter(str.isdigit, str(numero)))
    mensaje_codificado = urllib.parse.quote(mensaje)
    return f"https://wa.me/{num_limpio}?text={mensaje_codificado}"

# Función generadora del Boletín de Calificaciones
def generar_html_boletin(alumno_id):
    alumno_info = st.session_state["alumnos_df"][st.session_state["alumnos_df"]['ID_Alumno'] == alumno_id]
    if alumno_info.empty:
        return "<p>Alumno no encontrado</p>"
    
    alumno = alumno_info.iloc[0]
    notas_alumno = st.session_state["notas_df"][st.session_state["notas_df"]['ID_Alumno'] == alumno_id]
    
    filas = ""
    for _, row in notas_alumno.iterrows():
        n1 = f"{row['Nota_1er_Trim.']:.2f}" if row['Nota_1er_Trim.'] > 0 else "-"
        n2 = f"{row['Nota_2do_Trim.']:.2f}" if row['Nota_2do_Trim.'] > 0 else "-"
        n3 = f"{row['Nota_3er_Trim.']:.2f}" if row['Nota_3er_Trim.'] > 0 else "-"
        prom = f"{row['Promedio']:.2f}" if row['Promedio'] > 0 else "-"
        cond = row['Condición'] if row['Promedio'] > 0 else "En Cursada"
        color_cond = "#196F3D" if cond == "Aprobado" else "#922B21"
        
        filas += f"""
        <tr>
            <td style="text-align: left; font-weight: bold; padding: 6px; border: 1px solid #ddd;">{row.get('Materia', 'Asignatura')}</td>
            <td style="padding: 6px; border: 1px solid #ddd;">{n1}</td>
            <td style="padding: 6px; border: 1px solid #ddd;">{n2}</td>
            <td style="padding: 6px; border: 1px solid #ddd;">{n3}</td>
            <td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">{prom}</td>
            <td style="padding: 6px; border: 1px solid #ddd; color: {color_cond}; font-weight: bold;">{cond}</td>
        </tr>
        """

    html_content = f"""
    <div style="background-color: white; padding: 20px; border-radius: 8px; font-family: Arial, sans-serif; border: 2px solid #6C3483;">
        <div style="text-align: center; border-bottom: 2px solid #6C3483; padding-bottom: 10px; margin-bottom: 15px;">
            <h2 style="color: #512E5F; margin: 0;">BOLETÍN OFICIAL DE CALIFICACIONES</h2>
            <h4 style="color: #555; margin: 5px 0 0 0;">Ciclo Lectivo 2026 — Nivel Secundario</h4>
        </div>
        
        <table style="width: 100%; margin-bottom: 15px; font-size: 13px;">
            <tr>
                <td><strong>Estudiante:</strong> {alumno['Apellido']}, {alumno['Nombre']}</td>
                <td><strong>DNI:</strong> {alumno.get('DNI', '-')}</td>
            </tr>
            <tr>
                <td><strong>Curso:</strong> {alumno['Año']}° "{alumno['División']}" ({alumno.get('Turno', 'Mañana')})</td>
                <td><strong>Tutor:</strong> {alumno.get('Nombre_Tutor', '')} {alumno.get('Apellido_Tutor', '')}</td>
            </tr>
        </table>

        <table style="width: 100%; border-collapse: collapse; text-align: center; font-size: 12px;">
            <thead>
                <tr style="background-color: #6C3483; color: white;">
                    <th style="padding: 8px; text-align: left;">Asignatura</th>
                    <th style="padding: 8px;">1° Trim.</th>
                    <th style="padding: 8px;">2° Trim.</th>
                    <th style="padding: 8px;">3° Trim.</th>
                    <th style="padding: 8px;">Prom. Final</th>
                    <th style="padding: 8px;">Condición</th>
                </tr>
            </thead>
            <tbody>
                {filas}
            </tbody>
        </table>
        
        <div style="margin-top: 30px; display: flex; justify-content: space-around; text-align: center; font-size: 11px; color: #555;">
            <div>_______________________<br>Firma Tutor</div>
            <div>_______________________<br>Sello y Firma Dirección</div>
        </div>
    </div>
    """
    return html_content

@st.cache_data
def cargar_datos_iniciales():
    alumnos = pd.read_excel(EXCEL_FILE, sheet_name='Alumnos')
    profesores = pd.read_excel(EXCEL_FILE, sheet_name='Profesores')
    plan = pd.read_excel(EXCEL_FILE, sheet_name='Plan_Materias')
    notas = pd.read_excel(EXCEL_FILE, sheet_name='Notas')
    asistencia = pd.read_excel(EXCEL_FILE, sheet_name='Asistencia')
    
    try:
        conducta = pd.read_excel(EXCEL_FILE, sheet_name='Conducta')
    except Exception:
        conducta = pd.DataFrame(columns=[
            "ID_Parte", "ID_Alumno", "Alumno", "Año", "División", 
            "Fecha", "Tipo_Evento", "Motivo_Detalle", "Registrado_Por"
        ])

    try:
        agenda = pd.read_excel(EXCEL_FILE, sheet_name='Agenda')
    except Exception:
        agenda = pd.DataFrame(columns=[
            "ID_Evento", "Fecha", "Año", "División", "Materia", 
            "Tipo_Evento", "Título_Descripción", "Publicado_Por"
        ])

    try:
        cuotas = pd.read_excel(EXCEL_FILE, sheet_name='Cuotas')
    except Exception:
        cuotas = pd.DataFrame(columns=[
            "ID_Pago", "ID_Alumno", "Alumno", "Año", "División", 
            "Mes_Cuota", "Monto", "Estado_Pago", "Fecha_Pago"
        ])
        
    return alumnos, profesores, plan, notas, asistencia, conducta, agenda, cuotas

alumnos_init, profesores_df, plan_df, notas_init, asistencia_init, conducta_init, agenda_init, cuotas_init = cargar_datos_iniciales()

if "alumnos_df" not in st.session_state:
    st.session_state["alumnos_df"] = alumnos_init.copy()
if "notas_df" not in st.session_state:
    st.session_state["notas_df"] = notas_init.copy()
if "asistencia_df" not in st.session_state:
    st.session_state["asistencia_df"] = asistencia_init.copy()
if "conducta_df" not in st.session_state:
    st.session_state["conducta_df"] = conducta_init.copy()
if "agenda_df" not in st.session_state:
    st.session_state["agenda_df"] = agenda_init.copy()
if "cuotas_df" not in st.session_state:
    st.session_state["cuotas_df"] = cuotas_init.copy()

cols_eval = [
    '1T_Prueba1', '1T_Oral1', '1T_Prueba2', '1T_Oral2', '1T_Participacion', '1T_Carpeta',
    '2T_Prueba1', '2T_Oral1', '2T_Prueba2', '2T_Oral2', '2T_Participacion', '2T_Carpeta',
    '3T_Prueba1', '3T_Oral1', '3T_Prueba2', '3T_Oral2', '3T_Participacion', '3T_Carpeta'
]

columnas_numericas_notas = cols_eval + ['Nota_1er_Trim.', 'Nota_2do_Trim.', 'Nota_3er_Trim.', 'Promedio']
for col in columnas_numericas_notas:
    if col not in st.session_state["notas_df"].columns:
        st.session_state["notas_df"][col] = 0.0
    st.session_state["notas_df"][col] = pd.to_numeric(st.session_state["notas_df"][col], errors='coerce').fillna(0.0).astype(float)

# Función auxiliar para guardar todas las pestañas de Excel
def guardar_excel_completo():
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        st.session_state["alumnos_df"].to_excel(writer, sheet_name='Alumnos', index=False)
        st.session_state["notas_df"].to_excel(writer, sheet_name='Notas', index=False)
        st.session_state["asistencia_df"].to_excel(writer, sheet_name='Asistencia', index=False)
        st.session_state["conducta_df"].to_excel(writer, sheet_name='Conducta', index=False)
        st.session_state["agenda_df"].to_excel(writer, sheet_name='Agenda', index=False)
        st.session_state["cuotas_df"].to_excel(writer, sheet_name='Cuotas', index=False)
        profesores_df.to_excel(writer, sheet_name='Profesores', index=False)
        plan_df.to_excel(writer, sheet_name='Plan_Materias', index=False)

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
# 🏫 SISTEMA PRINCIPAL
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
        st.write("Visión general del rendimiento escolar, asistencia, partes disciplinarios, agenda y cobranzas.")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.markdown(f'<div class="metric-card"><h4>Total Alumnos</h4><h2>{len(st.session_state["alumnos_df"])}</h2></div>', unsafe_allow_html=True)
        with col2:
            st.markdown(f'<div class="metric-card"><h4>Total Profesores</h4><h2>{len(profesores_df)}</h2></div>', unsafe_allow_html=True)
        with col3:
            prom_gen = round(st.session_state["notas_df"]['Promedio'].mean(), 2) if not st.session_state["notas_df"].empty else 0
            st.markdown(f'<div class="metric-card"><h4>Promedio General</h4><h2>{prom_gen}</h2></div>', unsafe_allow_html=True)
        with col4:
            st.markdown(f'<div class="metric-card"><h4>Partes Conducta</h4><h2>{len(st.session_state["conducta_df"])}</h2></div>', unsafe_allow_html=True)
        with col5:
            rec_cuotas = st.session_state["cuotas_df"][st.session_state["cuotas_df"]['Estado_Pago'] == 'Pagado']['Monto'].sum() if not st.session_state["cuotas_df"].empty else 0
            st.markdown(f'<div class="metric-card"><h4>Cobranza Cuotas</h4><h2>${rec_cuotas:,.0f}</h2></div>', unsafe_allow_html=True)

        st.markdown("---")

        tab_alumnos, tab_notas, tab_profesores, tab_boletin, tab_conducta_dir, tab_graficos, tab_agenda_dir, tab_cuotas = st.tabs([
            "👨‍🎓 Padrón de Alumnos", 
            "📊 Calificaciones", 
            "👩‍🏫 Planta Docente",
            "📄 Boletín Oficial",
            "📋 Conducta",
            "📈 Tableros Gráficos",
            "📆 Agenda Escolar",
            "💳 Aranceles y Cuotas"
        ])

        with tab_alumnos:
            st.subheader("👨‍🎓 Registro General de Alumnos y Legajos")

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

                        st.session_state["alumnos_df"] = pd.concat([st.session_state["alumnos_df"], pd.DataFrame([nuevo_registro])], ignore_index=True)

                        nueva_asistencia = {
                            "ID_Alumno": id_nuevo,
                            "Alumno": f"{apellido.strip()}, {nombre.strip()}",
                            "Año": int(anio),
                            "División": division,
                            "Estado": "Presente",
                            "Observación_Preceptor": "Nuevo Ingreso"
                        }
                        st.session_state["asistencia_df"] = pd.concat([st.session_state["asistencia_df"], pd.DataFrame([nueva_asistencia])], ignore_index=True)

                        guardar_excel_completo()
                        st.success(f"✅ ¡Alumno **{apellido.strip()}, {nombre.strip()}** guardado permanentemente en el Excel!")
                        st.rerun()
                    else:
                        st.error("⚠️ Ingrese al menos el Apellido y Nombre del alumno para realizar el registro.")

            st.markdown("<br>", unsafe_allow_html=True)
            
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

            st.markdown("---")
            st.markdown("##### 📥 Exportar Registro de Alumnos en Excel")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state["alumnos_df"].to_excel(writer, sheet_name='Alumnos', index=False)
                st.session_state["notas_df"].to_excel(writer, sheet_name='Notas', index=False)
                st.session_state["asistencia_df"].to_excel(writer, sheet_name='Asistencia', index=False)
                st.session_state["conducta_df"].to_excel(writer, sheet_name='Conducta', index=False)
                st.session_state["agenda_df"].to_excel(writer, sheet_name='Agenda', index=False)
                st.session_state["cuotas_df"].to_excel(writer, sheet_name='Cuotas', index=False)
            
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

        with tab_boletin:
            st.subheader("📄 Generación e Impresión del Boletín Oficial")
            
            col_b1, col_b2 = st.columns([1, 2])
            with col_b1:
                curso_b = st.selectbox("Seleccionar Curso para Generar Boletín:", sorted(list(cursos_unicos)), key="bol_curso")
                partes_b = curso_b.split("°")
                anio_b, div_b = int(partes_b[0]), partes_b[1]
                
                alumnos_curso_b = st.session_state["alumnos_df"][
                    (st.session_state["alumnos_df"]['Año'] == anio_b) & 
                    (st.session_state["alumnos_df"]['División'] == div_b)
                ]
                
                if not alumnos_curso_b.empty:
                    opciones_alumnos_b = {
                        f"{r['Apellido']}, {r['Nombre']} (DNI: {r.get('DNI', '-')})": r['ID_Alumno'] 
                        for _, r in alumnos_curso_b.iterrows()
                    }
                    alumno_sel_nom = st.selectbox("Seleccionar Estudiante:", list(opciones_alumnos_b.keys()))
                    id_alumno_sel = opciones_alumnos_b[alumno_sel_nom]
                else:
                    id_alumno_sel = None
                    st.warning("No hay alumnos registrados en este curso.")

            with col_b2:
                if id_alumno_sel:
                    boletin_html = generar_html_boletin(id_alumno_sel)
                    st.components.v1.html(boletin_html, height=450, scrolling=True)

        with tab_conducta_dir:
            st.subheader("📋 Consolidado de Partes Disciplinarios e Infracciones")
            st.dataframe(st.session_state["conducta_df"], use_container_width=True)

        # -----------------------------------------------------------------------------
        # OPTION 1: 📈 TABLEROS GRÁFICOS Y ESTADÍSTICAS AVANZADAS (DIRECTIVOS)
        # -----------------------------------------------------------------------------
        with tab_graficos:
            st.subheader("📈 Indicadores Institucionales y Estadísticas Avanzadas")

            g_col1, g_col2 = st.columns(2)

            with g_col1:
                st.markdown("##### 📊 Promedio General por Materia")
                if not st.session_state["notas_df"].empty and 'Materia' in st.session_state["notas_df"].columns:
                    promedios_materia = st.session_state["notas_df"].groupby('Materia')['Promedio'].mean().round(2)
                    st.bar_chart(promedios_materia)
                else:
                    st.info("Sin datos suficientes de calificaciones para graficar.")

            with g_col2:
                st.markdown("##### 📋 Distribución de Inasistencias por Estado")
                if not st.session_state["asistencia_df"].empty and 'Estado' in st.session_state["asistencia_df"].columns:
                    conteo_asistencia = st.session_state["asistencia_df"]['Estado'].value_counts()
                    st.bar_chart(conteo_asistencia)
                else:
                    st.info("Sin datos de asistencia registrados.")

            st.markdown("---")
            g_col3, g_col4 = st.columns(2)

            with g_col3:
                st.markdown("##### 🚨 Partes Disciplinarios por Categoria")
                if not st.session_state["conducta_df"].empty and 'Tipo_Evento' in st.session_state["conducta_df"].columns:
                    conteo_conducta = st.session_state["conducta_df"]['Tipo_Evento'].value_counts()
                    st.bar_chart(conteo_conducta)
                else:
                    st.info("Sin partes de conducta registrados.")

            with g_col4:
                st.markdown("##### 💳 Estado de Cobranza de Aranceles")
                if not st.session_state["cuotas_df"].empty and 'Estado_Pago' in st.session_state["cuotas_df"].columns:
                    conteo_cuotas = st.session_state["cuotas_df"]['Estado_Pago'].value_counts()
                    st.bar_chart(conteo_cuotas)
                else:
                    st.info("Sin registros de aranceles o cuotas.")

        # -----------------------------------------------------------------------------
        # OPTION 2: 📆 AGENDA ESCOLAR Y CALENDARIO DE EVALUACIONES
        # -----------------------------------------------------------------------------
        with tab_agenda_dir:
            st.subheader("📆 Cronograma General de Exámenes y Eventos Institucionales")

            st.dataframe(st.session_state["agenda_df"], use_container_width=True)

        # -----------------------------------------------------------------------------
        # OPTION 3: 💳 MÓDULO DE ARANCELES, CUOTAS Y COBRANZAS
        # -----------------------------------------------------------------------------
        with tab_cuotas:
            st.subheader("💳 Gestión de Aranceles, Cuotas y Cooperadora")

            with st.expander("➕ Registrar Nuevo Cobro de Cuota / Arancel", expanded=False):
                with st.form("form_nuevo_cobro"):
                    c_col1, c_col2 = st.columns(2)
                    with c_col1:
                        curso_cuota_sel = st.selectbox("Seleccionar Curso:", sorted(list(cursos_unicos)), key="cuota_curso")
                        partes_cuota = curso_cuota_sel.split("°")
                        anio_cuota, div_cuota = int(partes_cuota[0]), partes_cuota[1]

                        alumnos_cuota = st.session_state["alumnos_df"][
                            (st.session_state["alumnos_df"]['Año'] == anio_cuota) & 
                            (st.session_state["alumnos_df"]['División'] == div_cuota)
                        ]
                        
                        if not alumnos_cuota.empty:
                            mapa_alumnos_cuota = {
                                f"{r['Apellido']}, {r['Nombre']}": r['ID_Alumno'] 
                                for _, r in alumnos_cuota.iterrows()
                            }
                            alumno_cuota_nom = st.selectbox("Seleccionar Alumno:", list(mapa_alumnos_cuota.keys()))
                            id_alumno_cuota = mapa_alumnos_cuota[alumno_cuota_nom]
                        else:
                            id_alumno_cuota = None

                    with c_col2:
                        mes_cuota = st.selectbox("Mes de Cuota:", ["Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", "Matrícula"])
                        monto_cuota = st.number_input("Monto ($):", min_value=0.0, value=15000.0, step=1000.0)
                        estado_pago = st.selectbox("Estado del Pago:", ["Pagado", "Pendiente", "Atrasado"])
                        fecha_cobro = st.date_input("Fecha del Pago:", datetime.now())

                    btn_guardar_cuota = st.form_submit_button("💾 Registrar Cobro")

                if btn_guardar_cuota and id_alumno_cuota:
                    id_pago = f"PAG-{len(st.session_state['cuotas_df']) + 1:03d}"
                    nuevo_pago = {
                        "ID_Pago": id_pago,
                        "ID_Alumno": id_alumno_cuota,
                        "Alumno": alumno_cuota_nom,
                        "Año": anio_cuota,
                        "División": div_cuota,
                        "Mes_Cuota": mes_cuota,
                        "Monto": monto_cuota,
                        "Estado_Pago": estado_pago,
                        "Fecha_Pago": fecha_cobro.strftime('%d/%m/%Y')
                    }
                    st.session_state["cuotas_df"] = pd.concat([
                        st.session_state["cuotas_df"], 
                        pd.DataFrame([nuevo_pago])
                    ], ignore_index=True)

                    guardar_excel_completo()
                    st.success(f"✅ ¡Cobro de {mes_cuota} registrado exitosamente!")
                    st.rerun()

            st.markdown("---")
            st.subheader("📋 Estado de Pagos y Notificación a Tutores")
            st.dataframe(st.session_state["cuotas_df"], use_container_width=True)

            if not st.session_state["cuotas_df"].empty:
                pagos_pendientes = st.session_state["cuotas_df"][st.session_state["cuotas_df"]['Estado_Pago'] != 'Pagado']
                if not pagos_pendientes.empty:
                    pago_sel_id = st.selectbox("Seleccionar Pago para Notificar Recordatorio:", pagos_pendientes['ID_Pago'].tolist())
                    row_pago = pagos_pendientes[pagos_pendientes['ID_Pago'] == pago_sel_id].iloc[0]

                    info_al_cuota = st.session_state["alumnos_df"][st.session_state["alumnos_df"]['ID_Alumno'] == row_pago['ID_Alumno']]
                    if not info_al_cuota.empty:
                        info_al_cuota = info_al_cuota.iloc[0]
                        tutor_nom = f"{info_al_cuota.get('Nombre_Tutor', '')} {info_al_cuota.get('Apellido_Tutor', '')}".strip() or info_al_cuota.get('Tutor_Responsab', 'Tutor/a')
                        tel_contacto = info_al_cuota.get('Telefono_Contacto', '')

                        msg_cuota = (
                            f"Estimado/a {tutor_nom}, le enviamos un recordatorio desde la Administración Escolar "
                            f"respecto al arancel correspondiente al mes de *{row_pago['Mes_Cuota']}* "
                            f"del estudiante *{row_pago['Alumno']}* por un monto de *${row_pago['Monto']:,.2f}*.\n\n"
                            f"Estado actual: *{row_pago['Estado_Pago']}*.\n"
                            f"Agradecemos su regularización."
                        )

                        st.text_area("Mensaje de Recordatorio de Pago:", value=msg_cuota, height=120)

                        if pd.notna(tel_contacto) and str(tel_contacto) != '':
                            link_wa_cuota = generar_link_whatsapp(tel_contacto, msg_cuota)
                            st.markdown(f'<a href="{link_wa_cuota}" target="_blank"><button style="background-color:#25D366; color:white; padding:8px 16px; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📲 Enviar Recordatorio por WhatsApp</button></a>', unsafe_allow_html=True)

    # --- VISTA DOCENTE ---
    elif user_info['rol'] == 'Docente':
        prof_data = profesores_df[profesores_df['ID_Profesor'] == user_info['id']].iloc[0]
        st.title(f"📚 Gestión Académica — Prof. {prof_data['Nombre']} {prof_data['Apellido']}")
        st.markdown(f"**Materia Asignada:** `{prof_data['Materia_Principal']}`")

        tab_cargar_notas, tab_enviar_informe, tab_agenda_doc = st.tabs([
            "📝 Cargar/Modificar Notas", 
            "📲 Enviar Informe Trimestral WhatsApp",
            "📆 Cargar Exámenes a Agenda Escolar"
        ])

        notas_prof = st.session_state["notas_df"][st.session_state["notas_df"]['ID_Profesor'] == user_info['id']]
        cursos = notas_prof[['Año', 'División']].drop_duplicates()
        opciones_cursos = [f"{row['Año']}° {row['División']}" for _, row in cursos.iterrows()]

        with tab_cargar_notas:
            curso_seleccionado = st.selectbox("Seleccionar Curso:", opciones_cursos, key="doc_curso")
            trimestre_trabajo = st.radio("Seleccionar Trimestre a Evaluar:", ["1° Trimestre", "2° Trimestre", "3° Trimestre"], horizontal=True)

            anio_sel = int(curso_seleccionado.split("°")[0])
            div_sel = curso_seleccionado.split(" ")[1]

            notas_curso = st.session_state["notas_df"][
                (st.session_state["notas_df"]['ID_Profesor'] == user_info['id']) &
                (st.session_state["notas_df"]['Año'] == anio_sel) &
                (st.session_state["notas_df"]['División'] == div_sel)
            ].copy()

            st.subheader(f"Planilla de Evaluación: {prof_data['Materia_Principal']} — {curso_seleccionado}")
            st.info("💡 Las casillas con '0' se ignoran. El promedio trimestral calcula automáticamente el promedio solo de las notas evaluadas.")

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
                    cols_sub[0]: st.column_config.NumberColumn("Prueba 1", min_value=0, max_value=10, format="%d"),
                    cols_sub[1]: st.column_config.NumberColumn("L. Oral 1", min_value=0, max_value=10, format="%d"),
                    cols_sub[2]: st.column_config.NumberColumn("Prueba 2", min_value=0, max_value=10, format="%d"),
                    cols_sub[3]: st.column_config.NumberColumn("L. Oral 2", min_value=0, max_value=10, format="%d"),
                    cols_sub[4]: st.column_config.NumberColumn("Participación", min_value=0, max_value=10, format="%d"),
                    cols_sub[5]: st.column_config.NumberColumn("Carpeta", min_value=0, max_value=10, format="%d"),
                    col_trim_res: st.column_config.NumberColumn("Prom. Trimestre", format="%.2f"),
                    "Promedio": st.column_config.NumberColumn("Prom. Final", format="%.2f"),
                    "Condición": st.column_config.TextColumn("Condición")
                },
                disabled=['ID_Alumno', 'Alumno', col_trim_res, 'Promedio', 'Condición'],
                use_container_width=True,
                key=f"editor_{trimestre_trabajo}_{anio_sel}_{div_sel}"
            )

            if st.button("💾 Recalcular y Guardar en Excel"):
                for col in cols_sub:
                    edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce').fillna(0.0).astype(float)

                edited_df[col_trim_res] = edited_df[cols_sub].replace(0, pd.NA).mean(axis=1).round(2).fillna(0.0).astype(float)

                for t_col in ['Nota_1er_Trim.', 'Nota_2do_Trim.', 'Nota_3er_Trim.']:
                    if t_col != col_trim_res:
                        edited_df[t_col] = pd.to_numeric(
                            st.session_state["notas_df"].loc[edited_df.index, t_col], 
                            errors='coerce'
                        ).fillna(0.0).astype(float)

                edited_df['Promedio'] = (
                    edited_df['Nota_1er_Trim.'] + edited_df['Nota_2do_Trim.'] + edited_df['Nota_3er_Trim.']
                ) / 3
                edited_df['Promedio'] = edited_df['Promedio'].round(2).astype(float)
                edited_df['Condición'] = edited_df['Promedio'].apply(lambda x: 'Aprobado' if x >= 6 else 'Desaprobado')

                for col in edited_df.columns:
                    if col in columnas_numericas_notas:
                        st.session_state["notas_df"][col] = st.session_state["notas_df"][col].astype(float)
                        st.session_state["notas_df"].loc[edited_df.index, col] = edited_df[col].astype(float).values
                    else:
                        st.session_state["notas_df"].loc[edited_df.index, col] = edited_df[col].values

                guardar_excel_completo()
                st.success("✅ ¡Notas y promedios recalculados y guardados correctamente en Excel!")
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

        # -----------------------------------------------------------------------------
        # OPTION 2: 📆 CARGA DE EVALUACIONES A LA AGENDA (DOCENTE)
        # -----------------------------------------------------------------------------
        with tab_agenda_doc:
            st.subheader("📆 Publicar Evaluación o Entrega en la Agenda Escolar")

            with st.form("form_nueva_evaluacion_agenda"):
                a_col1, a_col2 = st.columns(2)
                with a_col1:
                    curso_agenda = st.selectbox("Curso:", opciones_cursos)
                    tipo_eval = st.selectbox("Tipo de Evento:", ["Prueba Escrita", "Lección Oral", "Entrega de Trabajo Práctico", "Exposición Groupal"])
                with a_col2:
                    fecha_eval = st.date_input("Fecha programada:", datetime.now())
                    titulo_eval = st.text_input("Tema / Contenidos a Evaluar:", placeholder="Ej. Capítulos 1 al 3 - Historia Argentina")

                btn_agendar = st.form_submit_button("🗓️ Guardar Examen en Agenda")

            if btn_agendar:
                if titulo_eval.strip():
                    anio_ag = int(curso_agenda.split("°")[0])
                    div_ag = curso_agenda.split(" ")[1]

                    id_ev = f"EVA-{len(st.session_state['agenda_df']) + 1:03d}"
                    nuevo_evento = {
                        "ID_Evento": id_ev,
                        "Fecha": fecha_eval.strftime('%d/%m/%Y'),
                        "Año": anio_ag,
                        "División": div_ag,
                        "Materia": prof_data['Materia_Principal'],
                        "Tipo_Evento": tipo_eval,
                        "Título_Descripción": titulo_eval.strip(),
                        "Publicado_Por": f"Prof. {prof_data['Nombre']} {prof_data['Apellido']}"
                    }

                    st.session_state["agenda_df"] = pd.concat([
                        st.session_state["agenda_df"], 
                        pd.DataFrame([nuevo_evento])
                    ], ignore_index=True)

                    guardar_excel_completo()
                    st.success(f"✅ ¡Evaluación agendada para el {fecha_eval.strftime('%d/%m/%Y')}!")
                    st.rerun()
                else:
                    st.error("⚠️ Ingrese el tema o descripción de la evaluación.")

            st.markdown("---")
            st.markdown("##### 📅 Próximas Evaluaciones Agendadas:")
            st.dataframe(st.session_state["agenda_df"][st.session_state["agenda_df"]['Materia'] == prof_data['Materia_Principal']], use_container_width=True)

    # --- VISTA PRECEPTORÍA ---
    elif user_info['rol'] == 'Preceptor':
        st.title("📋 Control Diario, Asistencia, Conducta y Agenda")
        
        tab_asistencia, tab_conducta, tab_agenda_prec = st.tabs(["📋 Tomar Asistencia", "📝 Partes Disciplinarios", "📆 Agenda Escolar"])

        with tab_asistencia:
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
                guardar_excel_completo()
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

        with tab_conducta:
            st.subheader("📝 Registrar Nuevo Parte Disciplinario / Felicitación")

            cursos_unicos_p = (
                st.session_state["alumnos_df"]['Año'].astype(str).str.strip() + "°" + 
                st.session_state["alumnos_df"]['División'].astype(str).str.strip()
            ).unique()

            col_p1, col_p2 = st.columns([1, 2])
            with col_p1:
                curso_cond_sel = st.selectbox("Seleccionar Curso:", sorted(list(cursos_unicos_p)), key="cond_curso")
                partes_c = curso_cond_sel.split("°")
                anio_c, div_c = int(partes_c[0]), partes_c[1]

                alumnos_curso_c = st.session_state["alumnos_df"][
                    (st.session_state["alumnos_df"]['Año'] == anio_c) & 
                    (st.session_state["alumnos_df"]['División'] == div_c)
                ]

                if not alumnos_curso_c.empty:
                    mapa_alumnos_c = {
                        f"{r['Apellido']}, {r['Nombre']}": r['ID_Alumno'] 
                        for _, r in alumnos_curso_c.iterrows()
                    }
                    alumno_cond_nom = st.selectbox("Seleccionar Alumno:", list(mapa_alumnos_c.keys()))
                    id_alumno_cond = mapa_alumnos_c[alumno_cond_nom]
                else:
                    id_alumno_cond = None
                    st.warning("Sin alumnos registrados en este curso.")

            with col_p2:
                if id_alumno_cond:
                    with st.form("form_nuevo_parte_conducta"):
                        st.markdown(f"##### Parte de Conducta para: **{alumno_cond_nom}**")
                        
                        f_col1, f_col2 = st.columns(2)
                        with f_col1:
                            tipo_evento = st.selectbox("Tipo de Evento:", [
                                "Llamado de Atención Verbal",
                                "Amonestación",
                                "Sanción / Suspensión",
                                "Felicitación / Reconocimiento"
                            ])
                        with f_col2:
                            fecha_evento = st.date_input("Fecha del Suceso:", datetime.now())

                        motivo_detalle = st.text_area("Descripción Detallada del Motivo / Incidente:", placeholder="Escriba aquí lo sucedido...")

                        btn_registrar_parte = st.form_submit_button("💾 Guardar Parte de Conducta")

                    if btn_registrar_parte:
                        if motivo_detalle.strip():
                            id_parte = f"PAR-{len(st.session_state['conducta_df']) + 1:03d}"
                            nuevo_parte = {
                                "ID_Parte": id_parte,
                                "ID_Alumno": id_alumno_cond,
                                "Alumno": alumno_cond_nom,
                                "Año": anio_c,
                                "División": div_c,
                                "Fecha": fecha_evento.strftime('%d/%m/%Y'),
                                "Tipo_Evento": tipo_evento,
                                "Motivo_Detalle": motivo_detalle.strip(),
                                "Registrado_Por": user_info['nombre']
                            }

                            st.session_state["conducta_df"] = pd.concat([
                                st.session_state["conducta_df"], 
                                pd.DataFrame([nuevo_parte])
                            ], ignore_index=True)

                            guardar_excel_completo()
                            st.success(f"✅ ¡Parte registrado exitosamente bajo el ID `{id_parte}`!")
                            st.rerun()
                        else:
                            st.error("⚠️ Complete la descripción del motivo antes de guardar.")

            st.markdown("---")
            st.subheader("📲 Notificar Parte Disciplinario por WhatsApp al Tutor")

            if not st.session_state["conducta_df"].empty:
                partes_recientes = st.session_state["conducta_df"].tail(10)
                st.dataframe(partes_recientes[['ID_Parte', 'Alumno', 'Fecha', 'Tipo_Evento', 'Motivo_Detalle', 'Registrado_Por']], use_container_width=True)

                parte_sel_id = st.selectbox("Seleccionar ID de Parte para notificar al Tutor:", partes_recientes['ID_Parte'].tolist())
                row_parte = partes_recientes[partes_recientes['ID_Parte'] == parte_sel_id].iloc[0]

                info_al = st.session_state["alumnos_df"][st.session_state["alumnos_df"]['ID_Alumno'] == row_parte['ID_Alumno']]
                if not info_al.empty:
                    info_al = info_al.iloc[0]
                    tutor_nom = f"{info_al.get('Nombre_Tutor', '')} {info_al.get('Apellido_Tutor', '')}".strip() or info_al.get('Tutor_Responsab', 'Tutor/a')
                    tel_contacto = info_al.get('Telefono_Contacto', '')

                    msg_conducta = (
                        f"Estimado/a {tutor_nom}, le notificamos desde la Preceptoría que se ha registrado "
                        f"un evento de conducta (*{row_parte['Tipo_Evento']}*) para el estudiante *{row_parte['Alumno']}* "
                        f"con fecha {row_parte['Fecha']}.\n\n"
                        f"Detalle: {row_parte['Motivo_Detalle']}.\n\n"
                        f"Por favor, póngase en contacto con el establecimiento."
                    )

                    st.text_area("Mensaje de Notificación:", value=msg_conducta, height=120)

                    if pd.notna(tel_contacto) and str(tel_contacto) != '':
                        link_wa_c = generar_link_whatsapp(tel_contacto, msg_conducta)
                        st.markdown(f'<a href="{link_wa_c}" target="_blank"><button style="background-color:#25D366; color:white; padding:8px 16px; border:none; border-radius:8px; font-weight:bold; cursor:pointer;">📲 Notificar por WhatsApp</button></a>', unsafe_allow_html=True)
                    else:
                        st.warning("⚠️ Sin teléfono de contacto registrado para el tutor.")

        with tab_agenda_prec:
            st.subheader("📆 Calendario General de Exámenes y Eventos por Curso")
            st.dataframe(st.session_state["agenda_df"], use_container_width=True)
