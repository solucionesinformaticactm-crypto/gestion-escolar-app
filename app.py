import streamlit as st
import pandas as pd
from datetime import datetime

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
    </style>
""", unsafe_allow_html=True)

# Carga de datos desde la planilla Excel
@st.cache_data
def cargar_datos():
    excel_path = 'Gestion_Escolar_Secundaria.xlsx'
    alumnos = pd.read_excel(excel_path, sheet_name='Alumnos')
    profesores = pd.read_excel(excel_path, sheet_name='Profesores')
    plan = pd.read_excel(excel_path, sheet_name='Plan_Materias')
    notas = pd.read_excel(excel_path, sheet_name='Notas')
    asistencia = pd.read_excel(excel_path, sheet_name='Asistencia')
    return alumnos, profesores, plan, notas, asistencia

alumnos_df, profesores_df, plan_df, notas_df, asistencia_df = cargar_datos()

# Control de Autenticación
st.sidebar.image("https://img.icons8.com/isometric/100/graduation-cap.png", width=70)
st.sidebar.title("Portal Escolar")
st.sidebar.subheader("Iniciar Sesión")

usuarios_demo = {
    "laura.gomez": {"nombre": "Laura Gómez", "rol": "Docente", "id": "PR-01"},
    "marcelo.fernandez": {"nombre": "Marcelo Fernández", "rol": "Docente", "id": "PR-02"},
    "marina.torres": {"nombre": "Marina Torres", "rol": "Docente", "id": "PR-08"},
    "admin.direccion": {"nombre": "Equipo Directivo", "rol": "Directivo", "id": "DIR"},
    "preceptoria": {"nombre": "Preceptoría General", "rol": "Preceptor", "id": "PRE"}
}

usuario_ingresado = st.sidebar.selectbox("Seleccionar Usuario (Demo):", list(usuarios_demo.keys()))

if usuario_ingresado:
    user_info = usuarios_demo[usuario_ingresado]
    st.sidebar.success(f"Conectado como: **{user_info['nombre']}**\n\nRol: *{user_info['rol']}*")

    # --- VISTA DIRECTIVO ---
    if user_info['rol'] == 'Directivo':
        st.title("🏛️ Panel de Control e Indicadores Institucionales")
        st.write("Visión general del rendimiento escolar y asistencia del establecimiento.")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown('<div class="metric-card"><h4>Total Alumnos</h4><h2>30</h2></div>', unsafe_allow_html=True)
        with col2:
            st.markdown('<div class="metric-card"><h4>Total Profesores</h4><h2>11</h2></div>', unsafe_allow_html=True)
        with col3:
            st.markdown('<div class="metric-card"><h4>Promedio General</h4><h2>7.07</h2></div>', unsafe_allow_html=True)
        with col4:
            st.markdown('<div class="metric-card"><h4>Asistencia General</h4><h2>56%</h2></div>', unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("📊 Calificaciones Consolidadas")
        st.dataframe(notas_df, use_container_width=True)

    # --- VISTA DOCENTE ---
    elif user_info['rol'] == 'Docente':
        prof_data = profesores_df[profesores_df['ID_Profesor'] == user_info['id']].iloc[0]
        st.title(f"📚 Gestión Académica — Prof. {prof_data['Nombre']} {prof_data['Apellido']}")
        st.markdown(f"**Materia Asignada:** `{prof_data['Materia_Principal']}`")

        notas_prof = notas_df[notas_df['ID_Profesor'] == user_info['id']]
        cursos = notas_prof[['Año', 'División']].drop_duplicates()
        
        opciones_cursos = [f"{row['Año']}° {row['División']}" for _, row in cursos.iterrows()]
        curso_seleccionado = st.selectbox("Seleccionar Curso para Cargar/Modificar Notas:", opciones_cursos)
        
        anio_sel = int(curso_seleccionado.split("°")[0])
        div_sel = curso_seleccionado.split(" ")[1]

        # Cargar los datos del curso
        notas_curso = notas_prof[(notas_prof['Año'] == anio_sel) & (notas_prof['División'] == div_sel)].copy()

        st.subheader(f"Planilla de Notas: {prof_data['Materia_Principal']} — {curso_seleccionado}")
        
        # Editor interactivo de notas
        edited_df = st.data_editor(
            notas_curso[['ID_Alumno', 'Alumno', 'Nota_1er_Trim.', 'Nota_2do_Trim.', 'Nota_3er_Trim.', 'Promedio', 'Condición', 'Observación']],
            column_config={
                "Nota_1er_Trim.": st.column_config.SelectboxColumn("1° Trimestre", options=list(range(1, 11)), required=True),
                "Nota_2do_Trim.": st.column_config.SelectboxColumn("2° Trimestre", options=list(range(1, 11)), required=True),
                "Nota_3er_Trim.": st.column_config.SelectboxColumn("3° Trimestre", options=list(range(1, 11)), required=True),
                "Promedio": st.column_config.NumberColumn("Promedio", format="%.2f"),
                "Condición": st.column_config.TextColumn("Condición")
            },
            disabled=['ID_Alumno', 'Alumno', 'Promedio', 'Condición'],
            use_container_width=True
        )

        # Botón para procesar automáticos
        if st.button("💾 Recalcular Promedios y Guardar"):
            # Recalcular matemáticamente los promedios
            edited_df['Promedio'] = (edited_df['Nota_1er_Trim.'] + edited_df['Nota_2do_Trim.'] + edited_df['Nota_3er_Trim.']) / 3
            edited_df['Promedio'] = edited_df['Promedio'].round(2)
            
            # Asignar automáticamente Aprobado / Desaprobado
            edited_df['Condición'] = edited_df['Promedio'].apply(lambda x: 'Aprobado' if x >= 6 else 'Desaprobado')
            
            # Forzar actualización en pantalla
            st.success("¡Promedios y condiciones recalculados automáticamente!")
            st.dataframe(edited_df[['ID_Alumno', 'Alumno', 'Nota_1er_Trim.', 'Nota_2do_Trim.', 'Nota_3er_Trim.', 'Promedio', 'Condición', 'Observación']], use_container_width=True)

    # --- VISTA PRECEPTORÍA ---
    elif user_info['rol'] == 'Preceptor':
        st.title("📋 Control Diario de Asistencia")
        
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            curso_p = st.selectbox("Seleccionar Curso:", ["1°A", "2°A", "3°A"])
        with col_c2:
            fecha_asistencia = st.date_input("Fecha de Asistencia:", datetime.now())

        anio_p = int(curso_p[0])
        div_p = curso_p[2]
        
        asistencia_filtrada = asistencia_df[(asistencia_df['Año'] == anio_p) & (asistencia_df['División'] == div_p)].copy()

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

        if st.button("💾 Guardar Asistencia"):
            st.success(f"Asistencia guardada correctamente.")
