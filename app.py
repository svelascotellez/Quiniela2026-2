import os
import streamlit as st
import pandas as pd
from io import BytesIO
from consolidate_quinielas import (
    process_master_file,
    evaluate_participant,
    create_premium_leaderboard,
    generate_whatsapp_report
)

# Configuración de página
st.set_page_config(
    page_title="Quiniela Mundial 2026",
    page_icon="🏆",
    layout="wide"
)

# ==========================================
# ESTILOS CSS - ESTÉTICA FIFA 2026
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600&family=Oswald:wght@500;700&display=swap');

/* Fuente general y colores */
html, body, [class*="css"] {
    font-family: 'Montserrat', sans-serif;
}

/* Títulos con fuente condensada y estilo deportivo */
h1, h2, h3 {
    font-family: 'Oswald', sans-serif !important;
    text-transform: uppercase;
    letter-spacing: 1px;
}

h1 {
    color: #FFFFFF !important;
    font-size: 3.5rem !important;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.5);
    border-bottom: 3px solid #E10600; /* Acento rojo/magenta */
    padding-bottom: 10px;
    margin-bottom: 30px !important;
}

h2 {
    color: #00E5FF !important; /* Acento cyan Neón */
}

/* Estilo de botones estilo FIFA */
div.stButton > button:first-child {
    background: linear-gradient(135deg, #0253CC 0%, #002266 100%);
    color: white;
    font-family: 'Oswald', sans-serif;
    font-size: 1.2rem;
    padding: 0.5rem 2rem;
    border: none;
    border-radius: 30px; /* Botones redondeados tipo píldora */
    text-transform: uppercase;
    box-shadow: 0 4px 15px rgba(2, 83, 204, 0.4);
    transition: all 0.3s ease;
}

div.stButton > button:first-child:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(2, 83, 204, 0.6);
    background: linear-gradient(135deg, #002266 0%, #0253CC 100%);
    color: #00E5FF !important; /* Texto cyan brillante al hover */
    border: 1px solid #00E5FF !important;
}

div.stDownloadButton > button:first-child {
    background: linear-gradient(135deg, #E10600 0%, #8B0000 100%);
    color: white;
    font-family: 'Oswald', sans-serif;
    border-radius: 30px;
    box-shadow: 0 4px 15px rgba(225, 6, 0, 0.4);
    border: none;
}

div.stDownloadButton > button:first-child:hover {
    transform: translateY(-2px);
    background: linear-gradient(135deg, #8B0000 0%, #E10600 100%);
    box-shadow: 0 6px 20px rgba(225, 6, 0, 0.6);
    border: 1px solid #FF4B4B !important;
    color: white !important;
}

/* Estilos para las pestañas (Tabs) */
[data-baseweb="tab"] {
    background-color: transparent !important;
    border: none !important;
}

[data-baseweb="tab-list"] {
    gap: 20px;
    border-bottom: 2px solid #141C2A;
}

[data-baseweb="tab"] p {
    font-family: 'Oswald', sans-serif;
    font-size: 1.3rem;
    color: #8892A3;
    text-transform: uppercase;
}

[aria-selected="true"] {
    border-bottom: 4px solid #00E5FF !important;
}

[aria-selected="true"] p {
    color: #FFFFFF !important;
}

/* Estilizar las cajas/paneles y tarjetas */
.stAlert {
    border-radius: 10px;
    border-left: 5px solid;
}

/* Efecto en los DataFrames */
[data-testid="stDataFrame"] {
    border: 1px solid #141C2A;
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 4px 10px rgba(0,0,0,0.3);
}

</style>
""", unsafe_allow_html=True)

# ==========================================
# AUTENTICACIÓN
# ==========================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔒 Acceso a la Quiniela")
    st.markdown("Por favor, ingresa tus credenciales para acceder al sistema.")
    with st.form("login_form"):
        username = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submit = st.form_submit_button("Ingresar")
        
        if submit:
            # Aquí se define el usuario y contraseña
            if username == "admin" and password == "mundial2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
    st.stop()

# ==========================================
# CONFIGURACIÓN DE DIRECTORIOS LOCALES
# ==========================================
DATA_DIR = "data"
MASTER_DIR = os.path.join(DATA_DIR, "master")
PARTICIPANTS_DIR = os.path.join(DATA_DIR, "participants")

# Crear carpetas si no existen
os.makedirs(MASTER_DIR, exist_ok=True)
os.makedirs(PARTICIPANTS_DIR, exist_ok=True)

MASTER_FILE_PATH = os.path.join(MASTER_DIR, "master.xlsx")

def get_saved_participants():
    """Devuelve la lista de archivos de participantes guardados."""
    if not os.path.exists(PARTICIPANTS_DIR):
        return []
    return [f for f in os.listdir(PARTICIPANTS_DIR) if f.endswith('.xlsx')]

def delete_participant(filename):
    """Elimina un archivo de participante del almacenamiento local."""
    path = os.path.join(PARTICIPANTS_DIR, filename)
    if os.path.exists(path):
        os.remove(path)

# ==========================================
# INTERFAZ GRÁFICA
# ==========================================
st.title("🏆 Consolidación de Quinielas - Mundial 2026")
st.markdown("""
Sube el archivo maestro y los archivos de los participantes. El sistema los **guardará automáticamente** para que no tengas que volver a subirlos la próxima vez.
""")

# Botones de descargas útiles
TEMPLATE_PATH = "Quiniela_Mundial_2026_Final_vacia.xlsx"
SCHEDULE_PATH = "FWC26 Match Schedule_v17_10042026_EN.pdf"

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    if os.path.exists(TEMPLATE_PATH):
        with open(TEMPLATE_PATH, "rb") as f:
            template_bytes = f.read()
        st.download_button(
            label="📄 Descargar Plantilla para Participar",
            data=template_bytes,
            file_name="Quiniela_Mundial_2026_Final_vacia.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga este archivo, llénalo con tus pronósticos y súbelo para participar.",
            use_container_width=True
        )

with col_btn2:
    if os.path.exists(SCHEDULE_PATH):
        with open(SCHEDULE_PATH, "rb") as f:
            schedule_bytes = f.read()
        st.download_button(
            label="📅 Descargar Calendario Oficial FIFA",
            data=schedule_bytes,
            file_name="Calendario_Oficial_Mundial_2026.pdf",
            mime="application/pdf",
            help="Descarga el calendario oficial de partidos de la FIFA en formato PDF.",
            use_container_width=True
        )

st.write("") # Espacio


tab_app, tab_instructions = st.tabs(["🎮 Panel de Control", "📖 Instrucciones y Reglas"])

with tab_app:
    col1, col2 = st.columns(2)

    with col1:
        st.header("1. Archivo Maestro")
        
        # Mostrar estado del archivo maestro actual
        if os.path.exists(MASTER_FILE_PATH):
            st.success("✅ Hay un archivo Maestro guardado en el sistema.")
        else:
            st.warning("⚠️ No hay archivo Maestro guardado.")
            
        master_file_upload = st.file_uploader("Actualizar archivo Maestro", type=["xlsx"], key="master")
        if master_file_upload:
            # Guardar archivo maestro físico
            with open(MASTER_FILE_PATH, "wb") as f:
                f.write(master_file_upload.getbuffer())
            st.session_state.calcular = False
            st.success("Archivo Maestro guardado y actualizado con éxito.")

    with col2:
        st.header("2. Gestión de Participantes")
        
        # Subir nuevos participantes
        participant_uploads = st.file_uploader("Añadir nuevos participantes", type=["xlsx"], accept_multiple_files=True, key="participants")
        if participant_uploads:
            for p_file in participant_uploads:
                path = os.path.join(PARTICIPANTS_DIR, p_file.name)
                with open(path, "wb") as f:
                    f.write(p_file.getbuffer())
            st.session_state.calcular = False
            st.success(f"Se han guardado {len(participant_uploads)} nuevos participantes.")
            
        st.subheader("Participantes Registrados")
        saved_participants = get_saved_participants()
        
        if not saved_participants:
            st.info("No hay participantes registrados aún.")
        else:
            # Mostrar lista con botón de eliminar
            for p_file in saved_participants:
                col_name, col_btn = st.columns([4, 1])
                with col_name:
                    st.write(f"📄 {p_file}")
                with col_btn:
                    if st.button("🗑️", key=f"del_{p_file}", help=f"Eliminar {p_file}"):
                        delete_participant(p_file)
                        st.session_state.calcular = False
                        st.rerun()

    st.divider()

    # ==========================================
    # SECCIÓN DE CÁLCULO
    # ==========================================
    if st.button("🚀 Calcular Clasificación", type="primary", use_container_width=True):
        st.session_state.calcular = True

    if st.session_state.get("calcular", False):
        if not os.path.exists(MASTER_FILE_PATH):
            st.error("Debes subir un archivo Maestro primero.")
            st.session_state.calcular = False
            st.stop()
            
        saved_participants = get_saved_participants()
        if not saved_participants:
            st.error("Debes añadir al menos un participante para calcular.")
            st.session_state.calcular = False
            st.stop()
            
        with st.spinner("Procesando resultados oficiales y calculando puntos..."):
            try:
                # 1. Procesar archivo maestro guardado
                master_groups, master_groups_list, master_ko_advances, master_ko_matches, max_possible_pts, group_matches_played = process_master_file(MASTER_FILE_PATH)
                
                # 2. Procesar participantes guardados
                participant_results = []
                progress_bar = st.progress(0)
                
                for i, p_filename in enumerate(saved_participants):
                    p_path = os.path.join(PARTICIPANTS_DIR, p_filename)
                    
                    # Extraer el nombre del participante
                    p_name = p_filename.replace('.xlsx', '').replace('Quiniela_Mundial_2026_', '').replace('Quiniela_Mundial_', '').replace('Quiniela_', '')
                    p_name = p_name.replace('_', ' ').strip()
                    
                    try:
                        scores = evaluate_participant(p_path, master_groups, master_groups_list, master_ko_advances, master_ko_matches)
                        scores["name"] = p_name
                        scores["file"] = p_filename
                        participant_results.append(scores)
                    except Exception as e:
                        st.error(f"Error procesando la quiniela de {p_name}: {e}")
                    
                    progress_bar.progress((i + 1) / len(saved_participants))
                
                if not participant_results:
                    st.error("No se pudo evaluar a los participantes. Revisa los archivos.")
                    st.session_state.calcular = False
                    st.stop()
                    
                # 3. Ordenar resultados
                participant_results.sort(
                    key=lambda x: (-x["total_points"], -x["exact_aciertos_totales"], x["name"].lower())
                )
                
                # CREAR PESTAÑAS
                tab_clasif, tab_resul = st.tabs(["🏆 Clasificación General", "⚽ Resultados Oficiales"])
                
                with tab_clasif:
                    # 4. Mostrar DataFrame (Tabla Resumen en Web)
                    st.subheader("📊 Tabla de Clasificación")
                    df_data = []
                    for idx, r in enumerate(participant_results, 1):
                        df_data.append({
                            "Posición": idx,
                            "Nombre": r["name"],
                            "Puntos Totales": r["total_points"],
                            "Efectividad": f"{(r['total_points']/max_possible_pts*100) if max_possible_pts > 0 else 0:.1f}%",
                            "Marcadores Exactos": r["exact_aciertos_totales"]
                        })
                    st.dataframe(pd.DataFrame(df_data), use_container_width=True)
                    
                    # 5. Generar Excel en memoria
                    excel_buffer = BytesIO()
                    create_premium_leaderboard(participant_results, excel_buffer, max_possible_pts, master_groups_list, master_ko_matches)
                    excel_buffer.seek(0)
                    
                    # 6. Botón de descarga de Excel
                    st.download_button(
                        label="📥 Descargar Leaderboard Oficial (Excel)",
                        data=excel_buffer,
                        file_name="Clasificacion_Quiniela_Mundial_2026.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                    # 7. Texto de WhatsApp
                    st.subheader("📱 Reporte para WhatsApp")
                    whatsapp_text = generate_whatsapp_report(participant_results, max_possible_pts)
                    st.code(whatsapp_text, language="text")

                with tab_resul:
                    st.subheader("⚽ Resultados de Partidos Disputados")
                    st.write("A continuación se muestran los partidos que ya tienen un resultado oficial cargado en el archivo Maestro.")
                    
                    # Partidos de Fase de Grupos
                    played_groups = [m for m in master_groups_list if m["score1"] is not None and m["score2"] is not None]
                    
                    if played_groups:
                        st.markdown("### Fase de Grupos")
                        for m in played_groups:
                            st.markdown(f"- **{m['team1']}** `{m['score1']} - {m['score2']}` **{m['team2']}**")
                    else:
                        st.info("Aún no hay resultados de Fase de Grupos.")
                        
                    # Partidos de Eliminatorias
                    played_ko = [m for m in master_ko_matches.values() if m["score1"] is not None and m["score2"] is not None]
                    
                    if played_ko:
                        st.markdown("### Eliminatorias")
                        for m in played_ko:
                            phase_map = {
                                "r32": "Dieciseisavos de Final",
                                "r16": "Octavos de Final",
                                "qf": "Cuartos de Final",
                                "sf": "Semifinales",
                                "t3p": "Tercer Lugar",
                                "final": "Gran Final"
                            }
                            phase_name = phase_map.get(m.get("phase", ""), "Eliminatoria")
                            st.markdown(f"- *{phase_name}:* **{m['team1']}** `{m['score1']} - {m['score2']}` **{m['team2']}**")
                    
            except Exception as e:
                st.error(f"Ocurrió un error general: {e}")

with tab_instructions:
    st.markdown("""
    ## 📖 Cómo funciona el sistema
    
    Esta aplicación automatiza el proceso de evaluación de una quiniela deportiva basada en archivos Excel.
    
    ### 1. El Archivo Maestro
    El organizador (administrador) debe mantener un único archivo de Excel llamado "Maestro", en el cual irá capturando los resultados reales de los partidos a medida que el Mundial avanza. Este archivo se sube en la sección **Archivo Maestro**.
    
    ### 2. Los Archivos de Participantes
    Cada persona que entra a la quiniela llena un archivo de Excel con sus pronósticos antes de que empiece el mundial. Estos archivos deben ser subidos en la sección **Añadir nuevos participantes**. El sistema los guardará automáticamente.
    
    ### 3. Calcular Clasificación
    Al hacer clic en el botón **Calcular Clasificación**, el sistema:
    1. Lee los resultados reales del Maestro.
    2. Lee las predicciones de todos los participantes registrados.
    3. Compara y asigna puntos automáticamente según las reglas de puntuación.
    4. Genera una tabla de posiciones en la pantalla y permite descargar un reporte en Excel de aspecto profesional.
    
    ---
    
    ## 🎯 Sistema de Puntuación
    
    El algoritmo está pre-configurado de la siguiente manera:
    
    **Fase de Grupos:**
    - **3 puntos** por acertar el marcador exacto de un partido.
    - **1 punto** por acertar el resultado (ganador o empate) sin atinarle al marcador exacto.
    
    **Fase Eliminatoria:**
    - **5 puntos** por acertar el marcador exacto (solo si atinaste a los equipos que jugaban).
    - **2 puntos** por acertar al ganador del partido.
    - **Puntos adicionales por clasificación:** (Si el equipo que predijiste realmente avanza a esa fase)
      - *Dieciseisavos:* +2 puntos por equipo
      - *Octavos:* +4 puntos por equipo
      - *Cuartos:* +6 puntos por equipo
      - *Semifinales:* +8 puntos por equipo
      - *Tercer Lugar / Final:* +10 / +12 puntos por equipo
      
    **Podio (Predicciones Finales):**
    - **15 puntos** al acertar al Campeón.
    - **10 puntos** al acertar al Subcampeón.
    - **10 puntos** al acertar al 3er Lugar.
    """)
