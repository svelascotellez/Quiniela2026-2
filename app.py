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
        st.success("Archivo Maestro guardado y actualizado con éxito.")
        st.rerun()

with col2:
    st.header("2. Gestión de Participantes")
    
    # Subir nuevos participantes
    participant_uploads = st.file_uploader("Añadir nuevos participantes", type=["xlsx"], accept_multiple_files=True, key="participants")
    if participant_uploads:
        for p_file in participant_uploads:
            path = os.path.join(PARTICIPANTS_DIR, p_file.name)
            with open(path, "wb") as f:
                f.write(p_file.getbuffer())
        st.success(f"Se han guardado {len(participant_uploads)} nuevos participantes.")
        st.rerun()
        
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
                    st.rerun()

st.divider()

# ==========================================
# SECCIÓN DE CÁLCULO
# ==========================================
if st.button("🚀 Calcular Clasificación", type="primary", use_container_width=True):
    if not os.path.exists(MASTER_FILE_PATH):
        st.error("Debes subir un archivo Maestro primero.")
        st.stop()
        
    saved_participants = get_saved_participants()
    if not saved_participants:
        st.error("Debes añadir al menos un participante para calcular.")
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
                st.stop()
                
            # 3. Ordenar resultados
            participant_results.sort(
                key=lambda x: (-x["total_points"], -x["exact_aciertos_totales"], x["name"].lower())
            )
            
            # CREAR PESTAÑAS
            tab1, tab2 = st.tabs(["🏆 Clasificación General", "⚽ Resultados Oficiales"])
            
            with tab1:
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

            with tab2:
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
