import streamlit as st
import pandas as pd
from io import BytesIO
from consolidate_quinielas import (
    process_master_file,
    evaluate_participant,
    create_premium_leaderboard,
    generate_whatsapp_report
)

st.set_page_config(
    page_title="Quiniela Mundial 2026",
    page_icon="🏆",
    layout="wide"
)

st.title("🏆 Consolidación de Quinielas - Mundial 2026")
st.markdown("""
Sube el archivo maestro (con los resultados oficiales hasta la fecha) y los archivos de los participantes. 
El sistema evaluará automáticamente las puntuaciones y generará la tabla de clasificación.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Archivo Maestro")
    master_file = st.file_uploader("Sube el archivo Maestro con los resultados reales", type=["xlsx"], key="master")

with col2:
    st.subheader("2. Archivos de Participantes")
    participant_files = st.file_uploader("Sube los archivos de todos los participantes", type=["xlsx"], accept_multiple_files=True, key="participants")

if master_file and participant_files:
    if st.button("🚀 Calcular Clasificación", type="primary"):
        with st.spinner("Procesando resultados oficiales..."):
            try:
                # 1. Procesar archivo maestro
                master_groups, master_groups_list, master_ko_advances, master_ko_matches, max_possible_pts, group_matches_played = process_master_file(master_file)
                st.success(f"Archivo maestro procesado. Puntos máximos disputados: {max_possible_pts}")
                
                # 2. Procesar participantes
                participant_results = []
                progress_bar = st.progress(0)
                
                for i, p_file in enumerate(participant_files):
                    # Extraer el nombre del archivo
                    filename = p_file.name
                    p_name = filename.replace('.xlsx', '').replace('Quiniela_Mundial_2026_', '').replace('Quiniela_Mundial_', '').replace('Quiniela_', '')
                    p_name = p_name.replace('_', ' ').strip()
                    
                    try:
                        scores = evaluate_participant(p_file, master_groups, master_groups_list, master_ko_advances, master_ko_matches)
                        scores["name"] = p_name
                        scores["file"] = filename
                        participant_results.append(scores)
                    except Exception as e:
                        st.error(f"Error procesando la quiniela de {p_name}: {e}")
                    
                    progress_bar.progress((i + 1) / len(participant_files))
                
                if not participant_results:
                    st.error("No se pudo evaluar a los participantes. Revisa los archivos.")
                    st.stop()
                    
                # 3. Ordenar resultados
                participant_results.sort(
                    key=lambda x: (-x["total_points"], -x["exact_aciertos_totales"], x["name"].lower())
                )
                
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
                create_premium_leaderboard(participant_results, excel_buffer, max_possible_pts)
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
                
            except Exception as e:
                st.error(f"Ocurrió un error general: {e}")
