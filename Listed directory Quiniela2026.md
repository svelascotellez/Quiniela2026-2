**Sistema Integral de Gestión y Evaluación Automatizada para una Quiniela del Mundial de la FIFA 2026**. 

El proyecto combina hojas de cálculo de Excel con un backend en Python para automatizar el cálculo de puntuaciones, lo cual es ideal para gestionar quinielas con múltiples participantes.

### 📁 Componentes Principales del Proyecto

1. **La Plantilla Base (`Quiniela_Mundial_2026_Totalmente_Corregida.xlsx`)**
   - Es el archivo de Excel donde los participantes ingresan sus predicciones.
   - Cuenta con una estructura corregida para mapear correctamente las llaves desde la Fase de Grupos hasta la Gran Final. 
   - Soporta desempates automáticos y propagación de los equipos a través de Dieciseisavos, Octavos, Cuartos, Semifinales, Partido por el Tercer Lugar y la Final.

2. **El Motor de Evaluación (`consolidate_quinielas.py`)**
   - Un script de Python que funciona como el núcleo evaluador. Toma un archivo "Maestro" (con los resultados reales) y lo compara dinámicamente con los archivos de todos los participantes.
   - Utiliza la librería `openpyxl` para leer los datos de Excel evaluando las celdas resultantes de las fórmulas.
   - **Sistema de Puntuación Híbrido:**
     - **Marcadores (Fase de Grupos y Eliminatorias):** Otorga puntos por atinar al marcador exacto o por atinar al resultado (ganador/empate).
     - **Bracket Style (March Madness):** Premia la capacidad de predecir qué equipos avanzan en el torneo, otorgando puntos progresivos (desde 2 pts por clasificar a Dieciseisavos, hasta 12 pts por llegar a la Final).
     - **Podio de Honor:** Puntos de bonificación por predecir correctamente al Campeón, Subcampeón y Tercer Lugar.

3. **Generación de Reportes Automáticos**
   - **Excel Premium (`Clasificacion_Quiniela_Mundial_2026.xlsx`):** El script crea automáticamente una tabla de clasificación (Leaderboard) con formato avanzado (colores esmeralda, celdas de oro/plata/bronce para el Top 3 y cálculo de efectividad).
   - **Reporte para Redes Sociales (`Reporte_WhatsApp_Quiniela_2026.txt`):** Genera un texto con formato listo para ser copiado y pegado en WhatsApp, mostrando las posiciones, puntos y medallas.

### ⚙️ Flujo de Trabajo (Cómo Funciona)

1. El administrador comparte la plantilla Excel con los participantes.
2. Los participantes la llenan y la devuelven. El administrador guarda todos los archivos en una carpeta `Participantes`.
3. A medida que ocurren los partidos reales, el administrador actualiza su propio archivo "Maestro" con los resultados reales.
4. El administrador ejecuta `consolidate_quinielas.py`. El script lee el archivo maestro, evalúa todas las quinielas de la carpeta, y escupe instantáneamente la tabla de clasificación actualizada en Excel y el mensaje de texto para WhatsApp.
