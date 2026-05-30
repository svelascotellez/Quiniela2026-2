# 🏆 Quiniela Mundial 2026 - Plataforma Web

Bienvenido al sistema automatizado de consolidación y evaluación de quinielas para el Mundial de Fútbol 2026. Este proyecto transforma una serie de archivos de Excel (quinielas de los participantes) en una plataforma web interactiva y en tiempo real.

## ✨ Características Principales

* **Interfaz Gráfica (Web):** Construida con [Streamlit](https://streamlit.io/), permite gestionar la quiniela desde cualquier navegador.
* **Base de Datos Local Inteligente:** Guarda los pronósticos de los participantes y el archivo "Maestro" localmente, evitando tener que subir los archivos repetidamente.
* **Cálculo Automático de Puntos:** Compara las predicciones de docenas de participantes contra los resultados oficiales en cuestión de segundos.
* **Reportes Multicanal:**
  * Visualización en tiempo real de la tabla de posiciones en la web.
  * Generación de un reporte de texto con emojis, listo para copiar y enviar por WhatsApp.
  * Exportación de un reporte en Excel de aspecto **premium** con múltiples pestañas (Clasificación General y Resultados Oficiales).
* **Pestañas Integradas:**
  * *Panel de Control:* Para subir y gestionar participantes.
  * *Instrucciones y Reglas:* Para que los jugadores entiendan fácilmente cómo ganan puntos.

---

## 🛠️ Requisitos e Instalación

Para ejecutar este proyecto en tu computadora local, necesitas tener instalado **Python 3.8+**.

1. **Clonar el repositorio o descargar los archivos.**
2. **Instalar las dependencias:**
   Abre una terminal en la carpeta del proyecto y ejecuta:
   ```bash
   pip install -r requirements.txt
   ```

*(Las dependencias principales son `streamlit`, `pandas` y `openpyxl`).*

---

## 🚀 Cómo ejecutar la aplicación

Una vez instaladas las dependencias, levanta el servidor web de Streamlit ejecutando:

```bash
streamlit run app.py
```

Esto abrirá automáticamente una pestaña en tu navegador web (por defecto en `http://localhost:8501`).

---

## 📂 Estructura del Proyecto

* `app.py`: El archivo principal de la interfaz web de Streamlit. Contiene el diseño de las pestañas, el gestor de archivos y la tabla dinámica.
* `consolidate_quinielas.py`: El "motor" lógico de la aplicación. Contiene el algoritmo que lee los Excels, calcula los puntos y genera el Excel Premium de salida.
* `requirements.txt`: Lista de librerías necesarias para correr el proyecto.
* `data/` *(Generada automáticamente)*:
  * `data/master/`: Almacena el archivo maestro actual.
  * `data/participants/`: Almacena todos los Excels de los participantes registrados.

---

## 🎯 Sistema de Puntuación (Algoritmo)

El sistema evalúa las quinielas basándose en el siguiente esquema (configurable en `consolidate_quinielas.py`):

### Fase de Grupos
* **3 puntos:** Acertar el marcador exacto.
* **1 punto:** Acertar el resultado (ganador/empate) sin atinarle al marcador.

### Fase Eliminatoria
* **5 puntos:** Acertar el marcador exacto (aplicable solo si acertaste a los equipos que llegaron a ese partido).
* **2 puntos:** Acertar al ganador del partido.
* **Puntos por avance de equipo:**
  * *Dieciseisavos:* +2 pts por equipo
  * *Octavos:* +4 pts por equipo
  * *Cuartos:* +6 pts por equipo
  * *Semifinales:* +8 pts por equipo
  * *Tercer Lugar:* +10 pts por equipo
  * *Final:* +12 pts por equipo

### Podio
* **15 puntos:** Acertar al Campeón Mundial.
* **10 puntos:** Acertar al Subcampeón.
* **10 puntos:** Acertar al Tercer Lugar.

---

## 🌐 Publicación en Internet (Opcional)

Si deseas que cualquier persona interactúe con la plataforma sin instalar Python:
1. Sube tu código a un repositorio de **GitHub**.
2. Ingresa a **[Streamlit Community Cloud](https://share.streamlit.io/)** e inicia sesión con tu cuenta de GitHub.
3. Haz clic en "New app" y selecciona tu repositorio.
4. Asegúrate de apuntar al archivo `app.py`. ¡En 2 minutos tendrás un enlace público!

> **Nota:** La nube comunitaria de Streamlit es efímera. Si la aplicación entra en reposo por inactividad, los archivos temporales almacenados en la carpeta `data/` se borrarán. Para un despliegue permanente, se recomienda usar almacenamiento en la nube externo (como AWS S3).
