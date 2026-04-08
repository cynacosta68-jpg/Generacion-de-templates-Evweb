from google.colab import files

# Content for README.md
readme_content = """# Generador de Templates Excel con Streamlit

Esta aplicación permite a los usuarios subir un archivo Excel (`.xlsx`) y procesar sus datos para generar uno o varios archivos Excel con un formato de template específico, comprimiéndolos en un archivo ZIP para su descarga.

## Estructura del Proyecto

El proyecto consta de dos archivos principales:

- `requirements.txt`: Lista las dependencias de Python necesarias.
- `app.py`: Contiene el código fuente de la aplicación Streamlit.

## `requirements.txt`

Este archivo lista las librerías de Python que deben instalarse para que la aplicación funcione correctamente. Su contenido es:

```
pandas
openpyxl
streamlit
```

## `app.py`

Este es el script principal de la aplicación Streamlit. Realiza las siguientes funciones:

1.  **Definición del Template Modelo**: La estructura del template de Excel de salida (columnas, anchos, formatos, estilos) está definida directamente en el código Python, eliminando la necesidad de un archivo de template externo.
2.  **Mapeo de Columnas**: Identifica y mapea automáticamente las columnas del archivo Excel de entrada a las columnas del template de salida, siendo tolerante a variaciones en nombres (acentos, espacios, mayúsculas/minúsculas).
3.  **Procesamiento de Datos**: Lee el archivo Excel de entrada, construye las filas de salida y las distribuye en múltiples templates si la cantidad de filas excede un límite predefinido (1500 filas por template).
4.  **Generación de ZIP**: Combina todos los templates generados en un único archivo ZIP.
5.  **Interfaz de Usuario (Streamlit)**: Proporciona una interfaz web simple para que los usuarios suban su archivo Excel y descarguen el ZIP resultante.

## Cómo ejecutar la aplicación localmente

Para ejecutar esta aplicación en tu máquina local, sigue estos pasos:

1.  **Guarda los archivos**: Asegúrate de tener `requirements.txt` y `app.py` en la misma carpeta.

2.  **Crea un entorno virtual (opcional pero recomendado)**:
    ```bash
    python -m venv venv
    source venv/bin/activate  # En Linux/macOS
    .\venv\Scripts\activate   # En Windows
    ```

3.  **Instala las dependencias**: Utiliza `pip` para instalar todas las librerías listadas en `requirements.txt`:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Ejecuta la aplicación Streamlit**: Navega a la carpeta donde guardaste los archivos y ejecuta el siguiente comando:
    ```bash
    streamlit run app.py
    ```

    Esto abrirá automáticamente la aplicación en tu navegador web. Si no se abre, copia la URL que aparece en tu terminal (generalmente `http://localhost:8501`).

## Uso de la Aplicación

1.  **Subir archivo**: Haz clic en el botón "📤 Sube tu archivo Excel (.xlsx)" y selecciona el archivo Excel que deseas procesar.
2.  **Procesamiento**: La aplicación procesará automáticamente el archivo, identificará las columnas y generará los templates.
3.  **Descargar**: Una vez completado el procesamiento, aparecerá un botón "⬇️ Descargar Templates.zip" para que puedas descargar el archivo ZIP con todos los templates generados.

---

¡Esperamos que esta herramienta te sea útil!
"""

# Create README.md file
with open('README.md', 'w') as f:
    f.write(readme_content)

print("Archivos creados: README.md")

# Provide download link
files.download('README.md')
