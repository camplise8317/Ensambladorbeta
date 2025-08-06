import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import os
import zipfile
from io import BytesIO

# --- Configuración de la Página de Streamlit ---
st.set_page_config(
    page_title="Generador de Documentos",
    page_icon="📄",
    layout="centered"
)

# --- Funciones Auxiliares ---

def generar_documentos(plantilla_bytes, excel_bytes, columna_nombre_archivo):
    """
    Función principal que procesa los archivos y genera un zip con los documentos.
    """
    try:
        # Cargar el DataFrame de pandas desde los bytes del archivo Excel
        df = pd.read_excel(excel_bytes)
    except Exception as e:
        st.error(f"Error al leer el archivo Excel: {e}")
        return None

    # Verificar que la columna para el nombre del archivo exista en el Excel
    if columna_nombre_archivo not in df.columns:
        st.error(f"Error: La columna '{columna_nombre_archivo}' no se encuentra en el archivo Excel.")
        st.info(f"Columnas disponibles: {', '.join(df.columns)}")
        return None

    # Crear un objeto de BytesIO para guardar el archivo zip en memoria
    zip_buffer = BytesIO()

    # Usar un bloque 'with' para asegurarse de que el archivo zip se cierre correctamente
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
        total_documentos = len(df)
        progress_bar = st.progress(0, text="Iniciando generación...")

        # Iterar sobre cada fila del DataFrame
        for indice, fila in df.iterrows():
            # Cargar la plantilla desde los bytes en cada iteración
            # Esto es necesario para que cada documento sea independiente
            plantilla = DocxTemplate(BytesIO(plantilla_bytes))
            
            # Convertir la fila a un diccionario para el contexto
            contexto = fila.to_dict()
            
            # Renderizar la plantilla con los datos
            plantilla.render(contexto)
            
            # Crear un buffer en memoria para guardar el documento Word
            doc_buffer = BytesIO()
            plantilla.save(doc_buffer)
            doc_buffer.seek(0)
            
            # Crear el nombre del archivo de salida
            nombre_base = str(fila[columna_nombre_archivo]).replace('/', '_').replace('\\', '_')
            nombre_archivo_salida = f"{nombre_base}.docx"
            
            # Escribir el documento generado en el archivo zip
            zip_file.writestr(nombre_archivo_salida, doc_buffer.getvalue())

            # Actualizar la barra de progreso
            progreso_actual = (indice + 1) / total_documentos
            progress_bar.progress(progreso_actual, text=f"Generando documento: {nombre_archivo_salida} ({indice + 1}/{total_documentos})")

    progress_bar.empty() # Limpiar la barra de progreso al final
    zip_buffer.seek(0)
    return zip_buffer


# --- Interfaz de la Aplicación ---

st.title("📄 Ensamblador de Documentos")
st.markdown("Sube tu plantilla de Word y tu base de datos de Excel para generar documentos personalizados.")

# 1. Carga de la plantilla de Word
st.header("1. Sube la Plantilla de Word (.docx)")
archivo_plantilla = st.file_uploader("Selecciona tu archivo de plantilla", type=["docx"])

# 2. Carga del archivo de Excel
st.header("2. Sube la Base de Datos (.xlsx)")
archivo_excel = st.file_uploader("Selecciona tu archivo Excel con los datos", type=["xlsx"])

# 3. Ingreso del nombre de la columna
st.header("3. Define el Nombre de los Archivos")
columna_nombre_archivo = st.text_input(
    "Escribe el nombre exacto de la columna que usarás para nombrar cada documento",
    help="Ej: 'ItemId', 'Numero_Factura', 'Nombre_Cliente'. Debe coincidir con el encabezado en Excel."
)

# 4. Botón para generar los documentos
st.header("4. Genera los Documentos")

if st.button("✨ Generar Documentos", type="primary"):
    # Validaciones iniciales
    if archivo_plantilla and archivo_excel and columna_nombre_archivo:
        with st.spinner("Procesando... Por favor, espera."):
            # Leer los bytes de los archivos subidos
            plantilla_bytes = archivo_plantilla.getvalue()
            excel_bytes = archivo_excel.getvalue()

            # Llamar a la función principal para generar el zip
            zip_resultante = generar_documentos(plantilla_bytes, excel_bytes, columna_nombre_archivo)

            if zip_resultante:
                st.success("¡Documentos generados con éxito!")
                
                # 5. Botón de descarga
                st.download_button(
                    label="📥 Descargar todos los documentos (.zip)",
                    data=zip_resultante,
                    file_name="documentos_generados.zip",
                    mime="application/zip"
                )
    else:
        st.warning("Por favor, asegúrate de subir ambos archivos y de especificar el nombre de la columna.")

# Pie de página
st.markdown("---")
st.write("Aplicación creada para automatizar la generación de documentos.")
