import streamlit as st
import pandas as pd
import numpy as np
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from num2words import num2words
from io import BytesIO
import os

# Configuración inicial
st.set_page_config(page_title="Generador PQRS Convocatorias", layout="wide")

# Título principal
st.title("📄 Generador de PQRS para Convocatorias de Línea Pregrado")
st.subheader("Sapiencia - Medellín")

# Función para formatear números
def formato_numero(n):
    try:
        n = float(n)
        if n.is_integer():
            n = int(n)
        texto = num2words(n, lang='es')
        return f"{texto} ({n})"
    except (TypeError, ValueError):
        return n

# Carga de datos desde archivo Parquet interno
@st.cache_data
def cargar_datos():
    # Ruta interna del archivo Parquet
    ruta_parquet = r"C:\Users\Usuario\Documents\VSCode\Pqrs_26\Resultados_Linea_pregrado_2026-1.parquet"
   
    try:
        df = pd.read_parquet(ruta_parquet)
        df.fillna(0, inplace=True)
        df['Nombre'] = df['Nombre'].astype(str).str.upper()
        df['Documento'] = df['Documento'].astype(str)  # Asegurar tipo string
        return df
    except Exception as e:
        st.error(f"Error al cargar la base de datos: {str(e)}")
        return None

# Procesamiento de documentos con radicado e imágenes
def generar_documento(tipo_documento, row, radicado, imagen1=None, imagen2=None):
    # Preprocesar campos numéricos
    context = row.to_dict()
    for key in context:
        if key.startswith('cal'):
            context[key] = formato_numero(context[key])
    
    # Agregar radicado al contexto
    context['radicado'] = radicado
   
    # Seleccionar plantilla
    template_path = {
        "NO PRESELECCIONADO POR PUNTO DE CORTE PP": "No_preseleccionado_por_punto_corte_pp.docx",
        "NO CUMPLE HABILITANTE ART.70 LITERAL B": "No_cumple_habilitante_b.docx",
        "IMPEDIDO ART. 71 LITERAL A": "Impedido_literal_a.docx",
        "IMPEDIDO ART. 71 LITERAL C": "Impedido_literal_c.docx",
    }[tipo_documento]
   
    # Cargar plantilla
    doc = DocxTemplate(template_path)
    
    # Procesar imágenes si existen
    if imagen1 is not None:
        # Convertir imagen a formato compatible con docxtpl
        img1_stream = BytesIO(imagen1.getvalue())
        img1 = InlineImage(doc, img1_stream, width=Mm(120))  # Ajustar tamaño según necesidad
        context['imagen1'] = img1
    
    # Solo procesar imagen2 si es la plantilla "NO PRESELECCIONADO POR PUNTO DE CORTE PP"
    if tipo_documento == "NO PRESELECCIONADO POR PUNTO DE CORTE PP" and imagen2 is not None:
        # Convertir imagen a formato compatible con docxtpl
        img2_stream = BytesIO(imagen2.getvalue())
        img2 = InlineImage(doc, img2_stream, width=Mm(120))  # Ajustar tamaño según necesidad
        context['imagen2'] = img2
    else:
        # Para las otras plantillas, asegurarse de que imagen2 no esté en el contexto
        if 'imagen2' in context:
            context['imagen2'] = ""
   
    # Renderizar documento
    doc.render(context)
   
    # Preparar archivo para descarga
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
   
    return buffer

# Cargar datos automáticamente
df = cargar_datos()

if df is not None:
    st.success(f"Base de datos cargada internamente con {len(df)} registros")
    
    # Búsqueda por documento
    doc_busqueda = st.text_input("Ingrese el número de documento a buscar:")
    resultado = df[df['Documento'] == doc_busqueda] if doc_busqueda else pd.DataFrame()

    if not resultado.empty:
        row = resultado.iloc[0]
        st.success(f"Aspirante encontrado: {row['Nombre']}")
        
        # Mostrar información básica
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"**Documento:** {row['Documento']}")
        with col2:
            st.info(f"**Comuna:** {row['Comuna']}")
        with col2:
            st.info(f"**Estrato:** {row['Estrato']}")
        with col3:
            st.info(f"**Puntaje total:** {row['cal_total']}")
        with col3:
            st.info(f"**Puntaje de corte PP:** {row['punto_corte_pp']}")
        with col3:
            st.info(f"**RESULTADO CONVOCATORIA PP:** {row['Observaciones Presupuesto Participativo']}")
        
        # Sección para ingreso del radicado
        st.divider()
        st.subheader("📋 Información de la PQR")
        
        radicado = st.text_input(
            "Ingrese el número de radicado de la PQR:",
            placeholder="Ej: 2025-1244797-1",
            help="Este número aparecerá en el documento generado como {{radicado}}"
        )
        
        # Selección de documento a generar
        tipo_documento = st.selectbox(
            "Seleccione el tipo de documento a generar:",
            ["NO PRESELECCIONADO POR PUNTO DE CORTE PP", "NO CUMPLE HABILITANTE ART.70 LITERAL B",
             "IMPEDIDO ART. 71 LITERAL A",
             "IMPEDIDO ART. 71 LITERAL C"]
        )
        
        # Sección para carga de imágenes (opcional)
        st.divider()
        st.subheader("🖼️ Adjuntar imagenes en la PQRSDF")
        
        # Determinar qué imágenes solicitar según el tipo de documento
        if tipo_documento == "NO PRESELECCIONADO POR PUNTO DE CORTE PP":
            st.info("💡 Para este tipo de documento, se requieren dos imágenes como evidencia")
            col_img1, col_img2 = st.columns(2)
            
            with col_img1:
                imagen1 = st.file_uploader(
                    "Adjunte imagen del puntaje clúster:",
                    type=['png', 'jpg', 'jpeg', 'bmp', 'gif'],
                    help="Suba una imagen como evidencia complementaria"
                )
                if imagen1:
                    st.image(imagen1, caption="Vista previa imagen del puntaje clúster:", width=200)
            
            with col_img2:
                imagen2 = st.file_uploader(
                    "Adjunte imagen del puntaje total:",
                    type=['png', 'jpg', 'jpeg', 'bmp', 'gif'],
                    help="Suba una segunda imagen como evidencia complementaria"
                )
                if imagen2:
                    st.image(imagen2, caption="Vista previa imagen del puntaje total", width=200)
        
        else:
            # Para las demás plantillas, solo una imagen
            if tipo_documento in ["IMPEDIDO ART. 71 LITERAL A",
                                   "NO CUMPLE HABILITANTE ART.70 LITERAL B", 
                                  "IMPEDIDO ART. 71 LITERAL C"]:
                st.info("💡 Para este tipo de documento, se recomienda adjuntar evidencia de la situación")
            
            imagen1 = st.file_uploader(
                "Adjunte imagen de evidencia:",
                type=['png', 'jpg', 'jpeg', 'bmp', 'gif'],
                help="Suba una imagen como evidencia (opcional)"
            )
            if imagen1:
                st.image(imagen1, caption="Vista previa de la evidencia", width=200)
            
            # Para las otras plantillas, no hay imagen2
            imagen2 = None
        
        # Generar documento con validación
        st.divider()
        
        if st.button("📄 Generar Documento", type="primary"):
            if not radicado.strip():
                st.error("❌ Por favor ingrese el número de radicado antes de generar el documento.")
            else:
                try:
                    with st.spinner("Generando documento..."):
                        buffer = generar_documento(tipo_documento, row, radicado, imagen1, imagen2)
                        nombre_doc = f"{tipo_documento.replace(' ', '_')}-{row['Documento']}-{row['Nombre'][:30]}.docx"
                    
                    st.success("✅ Documento generado exitosamente!")
                    
                    st.download_button(
                        label="⬇️ Descargar Documento",
                        data=buffer,
                        file_name=nombre_doc,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        # icon="📥"
                    )
                    
                    # Mostrar resumen de lo generado
                    if tipo_documento == "NO PRESELECCIONADO POR PUNTO DE CORTE PP":
                        evidencias_text = f"Evidencias adjuntas: {bool(imagen1)} clúster, {bool(imagen2)} total"
                    else:
                        evidencias_text = f"Evidencia adjunta: {bool(imagen1)}"
                    
                    st.info(f"""
                    **Resumen del documento generado:**
                    - Tipo: {tipo_documento}
                    - Aspirante: {row['Nombre']}
                    - Documento: {row['Documento']}
                    - Radicado PQR: {radicado}
                    - {evidencias_text}
                    """)
                    
                except Exception as e:
                    st.error(f"❌ Error al generar el documento: {str(e)}")
                    st.info("⚠️ Asegúrese de que las plantillas tengan los marcadores correctos para radicado e imágenes.")
    
    elif doc_busqueda:
        st.warning("⚠️ No se encontró ningún aspirante con ese documento")