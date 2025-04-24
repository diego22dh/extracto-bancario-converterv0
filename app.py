import re
import pandas as pd
import PyPDF2
import pdfplumber
import os
from datetime import datetime
import streamlit as st
import tempfile
from io import BytesIO

def extraer_texto_de_pdf(archivo_pdf, metodo="pdfplumber"):
    """Extrae el texto completo de un archivo PDF usando el método especificado"""
    if metodo == "pdfplumber":
        try:
            texto_completo = ""
            with pdfplumber.open(archivo_pdf) as pdf:
                for pagina in pdf.pages:
                    texto_completo += pagina.extract_text() + "\n"
            return texto_completo
        except Exception as e:
            st.error(f"Error al leer el PDF con pdfplumber: {e}")
            return None
    else:  # PyPDF2 como respaldo
        try:
            texto_completo = ""
            with open(archivo_pdf, 'rb') as archivo:
                lector_pdf = PyPDF2.PdfReader(archivo)
                for pagina in lector_pdf.pages:
                    texto_completo += pagina.extract_text() + "\n"
            return texto_completo
        except Exception as e:
            st.error(f"Error al leer el PDF con PyPDF2: {e}")
            return None

def limpiar_valor_numerico(valor_str):
    """Limpia y convierte un valor numérico de texto a float"""
    if not valor_str:
        return 0.0
    
    # Reemplazar puntos como separadores de miles y comas como separadores decimales
    valor_limpio = valor_str.replace(".", "").replace(",", ".")
    try:
        return float(valor_limpio)
    except ValueError:
        return 0.0

def procesar_descripcion(descripcion_raw):
    """Procesa la descripción para separar la descripción principal del detalle"""
    descripcion = descripcion_raw.strip()
    detalle = ""
    
    # Si hay un guion, intentar separar en descripción y detalle
    if "-" in descripcion and descripcion.count("-") >= 1:
        partes = descripcion.split("-", 1)
        descripcion = partes[0].strip()
        if len(partes) > 1:
            detalle = partes[1].strip()
    
    return descripcion, detalle

def procesar_extracto_provincia(texto):
    """Procesa el texto del extracto del Banco Provincia y extrae las transacciones"""
    # Patrón para encontrar líneas de transacciones
    # Formato: DD-MM-YY DESCRIPCION IMPORTE DD-MM SALDO
    patron = r"(\d{2}-\d{2}-\d{2})\s+(.*?)\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)\s+(\d{2}-\d{2})\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)"
    
    # Patrón alternativo para líneas sin fecha (continuación de descripciones)
    patron_alt = r"^(\s+)(.+?)\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)\s+(\d{2}-\d{2})\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)"
    
    transacciones = []
    lineas = texto.split('\n')
    ultima_fecha = None
    
    for linea in lineas:
        # Intenta encontrar líneas de transacciones principales
        coincidencia = re.search(patron, linea)
        if coincidencia:
            fecha, descripcion, importe, fecha_valor, saldo = coincidencia.groups()
            
            # Limpia la descripción eliminando espacios múltiples
            descripcion = re.sub(r'\s+', ' ', descripcion.strip())
            
            # Extrae detalle (si existe)
            descripcion, detalle = procesar_descripcion(descripcion)
            
            # Procesa el importe para determinar si es débito o crédito
            importe_num = limpiar_valor_numerico(importe)
            tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
            # Para débitos, verificamos si ya tiene signo negativo
            if tipo_movimiento == "Débito" and importe_num > 0:
                importe_num = -importe_num
            
            # Procesa el saldo
            saldo_num = limpiar_valor_numerico(saldo)
            
            # Guarda la fecha actual para posibles continuaciones
            ultima_fecha = fecha
            
            transacciones.append({
                "fecha": fecha,
                "descripcion": descripcion,
                "detalle": detalle,
                "importe": importe_num,
                "saldo": saldo_num,
                "tipo_movimiento": tipo_movimiento
            })
        else:
            # Busca líneas alternativas (sin fecha)
            coincidencia_alt = re.search(patron_alt, linea)
            if coincidencia_alt and ultima_fecha:
                espacios, descripcion, importe, fecha_valor, saldo = coincidencia_alt.groups()
                
                # Limpia la descripción
                descripcion = re.sub(r'\s+', ' ', descripcion.strip())
                
                # Extrae detalle (si existe)
                descripcion, detalle = procesar_descripcion(descripcion)
                
                # Procesa el importe para determinar si es débito o crédito
                importe_num = limpiar_valor_numerico(importe)
                tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
                # Para débitos, verificamos si ya tiene signo negativo
                if tipo_movimiento == "Débito" and importe_num > 0:
                    importe_num = -importe_num
                
                # Procesa el saldo
                saldo_num = limpiar_valor_numerico(saldo)
                
                transacciones.append({
                    "fecha": ultima_fecha,
                    "descripcion": descripcion,
                    "detalle": detalle,
                    "importe": importe_num,
                    "saldo": saldo_num,
                    "tipo_movimiento": tipo_movimiento
                })
    
    return transacciones

def procesar_extracto_galicia(texto):
    """Procesa el texto del extracto del Banco Galicia y extrae las transacciones"""
    movimientos = []
    
    # Dividimos el texto por líneas
    lineas = texto.split('\n')
    inicio_movimientos = False
    
    for linea in lineas:
        # Detectamos el inicio de la tabla de movimientos
        if "Fecha" in linea and "Descripción" in linea and ("Crédito" in linea or "Débito" in linea) and "Saldo" in linea:
            inicio_movimientos = True
            continue
        
        # Si ya estamos en la sección de movimientos
        if inicio_movimientos:
            # Patrón para verificar si la línea comienza con una fecha
            patron_fecha = r"(\d{2}/\d{2}/\d{2})"
            
            # Si la línea comienza con una fecha, es un nuevo movimiento
            if re.match(patron_fecha, linea):
                # Patrones para diferentes formatos de línea
                # Patrón más completo
                patron1 = r"(\d{2}/\d{2}/\d{2})\s+(.*?)(?:\s+(\w+))?\s+(?:(\d+[\.,]?\d*\.?\d*,\d+)?\s+)?(?:(-\d+[\.,]?\d*\.?\d*,\d+)?\s+)?(\d+[\.,]?\d*\.?\d*,\d+)$"
                # Patrón alternativo para casos específicos
                patron2 = r"(\d{2}/\d{2}/\d{2})\s+(.*?)\s+(\d+[\.,]?\d*\.?\d*,\d+|\-\d+[\.,]?\d*\.?\d*,\d+)\s+(\d+[\.,]?\d*\.?\d*,\d+)$"
                
                match = re.search(patron1, linea)
                
                if match:
                    fecha, descripcion_raw, origen, credito, debito, saldo = match.groups()
                    
                    # Procesamos la descripción
                    descripcion, detalle = procesar_descripcion(descripcion_raw)
                    
                    # Determinamos tipo de movimiento e importe
                    if credito and credito.strip():
                        importe = limpiar_valor_numerico(credito)
                        tipo_movimiento = "Crédito"
                    elif debito and debito.strip():
                        # Mantenemos el signo negativo en el importe para débitos
                        importe = limpiar_valor_numerico(debito)  # Ya mantiene el signo negativo
                        tipo_movimiento = "Débito"
                    else:
                        # Intentar un enfoque alternativo para casos especiales
                        match_alt = re.search(patron2, linea)
                        if match_alt:
                            fecha, descripcion_raw, importe_str, saldo = match_alt.groups()
                            descripcion, detalle = procesar_descripcion(descripcion_raw)
                            
                            # Procesamos el importe manteniendo el signo
                            importe = limpiar_valor_numerico(importe_str)
                            tipo_movimiento = "Débito" if importe < 0 else "Crédito"
                        else:
                            importe = 0.0
                            tipo_movimiento = "Desconocido"
                    
                    # Procesamos el saldo
                    saldo_valor = limpiar_valor_numerico(saldo) if saldo else 0.0
                    
                    # Agregamos el movimiento a la lista
                    movimientos.append({
                        'fecha': fecha,
                        'descripcion': descripcion,
                        'detalle': detalle,
                        'importe': importe,
                        'saldo': saldo_valor,
                        'tipo_movimiento': tipo_movimiento
                    })
                else:
                    # Intentamos con el patrón alternativo
                    match_alt = re.search(patron2, linea)
                    if match_alt:
                        fecha, descripcion_raw, importe_str, saldo = match_alt.groups()
                        descripcion, detalle = procesar_descripcion(descripcion_raw)
                        
                        # Procesamos el importe manteniendo el signo
                        importe = limpiar_valor_numerico(importe_str)
                        tipo_movimiento = "Débito" if importe < 0 else "Crédito"
                        
                        # Procesamos el saldo
                        saldo_valor = limpiar_valor_numerico(saldo) if saldo else 0.0
                        
                        # Agregamos el movimiento a la lista
                        movimientos.append({
                            'fecha': fecha,
                            'descripcion': descripcion,
                            'detalle': detalle,
                            'importe': importe,
                            'saldo': saldo_valor,
                            'tipo_movimiento': tipo_movimiento
                        })
    
    return movimientos

def guardar_excel(transacciones):
    """Convierte las transacciones en un DataFrame para mostrar y descargar"""
    if not transacciones:
        st.error("No se encontraron transacciones para guardar")
        return None
    
    # Crea un DataFrame con los datos
    df = pd.DataFrame(transacciones)
    
    # Reformatea la fecha si es necesario (para asegurar formato uniforme)
    def reformatear_fecha(fecha_str):
        try:
            # Intenta varios formatos de fecha
            for formato in ["%d-%m-%y", "%d/%m/%Y", "%d/%m/%y"]:
                try:
                    return datetime.strptime(fecha_str, formato).strftime("%d/%m/%Y")
                except ValueError:
                    continue
            return fecha_str
        except Exception:
            return fecha_str
    
    df["fecha"] = df["fecha"].apply(reformatear_fecha)
    
    # Ordena por fecha (descendente) y saldo
    df = df.sort_values(by=["fecha", "saldo"], ascending=[False, False])
    
    return df

# Configurar la página de Streamlit
st.set_page_config(page_title="Procesador de Extractos Bancarios", layout="wide")

# Título de la aplicación
st.title("Procesador de Extractos Bancarios")

# Crear un sidebar para las opciones
st.sidebar.header("Configuración")

# Selección del tipo de banco
banco = st.sidebar.selectbox(
    "Seleccione el banco",
    options=["Provincia", "Galicia"],
    index=0
)

# Subida de archivo
archivo_pdf = st.sidebar.file_uploader("Seleccione el extracto bancario en PDF", type=["pdf"])

# Inicializar la sesión si no existe
if 'procesado' not in st.session_state:
    st.session_state.procesado = False
    st.session_state.transacciones = None
    st.session_state.df = None

# Botón para procesar
if st.sidebar.button("Procesar Extracto"):
    if archivo_pdf is not None:
        # Notificar el inicio del procesamiento
        with st.spinner(f"Procesando extracto del Banco {banco}..."):
            # Crear un archivo temporal para el PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(archivo_pdf.getvalue())
                temp_pdf_path = temp_pdf.name
                
            # Extraer texto del PDF
            texto = extraer_texto_de_pdf(temp_pdf_path, "pdfplumber")
            
            if not texto:
                st.error("Error al extraer texto con pdfplumber, intentando con PyPDF2...")
                texto = extraer_texto_de_pdf(temp_pdf_path, "pypdf2")
            
            if not texto:
                st.error("No se pudo extraer texto del PDF.")
            else:
                # Procesar según el banco seleccionado
                if banco == "Provincia":
                    transacciones = procesar_extracto_provincia(texto)
                else:  # banco == "Galicia"
                    transacciones = procesar_extracto_galicia(texto)
                
                # Limpiar el archivo temporal
                os.unlink(temp_pdf_path)
                
                if not transacciones:
                    st.error(f"No se encontraron transacciones en el extracto del Banco {banco}.")
                else:
                    st.session_state.procesado = True
                    st.session_state.transacciones = transacciones
                    st.session_state.df = guardar_excel(transacciones)
                    st.success(f"Se encontraron {len(transacciones)} transacciones.")
    else:
        st.error("Por favor, suba un archivo PDF primero.")

# Mostrar resultados si hay transacciones procesadas
if st.session_state.procesado and st.session_state.df is not None:
    st.header("Transacciones Encontradas")
    
    # Mostrar DataFrame
    st.dataframe(st.session_state.df)
        # Crear un objeto BytesIO para guardar el Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df.to_excel(writer, index=False)
    output.seek(0)

    st.download_button(
        label="Descargar Excel",
        data=output,
        file_name=f"extracto_{banco.lower()}_procesado.xlsx",
        mime="application/vnd.ms-excel"
    )
    
    
else:
    # Instrucciones iniciales
    st.info("Sube un extracto bancario en PDF y selecciona el banco para comenzar.")