import streamlit as st
import PyPDF2
import pandas as pd
import re
import io
from datetime import datetime

def extraer_texto_de_pdf(archivo_pdf):
    """Extrae el texto completo de un archivo PDF subido por Streamlit"""
    texto_completo = ""
    try:
        lector_pdf = PyPDF2.PdfReader(archivo_pdf)
        for pagina in lector_pdf.pages:
            texto_completo += pagina.extract_text()
        return texto_completo
    except Exception as e:
        st.error(f"Error al leer el PDF: {e}")
        return None

def procesar_extracto_provincia(texto):
    """Procesa el texto del extracto del Banco Provincia y extrae las transacciones"""
    patron = r"(\d{2}-\d{2}-\d{2})\s+(.*?)\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)\s+(\d{2}-\d{2})\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)"
    patron_alt = r"^(\s+)(.+?)\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)\s+(\d{2}-\d{2})\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)"
    
    transacciones = []
    lineas = texto.split('\n')
    ultima_fecha = None
    
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue
            
        coincidencia = re.search(patron, linea)
        if coincidencia:
            fecha, descripcion, importe, _, saldo = coincidencia.groups()
            ultima_fecha = fecha
            transacciones.append({
                "fecha": fecha,
                "descripcion": descripcion.strip(),
                "importe": float(importe.replace(",", ".")),
                "saldo": float(saldo.replace(",", "."))
            })
            
    return transacciones

def procesar_extracto_galicia(texto):
    """Procesa el texto del extracto del Banco Galicia y extrae las transacciones"""
    # Patrón más flexible para capturar diferentes formatos de líneas
    patron = r"(\d{2}/\d{2}/(?:\d{2}|\d{4}))\s+(.*?)\s+([-]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s*([-]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)"
    
    transacciones = []
    lineas = texto.split('\n')
    ultima_fecha = None
    descripcion_completa = ""
    
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue
            
        # Intenta encontrar líneas de transacciones
        coincidencia = re.search(patron, linea)
        
        if coincidencia:
            fecha, descripcion, importe, saldo = coincidencia.groups()
            
            # Guarda la última fecha válida
            ultima_fecha = fecha
            
            # Combina con la descripción acumulada si existe
            if descripcion_completa:
                descripcion = f"{descripcion_completa} {descripcion}"
                descripcion_completa = ""
            
            # Normalizar fecha
            if len(fecha.split('/')[2]) == 2:
                fecha = fecha.replace('/', '-')
            
            # Limpia y normaliza la descripción
            descripcion = re.sub(r'\s+', ' ', descripcion.strip())
            
            # Procesa importes
            importe = importe.replace(".", "").replace(",", ".")
            saldo = saldo.replace(".", "").replace(",", ".")
            
            try:
                importe_num = float(importe)
                saldo_num = float(saldo)
                tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
                
                # Extrae detalle si existe
                detalle = ""
                if " - " in descripcion:
                    partes = descripcion.split(" - ", 1)
                    descripcion = partes[0].strip()
                    detalle = partes[1].strip() if len(partes) > 1 else ""
                
                transacciones.append({
                    "fecha": fecha,
                    "descripcion": descripcion,
                    "detalle": detalle,
                    "importe": importe_num,
                    "saldo": saldo_num,
                    "tipo_movimiento": tipo_movimiento
                })
            except ValueError:
                continue
        else:
            # Si la línea no coincide con el patrón pero tiene contenido,
            # podría ser una descripción adicional
            if ultima_fecha and linea:
                descripcion_completa += " " + linea
    
    return transacciones

def guardar_excel(transacciones, ruta_salida):
    """Guarda las transacciones en un archivo Excel"""
    if not transacciones:
        print("No se encontraron transacciones para guardar")
        return False
    
    # Crea un DataFrame con los datos
    df = pd.DataFrame(transacciones)
    
    # Reformatea la fecha si es necesario (para asegurar formato uniforme)
    def reformatear_fecha(fecha_str):
        try:
            # Intenta varios formatos de fecha
            for formato in ["%d-%m-%y", "%d/%m/%Y"]:
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
    
    # Guarda el DataFrame en un archivo Excel
    df.to_excel(ruta_salida, index=False)
    print(f"Archivo Excel guardado exitosamente en: {ruta_salida}")
    return True

def main():
    st.title("Conversor de Extractos Bancarios")
    st.write("Convierte extractos bancarios PDF a Excel")

    # Agregar key única al file_uploader
    uploaded_file = st.file_uploader(
        "Sube el extracto bancario en PDF", 
        type=["pdf"],
        key="extracto_bancario_upload"
    )
    banco = st.selectbox(
        "Selecciona el banco", 
        ["Provincia", "Galicia"],
        key="selector_banco"
    ).lower()

    if uploaded_file is not None:
        texto = extraer_texto_de_pdf(uploaded_file)
        
        if texto:
            # Procesar según el banco
            transacciones = procesar_extracto_provincia(texto) if banco == "provincia" else procesar_extracto_galicia(texto)
            
            if transacciones:
                # Crear DataFrame
                df = pd.DataFrame(transacciones)
                
                # Mostrar vista previa
                st.write("### Vista previa de las transacciones:")
                st.dataframe(df)
                
                # Preparar Excel para descarga
                buffer = io.BytesIO()
                
                # Usar ExcelWriter con formato específico
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Transacciones', index=False)
                    # Obtener la hoja activa
                    worksheet = writer.sheets['Transacciones']
                    # Ajustar anchos de columna
                    for idx, col in enumerate(df.columns):
                        max_length = max(df[col].astype(str).apply(len).max(),
                                       len(col)) + 2
                        worksheet.column_dimensions[chr(65 + idx)].width = max_length

                # Crear botón de descarga
                nombre_excel = f"extracto_{banco}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                
                st.download_button(
                    label="📥 Descargar Excel",
                    data=buffer.getvalue(),
                    file_name=nombre_excel,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Haz clic para descargar el archivo Excel con las transacciones procesadas"
                )
                
                st.success(f"✅ Archivo {nombre_excel} listo para descargar")
            else:
                st.error(f"❌ No se encontraron transacciones en el extracto del Banco {banco.capitalize()}")
        else:
            st.error("❌ No se pudo extraer texto del PDF")

if __name__ == "__main__":
    main()