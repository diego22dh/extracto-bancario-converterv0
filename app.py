import re
import pandas as pd
import PyPDF2
import os
from datetime import datetime
import streamlit as st
import io

def extraer_texto_de_pdf(ruta_pdf):
    """Extrae el texto completo de un archivo PDF"""
    texto_completo = ""
    try:
        with open(ruta_pdf, 'rb') as archivo:
            lector_pdf = PyPDF2.PdfReader(archivo)
            for pagina in lector_pdf.pages:
                texto_completo += pagina.extract_text()
        return texto_completo
    except Exception as e:
        print(f"Error al leer el PDF: {e}")
        return None

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
            
            # Procesa el importe para determinar si es débito o crédito
            importe = importe.replace(".", "").replace(",", ".")
            try:
                importe_num = float(importe)
                tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
                # Para débitos, verificamos si ya tiene signo negativo
                if tipo_movimiento == "Débito" and not importe.startswith('-'):
                    importe_num = -importe_num
            except ValueError:
                importe_num = 0
                tipo_movimiento = "Indeterminado"
            importe = importe.replace(".", "").replace(",", ".")
            try:
                importe_num = float(importe)
                tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
                # Para débitos, verificamos si ya tiene signo negativo
                if tipo_movimiento == "Débito" and not importe.startswith('-'):
                    importe_num = -importe_num
            except ValueError:
                importe_num = 0
                tipo_movimiento = "Indeterminado"
            
            # Procesa el saldo
            saldo = saldo.replace(".", "").replace(",", ".")
            saldo = saldo.replace(".", "").replace(",", ".")
            
            # Guarda la fecha actual para posibles continuaciones
            ultima_fecha = fecha
            
            # Extrae detalle (si existe)
            detalle = ""
            if "-" in descripcion and descripcion.count("-") >= 1:
                partes = descripcion.split("-", 1)
                descripcion = partes[0].strip()
                if len(partes) > 1:
                    detalle = partes[1].strip()
            
            # Extrae detalle (si existe)
            detalle = ""
            if "-" in descripcion and descripcion.count("-") >= 1:
                partes = descripcion.split("-", 1)
                descripcion = partes[0].strip()
                if len(partes) > 1:
                    detalle = partes[1].strip()
            
            transacciones.append({
                "fecha": fecha,
                "descripcion": descripcion,
                "detalle": detalle,
                "importe": importe_num,
                "saldo": float(saldo),
                "tipo_movimiento": tipo_movimiento
            })
        else:
            # Busca líneas alternativas (sin fecha)
            coincidencia_alt = re.search(patron_alt, linea)
            if coincidencia_alt and ultima_fecha:
                espacios, descripcion, importe, fecha_valor, saldo = coincidencia_alt.groups()
                
                # Limpia la descripción
                descripcion = re.sub(r'\s+', ' ', descripcion.strip())
                
                # Procesa el importe para determinar si es débito o crédito
                importe = importe.replace(".", "").replace(",", ".")
                try:
                    importe_num = float(importe)
                    tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
                    # Para débitos, verificamos si ya tiene signo negativo
                    if tipo_movimiento == "Débito" and not importe.startswith('-'):
                        importe_num = -importe_num
                except ValueError:
                    importe_num = 0
                    tipo_movimiento = "Indeterminado"
                
                # Procesa el saldo
                saldo = saldo.replace(".", "").replace(",", ".")
                
                # Extrae detalle (si existe)
                detalle = ""
                if "-" in descripcion and descripcion.count("-") >= 1:
                    partes = descripcion.split("-", 1)
                    descripcion = partes[0].strip()
                    if len(partes) > 1:
                        detalle = partes[1].strip()
                
                transacciones.append({
                    "fecha": ultima_fecha,
                    "descripcion": descripcion,
                    "detalle": detalle,
                    "importe": importe_num,
                    "saldo": float(saldo),
                    "tipo_movimiento": tipo_movimiento
                })
                "saldo": float(saldo),
                "tipo_movimiento": tipo_movimiento
            })
        else:
            # Busca líneas alternativas (sin fecha)
            coincidencia_alt = re.search(patron_alt, linea)
            if coincidencia_alt and ultima_fecha:
                espacios, descripcion, importe, fecha_valor, saldo = coincidencia_alt.groups()
                
                # Limpia la descripción
                descripcion = re.sub(r'\s+', ' ', descripcion.strip())
                
                # Procesa el importe para determinar si es débito o crédito
                importe = importe.replace(".", "").replace(",", ".")
                try:
                    importe_num = float(importe)
                    tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
                    # Para débitos, verificamos si ya tiene signo negativo
                    if tipo_movimiento == "Débito" and not importe.startswith('-'):
                        importe_num = -importe_num
                except ValueError:
                    importe_num = 0
                    tipo_movimiento = "Indeterminado"
                
                # Procesa el saldo
                saldo = saldo.replace(".", "").replace(",", ".")
                
                # Extrae detalle (si existe)
                detalle = ""
                if "-" in descripcion and descripcion.count("-") >= 1:
                    partes = descripcion.split("-", 1)
                    descripcion = partes[0].strip()
                    if len(partes) > 1:
                        detalle = partes[1].strip()
                
                transacciones.append({
                    "fecha": ultima_fecha,
                    "descripcion": descripcion,
                    "detalle": detalle,
                    "importe": importe_num,
                    "saldo": float(saldo),
                    "tipo_movimiento": tipo_movimiento
                })
    
    return transacciones

def procesar_extracto_galicia(texto):
    """Procesa el texto del extracto del Banco Galicia y extrae las transacciones"""
    # Patrón más flexible para el Banco Galicia
    patron = r"(\d{2}/\d{2}/\d{2,4})\s+(.*?)\s+([-]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)\s*([-]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?)"
    
    # Añadir debug
    st.write("### Debugging texto extraído:")
    st.text(texto[:500])  # Muestra los primeros 500 caracteres
    
    transacciones = []
    lineas = texto.split('\n')
    
    for linea in lineas:
        # Limpiar la línea
        linea = linea.strip()
        if not linea:
            continue
            
        # Intenta encontrar líneas de transacciones
        coincidencia = re.search(patron, linea)
        if coincidencia:
            fecha, descripcion, importe, saldo = coincidencia.groups()
            
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
                tipo_movimiento = "Crédito" if importe_num > 0 else "Débito"
            except ValueError:
                continue  # Salta líneas con importes inválidos
            
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
                "saldo": float(saldo),
                "tipo_movimiento": tipo_movimiento
            })
    
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
            for formato in ["%d-%m-%y", "%d/%m/%Y"]:
                try:
                    return datetime.strptime(fecha_str, formato).strftime("%d/%m/%Y")
                except:
                except:
                    continue
            return fecha_str
        except:
        except:
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

    # Subida del archivo y selector de banco
    uploaded_file = st.file_uploader("Sube el extracto bancario en PDF", type=["pdf"])
    banco = st.selectbox("Selecciona el banco", ["Provincia", "Galicia"]).lower()

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
    main()