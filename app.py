import re
import pandas as pd
import PyPDF2
import os
from datetime import datetime

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
            
            # Procesa el saldo
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
    
    return transacciones

def procesar_extracto_galicia(texto):
    """Procesa el texto del extracto del Banco Galicia y extrae las transacciones"""
    # Patrón para encontrar líneas de transacciones en el formato del Banco Galicia
    # Adapta este patrón según el formato real de los extractos del Galicia
    patron = r"(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)\s+([-]?\d+(?:\.\d+)?(?:,\d+)?)"
    
    transacciones = []
    lineas = texto.split('\n')
    
    for linea in lineas:
        # Intenta encontrar líneas de transacciones
        coincidencia = re.search(patron, linea)
        if coincidencia:
            fecha, descripcion, importe, saldo = coincidencia.groups()
            
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
                try:
                    return datetime.strptime(fecha_str, formato).strftime("%d/%m/%Y")
                except:
                    continue
            return fecha_str
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
    # Solicitar la ruta del archivo PDF
    ruta_pdf = input("Ingrese la ruta del archivo PDF del extracto bancario: ")
    
    if not os.path.exists(ruta_pdf):
        print(f"El archivo {ruta_pdf} no existe.")
        return
    
    # Solicitar el tipo de banco
    while True:
        banco = input("Seleccione el banco (Provincia/Galicia): ").strip().lower()
        if banco in ["provincia", "galicia"]:
            break
        print("Por favor, ingrese 'Provincia' o 'Galicia'.")
    
    # Extraer texto del PDF
    texto = extraer_texto_de_pdf(ruta_pdf)
    
    if not texto:
        print("No se pudo extraer texto del PDF.")
        return
    
    # Procesar el extracto según el banco seleccionado
    if banco == "provincia":
        transacciones = procesar_extracto_provincia(texto)
    else:  # banco == "galicia"
        transacciones = procesar_extracto_galicia(texto)
    
    if not transacciones:
        print(f"No se encontraron transacciones en el extracto del Banco {banco.capitalize()}.")
        return
    
    # Determinar la ruta de salida del archivo Excel
    nombre_base = os.path.splitext(os.path.basename(ruta_pdf))[0]
    ruta_salida = os.path.join(os.path.dirname(ruta_pdf), f"{nombre_base}_{banco}_procesado.xlsx")
    
    # Guardar en Excel
    guardar_excel(transacciones, ruta_salida)

if __name__ == "__main__":
    main()