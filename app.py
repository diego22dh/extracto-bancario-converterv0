import re
import pandas as pd
import PyPDF2
import pdfplumber
import os
from datetime import datetime

def extraer_texto_de_pdf(ruta_pdf, metodo="pdfplumber"):
    """Extrae el texto completo de un archivo PDF usando el método especificado"""
    if metodo == "pdfplumber":
        try:
            texto_completo = ""
            with pdfplumber.open(ruta_pdf) as pdf:
                for pagina in pdf.pages:
                    texto_completo += pagina.extract_text() + "\n"
            return texto_completo
        except Exception as e:
            print(f"Error al leer el PDF con pdfplumber: {e}")
            return None
    else:  # PyPDF2 como respaldo
        try:
            texto_completo = ""
            with open(ruta_pdf, 'rb') as archivo:
                lector_pdf = PyPDF2.PdfReader(archivo)
                for pagina in lector_pdf.pages:
                    texto_completo += pagina.extract_text() + "\n"
            return texto_completo
        except Exception as e:
            print(f"Error al leer el PDF con PyPDF2: {e}")
            return None

def extraer_tablas_con_pdfplumber(ruta_pdf):
    """Extrae tablas específicamente usando pdfplumber para mejor manejo de estructuras tabulares"""
    todas_las_tablas = []
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                print(f"Procesando página {i+1}...")
                # Intentar extraer tablas de la página
                tablas = pagina.extract_tables()
                if tablas:
                    print(f"Se encontraron {len(tablas)} tablas en la página {i+1}")
                    for tabla in tablas:
                        todas_las_tablas.append((i+1, tabla))
                else:
                    # Si no hay tablas, intentar extraer texto estructurado
                    texto = pagina.extract_text()
                    if texto:
                        todas_las_tablas.append((i+1, texto))
        return todas_las_tablas
    except Exception as e:
        print(f"Error al extraer tablas con pdfplumber: {e}")
        return []

def limpiar_valor_numerico(valor_str):
    """Limpia y convierte un valor numérico de texto a float"""
    if not valor_str or valor_str.strip() == "":
        return 0.0
    
    # Remover espacios y caracteres no numéricos excepto puntos, comas y signos
    valor_limpio = str(valor_str).strip().replace(" ", "")
    
    # Manejar diferentes formatos de números
    # Si hay punto y coma, asumimos que el punto es separador de miles y la coma decimal
    if "." in valor_limpio and "," in valor_limpio:
        valor_limpio = valor_limpio.replace(".", "").replace(",", ".")
    # Si solo hay coma, probablemente es separador decimal
    elif "," in valor_limpio and "." not in valor_limpio:
        valor_limpio = valor_limpio.replace(",", ".")
    # Si solo hay puntos, verificamos si es separador de miles o decimal
    elif "." in valor_limpio:
        partes = valor_limpio.split(".")
        if len(partes[-1]) == 2:  # Probablemente decimal
            # Es decimal, no hacemos nada
            pass
        elif len(partes[-1]) == 3:  # Probablemente separador de miles
            valor_limpio = valor_limpio.replace(".", "")
    
    try:
        return float(valor_limpio)
    except ValueError:
        print(f"No se pudo convertir '{valor_str}' a número")
        return 0.0

def procesar_descripcion(descripcion_raw):
    """Procesa la descripción para separar la descripción principal del detalle"""
    if not descripcion_raw:
        return "", ""
    
    descripcion = str(descripcion_raw).strip()
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
                        importe = limpiar_valor_numerico(debito)
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
    
    return movimientos

def procesar_extracto_santander(ruta_pdf):
    """Procesa el extracto del Banco Santander extrayendo tablas directamente"""
    movimientos = []
    
    # Extraer tablas usando pdfplumber
    tablas_y_texto = extraer_tablas_con_pdfplumber(ruta_pdf)
    
    if not tablas_y_texto:
        print("No se encontraron tablas o texto en el PDF del Santander")
        return []
    
    for pagina_num, contenido in tablas_y_texto:
        print(f"Procesando contenido de página {pagina_num}...")
        
        # Si el contenido es una tabla (lista de listas)
        if isinstance(contenido, list) and len(contenido) > 0:
            # Buscar encabezados que indiquen una tabla de movimientos
            encabezados_encontrados = False
            indice_inicio = 0
            
            for i, fila in enumerate(contenido):
                if fila and any(col for col in fila if col and 
                               any(keyword in str(col).lower() for keyword in 
                                   ['fecha', 'descripcion', 'descripción', 'concepto', 'importe', 'saldo', 'débito', 'crédito'])):
                    encabezados_encontrados = True
                    indice_inicio = i + 1
                    print(f"Encabezados encontrados en fila {i}: {fila}")
                    break
            
            if encabezados_encontrados and indice_inicio < len(contenido):
                # Procesar las filas de datos
                for fila in contenido[indice_inicio:]:
                    if not fila or not any(col for col in fila if col and str(col).strip()):
                        continue  # Saltar filas vacías
                    
                    # Intentar identificar las columnas (esto puede necesitar ajuste según el formato real)
                    fecha, descripcion, detalle, importe, saldo, tipo_movimiento = extraer_datos_fila_santander(fila)
                    
                    if fecha:  # Solo agregar si tenemos al menos una fecha válida
                        movimientos.append({
                            'fecha': fecha,
                            'descripcion': descripcion,
                            'detalle': detalle,
                            'importe': importe,
                            'saldo': saldo,
                            'tipo_movimiento': tipo_movimiento
                        })
        
        # Si el contenido es texto plano
        elif isinstance(contenido, str):
            # Buscar patrones de movimientos en el texto
            movimientos_texto = procesar_texto_santander(contenido)
            movimientos.extend(movimientos_texto)
    
    return movimientos

def extraer_datos_fila_santander(fila):
    """Extrae los datos de una fila de tabla del Banco Santander"""
    fecha = ""
    descripcion = ""
    detalle = ""
    importe = 0.0
    saldo = 0.0
    tipo_movimiento = "Desconocido"
    
    # Limpiar la fila de valores None o vacíos
    fila_limpia = [str(col).strip() if col is not None else "" for col in fila]
    
    # Buscar fecha en la primera columna que tenga formato de fecha
    for i, col in enumerate(fila_limpia):
        if re.match(r'\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4}|\d{2}/\d{2}/\d{2}', col):
            fecha = col
            break
    
    # Buscar descripción (usualmente la columna más larga de texto)
    descripcion_candidatos = [col for col in fila_limpia if col and len(col) > 10 and not re.match(r'^[\d\.\,\-\s]+$', col)]
    if descripcion_candidatos:
        descripcion_raw = descripcion_candidatos[0]
        descripcion, detalle = procesar_descripcion(descripcion_raw)
    
    # Buscar valores numéricos (importe y saldo)
    valores_numericos = []
    for col in fila_limpia:
        if col and re.match(r'^[\d\.\,\-\s]+$', col):
            try:
                valor = limpiar_valor_numerico(col)
                if valor != 0.0:
                    valores_numericos.append(valor)
            except:
                continue
    
    # Asignar importe y saldo (el último suele ser el saldo)
    if len(valores_numericos) >= 2:
        importe = valores_numericos[-2]
        saldo = valores_numericos[-1]
        tipo_movimiento = "Crédito" if importe > 0 else "Débito"
    elif len(valores_numericos) == 1:
        saldo = valores_numericos[0]
    
    return fecha, descripcion, detalle, importe, saldo, tipo_movimiento

def procesar_texto_santander(texto):
    """Procesa texto plano del Santander buscando patrones de movimientos"""
    movimientos = []
    
    # Patrones comunes para extractos de Santander
    patrones = [
        r"(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([-]?\d+[.,]\d+)\s+([-]?\d+[.,]\d+)",
        r"(\d{2}/\d{2}/\d{2})\s+(.*?)\s+([-]?\d+[.,]\d+)\s+([-]?\d+[.,]\d+)",
        r"(\d{2}-\d{2}-\d{4})\s+(.*?)\s+([-]?\d+[.,]\d+)\s+([-]?\d+[.,]\d+)"
    ]
    
    lineas = texto.split('\n')
    
    for linea in lineas:
        for patron in patrones:
            match = re.search(patron, linea)
            if match:
                fecha, descripcion_raw, importe_str, saldo_str = match.groups()
                
                descripcion, detalle = procesar_descripcion(descripcion_raw)
                importe = limpiar_valor_numerico(importe_str)
                saldo = limpiar_valor_numerico(saldo_str)
                tipo_movimiento = "Crédito" if importe > 0 else "Débito"
                
                movimientos.append({
                    'fecha': fecha,
                    'descripcion': descripcion,
                    'detalle': detalle,
                    'importe': importe,
                    'saldo': saldo,
                    'tipo_movimiento': tipo_movimiento
                })
                break  # Salir del bucle de patrones una vez que encontramos uno
    
    return movimientos

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
            for formato in ["%d-%m-%y", "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y"]:
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
    print(f"Total de transacciones procesadas: {len(transacciones)}")
    return True

def main():
    # Solicitar la ruta del archivo PDF
    ruta_pdf = input("Ingrese la ruta del archivo PDF del extracto bancario: ")
    
    if not os.path.exists(ruta_pdf):
        print(f"El archivo {ruta_pdf} no existe.")
        return
    
    # Solicitar el tipo de banco
    while True:
        banco = input("Seleccione el banco (Provincia/Galicia/Santander): ").strip().lower()
        if banco in ["provincia", "galicia", "santander"]:
            break
        print("Por favor, ingrese 'Provincia', 'Galicia' o 'Santander'.")
    
    # Procesar el extracto según el banco seleccionado
    print(f"Procesando extracto del Banco {banco.capitalize()}...")
    
    if banco == "santander":
        # Para Santander, procesamos directamente desde el PDF
        transacciones = procesar_extracto_santander(ruta_pdf)
    else:
        # Para otros bancos, primero extraemos el texto
        texto = extraer_texto_de_pdf(ruta_pdf, "pdfplumber")
        
        if not texto:
            print("Error al extraer texto con pdfplumber, intentando con PyPDF2...")
            texto = extraer_texto_de_pdf(ruta_pdf, "pypdf2")
        
        if not texto:
            print("No se pudo extraer texto del PDF.")
            return
        
        if banco == "provincia":
            transacciones = procesar_extracto_provincia(texto)
        elif banco == "galicia":
            transacciones = procesar_extracto_galicia(texto)
    
    if not transacciones:
        print(f"No se encontraron transacciones en el extracto del Banco {banco.capitalize()}.")
        return
    
    print(f"Se encontraron {len(transacciones)} transacciones.")
    
    # Determinar la ruta de salida del archivo Excel
    nombre_base = os.path.splitext(os.path.basename(ruta_pdf))[0]
    ruta_salida = os.path.join(os.path.dirname(ruta_pdf), f"{nombre_base}_{banco}_procesado.xlsx")
    
    # Guardar en Excel
    guardar_excel(transacciones, ruta_salida)

if __name__ == "__main__":
    main() 