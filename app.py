import streamlit as st
import pdfplumber
import pandas as pd
import re
import base64
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Conversor de Extracto Bancario",
    page_icon="🏦",
    layout="centered"
)

st.title("Conversor de Extracto Bancario Galicia")
st.write("Convierte extractos bancarios del Banco Galicia a Excel. Detecta automáticamente si el PDF es formato "
         "**Office Banking** (fecha AAAA, importes con signo '$') o **Resumen de Cuenta Corriente** "
         "(fecha AA, columna Origen, importes sin '$').")

# --- Patrones formato "Office Banking" (fecha con año de 4 dígitos, importes "+ $ x,xx" / "- $ x,xx") ---
PATRON_FECHA_OB = re.compile(r'^(\d{2}/\d{2}/\d{4})\s+(.*)$')
PATRON_MONTO_OB = re.compile(r'([+-])\s*\$\s*([\d\.]+,\d{2})')
PATRON_SALDO_OB = re.compile(r'\$\s*([\d\.]+,\d{2})\s*$')

# --- Patrones formato "Resumen de Cuenta Corriente" (fecha con año de 2 dígitos, importes sin '$') ---
PATRON_FECHA_RC = re.compile(r'^(\d{2}/\d{2}/\d{2})\s+(.*)$')
PATRON_NUM_RC = re.compile(r'(-?\d{1,3}(?:\.\d{3})*,\d{2})')


def limpiar_valor_numerico(valor_str):
    """Convierte un número en formato argentino (1.234.567,89) a float, respetando el signo."""
    if not valor_str:
        return 0.0
    es_negativo = valor_str.strip().startswith('-')
    valor_limpio = valor_str.strip().lstrip('-').replace('.', '').replace(',', '.')
    try:
        resultado = float(valor_limpio)
    except ValueError:
        return 0.0
    return -resultado if es_negativo else resultado


def es_codigo_origen(token):
    """Códigos cortos tipo '00D5' de la columna Origen (formato Resumen de Cuenta), para no dejarlos pegados a la descripción."""
    return bool(re.fullmatch(r'[0-9A-Z]{2,6}', token)) and any(c.isdigit() for c in token) and any(c.isalpha() for c in token)


def es_linea_ruido(linea):
    """Encabezados, pies de página y numeración que no son parte de ningún movimiento (ambos formatos)."""
    ruido_exacto = {"Office Banking", "Galicia", "Resumen de Cuenta Corriente en Pesos"}
    if linea in ruido_exacto:
        return True
    if linea.startswith("Fecha de descarga"):
        return True
    if linea.startswith("Página"):
        return True
    # timestamp tipo "02/09/26 - 10:11hs" (Office Banking)
    if re.match(r'^\d{2}/\d{2}/\d{2}\s*-\s*\d{2}:\d{2}hs$', linea):
        return True
    # código de documento al pie de página (ej: 202608310468382697H)
    if re.match(r'^\d{15,}[A-Z]?$', linea):
        return True
    # numero de página suelto
    if re.match(r'^\d{1,2}$', linea):
        return True
    # fila de encabezado de la tabla
    if "Fecha" in linea and "Descripción" in linea and ("Débito" in linea or "Crédito" in linea) and "Saldo" in linea:
        return True
    return False


def extraer_movimientos_del_pdf(pdf_file):
    """
    Extrae movimientos de un extracto Galicia, detectando automáticamente el formato
    (Office Banking o Resumen de Cuenta Corriente) línea por línea. Cada movimiento
    puede ocupar varias líneas: la primera trae fecha + concepto + importe + saldo,
    y las siguientes agregan datos adicionales (beneficiario, CUIT, banco, categoría)
    a la columna 'detalle'. El estado "dentro de la tabla" persiste entre páginas,
    porque el encabezado de columnas no siempre se repite en cada una.
    """
    movimientos = []
    inicio_movimientos = False

    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            lineas = texto.split('\n')

            for linea_raw in lineas:
                linea = linea_raw.strip()
                if not linea:
                    continue

                if not inicio_movimientos:
                    if "Fecha" in linea and "Descripción" in linea and ("Débito" in linea or "Crédito" in linea) and "Saldo" in linea:
                        inicio_movimientos = True
                    continue

                if es_linea_ruido(linea):
                    continue

                # --- Probamos primero formato Office Banking (año de 4 dígitos) ---
                match_ob = PATRON_FECHA_OB.match(linea)
                match_rc = None if match_ob else PATRON_FECHA_RC.match(linea)

                if match_ob:
                    fecha = match_ob.group(1)
                    resto = match_ob.group(2)

                    montos = PATRON_MONTO_OB.findall(resto)
                    saldo_match = PATRON_SALDO_OB.search(resto)
                    saldo_valor = limpiar_valor_numerico(saldo_match.group(1)) if saldo_match else 0.0

                    idx_dollar = resto.find('$')
                    descripcion_raw = resto[:idx_dollar].strip() if idx_dollar != -1 else resto.strip()

                    if montos:
                        signo, valor_str = montos[0]
                        valor = limpiar_valor_numerico(valor_str)
                        importe = valor if signo == '+' else -valor
                        tipo_movimiento = "Credito" if signo == '+' else "Debito"
                    else:
                        importe = 0.0
                        tipo_movimiento = "Desconocido"

                    movimientos.append({
                        'fecha': fecha,
                        'descripcion': descripcion_raw,
                        'detalle_lineas': [],
                        'importe': importe,
                        'saldo': saldo_valor,
                        'tipo_movimiento': tipo_movimiento
                    })

                elif match_rc:
                    fecha = match_rc.group(1)
                    resto = match_rc.group(2)
                    numeros = PATRON_NUM_RC.findall(resto)

                    if numeros:
                        saldo_valor = limpiar_valor_numerico(numeros[-1])
                        pos_primer_num = resto.find(numeros[0])
                        descripcion_raw = resto[:pos_primer_num].strip()

                        if len(numeros) >= 2:
                            importe = limpiar_valor_numerico(numeros[-2])
                            tipo_movimiento = "Debito" if importe < 0 else "Credito"
                        else:
                            importe = 0.0
                            tipo_movimiento = "Desconocido"
                    else:
                        descripcion_raw = resto.strip()
                        saldo_valor = 0.0
                        importe = 0.0
                        tipo_movimiento = "Desconocido"

                    # separamos un posible código de "Origen" pegado al final (ej: "00D5")
                    partes = descripcion_raw.split()
                    if partes and es_codigo_origen(partes[-1]):
                        descripcion_raw = " ".join(partes[:-1])

                    movimientos.append({
                        'fecha': fecha,
                        'descripcion': descripcion_raw.strip(),
                        'detalle_lineas': [],
                        'importe': importe,
                        'saldo': saldo_valor,
                        'tipo_movimiento': tipo_movimiento
                    })

                else:
                    # ---- Línea de detalle del movimiento anterior (incluye continuaciones entre páginas) ----
                    if movimientos:
                        movimientos[-1]['detalle_lineas'].append(linea)

    for mov in movimientos:
        mov['detalle'] = " | ".join(mov.pop('detalle_lineas'))

    return movimientos


def get_table_download_link(df):
    """Genera un enlace para descargar el DataFrame como un archivo Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="extracto_bancario.xlsx">Descargar archivo Excel</a>'


# Widget para cargar archivo
uploaded_file = st.file_uploader("Carga tu extracto bancario en PDF", type=['pdf'])

if uploaded_file is not None:
    with st.spinner('Procesando el archivo PDF...'):
        try:
            movimientos = extraer_movimientos_del_pdf(uploaded_file)

            if movimientos:
                df = pd.DataFrame(movimientos, columns=['fecha', 'descripcion', 'detalle', 'importe', 'saldo', 'tipo_movimiento'])

                st.success(f'¡Procesamiento completado! Se encontraron {len(movimientos)} movimientos.')

                st.subheader("Vista previa de los datos extraídos:")
                st.dataframe(df)

                st.markdown(get_table_download_link(df), unsafe_allow_html=True)
            else:
                st.warning("No se encontraron movimientos en el PDF. Verifica que sea un extracto bancario del Banco Galicia.")
        except Exception as e:
            st.error(f"Error al procesar el archivo: {str(e)}")

# Información adicional
st.markdown("---")
st.markdown("""
### Información
- Esta aplicación procesa extractos del Banco Galicia en **ambos formatos** (Office Banking y Resumen de Cuenta Corriente), detectando automáticamente cuál es.
- El archivo resultante tiene las columnas: fecha, descripcion, detalle, importe, saldo y tipo_movimiento.
- Los valores de débito mantienen su signo negativo para facilitar los cálculos.
- Los datos adicionales de cada movimiento (nombre, CUIT, banco, categoría) se agrupan en la columna "detalle", separados por " | ".
""")

st.markdown("---")
st.markdown("Desarrollado con Streamlit")
