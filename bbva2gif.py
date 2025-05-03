#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Convierte extractos bancarios descargados como archivos Excel (.xls/.xlsx) desde
BBVA España en formato QIF (Quicken Interchange Format).

Este script parsea la estructura específica de las exportaciones Excel de BBVA,
extrayendo detalles de la transacción (fecha valor, importe, concepto, movimiento).
Realiza las siguientes acciones principales:
1. Busca la fila de cabecera esperada en el archivo Excel.
2. Lee los datos de transacciones usando pandas, forzando la lectura inicial como texto.
3. Procesa cada fila:
    - Parsea la fecha ('F.Valor').
    - Parsea el importe ('Importe'), manejando el formato decimal español (coma).
    - Busca una secuencia de 16+ dígitos en 'Movimiento' y la asigna al campo 'N' (Número) del QIF.
    - Combina 'Concepto' y 'Movimiento' originales para crear el campo 'M' (Memo) del QIF.
    - Deja los campos 'P' (Beneficiario) y 'L' (Categoría) vacíos.
4. Escribe las transacciones procesadas y ordenadas por fecha en un archivo QIF.

Incluye validación de datos (rango de fechas, importe válido) y un modo
detallado para depuración.

Requiere pandas y openpyxl (o xlrd para archivos .xls antiguos). Instalar con:
pip install pandas openpyxl xlrd
"""

# Basado en el trabajo original para ING: https://github.com/egofer/ing-es-excel-to-qif
# Adaptado para BBVA

__author__ = "https://github.com/egofer"
__license__ = "MIT"
__version__ = "0.4-bbva-docstrings"  # Incremento de versión por docstrings
__status__ = "Development"
__date__ = "Mayo, 2024"

import datetime
import re
import argparse
from decimal import Decimal, InvalidOperation
import pandas as pd
import sys
# import traceback # Descomentar si se necesita traza completa en errores inesperados

# --- Constantes ---
# Rango de fechas considerado razonable para una transacción bancaria.
REASONABLE_START_DATE = datetime.datetime(1990, 1, 1)
REASONABLE_END_DATE = datetime.datetime.now() + datetime.timedelta(days=5*365)

# Cabecera exacta esperada en el archivo Excel de BBVA.
# Es crucial que coincida con la exportación del banco.
EXPECTED_HEADER = ['F.Valor', 'Fecha', 'Concepto', 'Movimiento',
                   'Importe', 'Divisa', 'Disponible', 'Divisa', 'Observaciones']

# Mapeo de nombres internos usados en el script a los nombres de columna en el Excel de BBVA.
COL_MAP = {
    'date': 'F.Valor',          # Fecha principal para la transacción QIF.
    'concept': 'Concepto',      # Descripción principal.
    # Descripción secundaria, donde se busca el nº ref.
    'movement': 'Movimiento',
    'amount': 'Importe',        # Cantidad de la transacción.
    # 'op_date': 'Fecha',       # Fecha de operación (actualmente no usada en QIF).
    # 'notes': 'Observaciones'  # Campo de observaciones (actualmente no usada en QIF).
}

# Lista de claves internas que representan columnas absolutamente necesarias para procesar una fila.
REQUIRED_COLS_INTERNAL = ['date', 'concept', 'movement', 'amount']

# --- Compilación de Regex ---
# Expresión regular para buscar cualquier secuencia de 16 o más dígitos consecutivos.
# Se usará para extraer el número de referencia del campo 'Movimiento'.
# El grupo 1 captura los dígitos encontrados.
SIXTEEN_PLUS_DIGIT_PATTERN = re.compile(r'(\d{16,})')

# --- Funciones ---


def parse_arguments():
    """
    Parsea los argumentos proporcionados en la línea de comandos.

    Utiliza el módulo argparse para definir y leer los argumentos esperados:
    el archivo Excel de entrada y opciones para el archivo de salida,
    la codificación y el modo detallado (verbose).

    Returns:
        argparse.Namespace: Un objeto que contiene los argumentos parseados
                            como atributos (ej. args.excel_file, args.verbose).
    """
    parser = argparse.ArgumentParser(
        description="Convierte extracto Excel BBVA a QIF",
        epilog="Ejemplo: python bbva2qif.py extracto_bbva.xlsx -o salida.qif -v"
    )
    parser.add_argument("excel_file", help="Ruta al archivo Excel de BBVA.")
    parser.add_argument(
        "-o", "--output", help="Ruta QIF salida (defecto: nombre_excel.qif).")
    parser.add_argument("--encoding", default="utf-8", choices=[
                        "utf-8", "cp1252", "iso-8859-1"], help="Codificación salida QIF (defecto: utf-8).")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Activar mensajes detallados para depuración.")
    return parser.parse_args()


def parse_spanish_decimal(decimal_val, row_num, verbose=False):
    """
    Convierte un valor (generalmente string) a un objeto Decimal.

    Maneja específicamente el formato numérico español donde la coma (,) es el
    separador decimal y el punto (.) puede ser el separador de miles. También
    intenta manejar el formato estándar (punto como decimal) si no hay coma.
    Limpia espacios en blanco y el símbolo del Euro (€) antes de la conversión.

    Args:
        decimal_val (str | float | int | None): El valor a convertir. Se espera
            que sea un string leído del Excel, pero se intenta manejar otros tipos.
        row_num (int): El número original de la fila en el archivo Excel,
                       utilizado únicamente para mensajes de error/depuración.
        verbose (bool, optional): Si es True, imprime mensajes de depuración
                                  detallando el proceso de conversión.
                                  Por defecto es False.

    Returns:
        Decimal | None: El objeto Decimal representando el importe si la
                        conversión es exitosa y el valor es finito.
                        Retorna None si el valor de entrada es NaN, vacío,
                        no se puede convertir, o resulta en un valor no finito
                        (Inf, -Inf). Se imprime un aviso o error en caso de fallo.
    """
    if pd.isna(decimal_val):
        if verbose:
            print(f"  [DEBUG] Fila {row_num}: Importe NaN/Vacío.")
        return None

    # Asegurar que trabajamos con un string y quitar espacios laterales
    decimal_str = str(decimal_val).strip()
    if not decimal_str:
        if verbose:
            print(f"  [DEBUG] Fila {row_num}: Importe vacío tras strip.")
        return None

    # Paso 1: Quitar caracteres no numéricos básicos (espacios internos, €)
    chars_to_remove = " €"
    translation_table = str.maketrans('', '', chars_to_remove)
    intermediate_str = decimal_str.translate(translation_table)

    final_str_for_decimal = ""

    # Paso 2: Decidir estrategia basada en la presencia de coma
    if ',' in intermediate_str:
        # ASUMIMOS FORMATO ESPAÑOL: ',' es decimal, '.' es miles
        if verbose:
            print(f"  [DEBUG] Detectada coma en '{
                  intermediate_str}', aplicando formato español.")
        # Quitar separador de miles (punto)
        temp_str = intermediate_str.replace('.', '')
        # Reemplazar coma decimal por punto
        final_str_for_decimal = temp_str.replace(',', '.')
    else:
        # NO HAY COMA: Asumimos formato estándar (punto es decimal) o entero
        if verbose:
            print(f"  [DEBUG] No detectada coma en '{
                  intermediate_str}', usando como formato estándar/entero.")
        # No quitamos puntos aquí, ya que podrían ser el decimal correcto
        final_str_for_decimal = intermediate_str

    if verbose:
        print(f"  [DEBUG] String final para conversión Decimal: '{
              final_str_for_decimal}'")

    # Paso 3: Intentar conversión a Decimal
    try:
        # Validar que la cadena no esté vacía después de la limpieza
        if not final_str_for_decimal:
            if verbose:
                print(f"  [DEBUG] Fila {
                      row_num}: String vacío tras limpieza. Omitiendo.")
            return None

        dec_val = Decimal(final_str_for_decimal)

        # Validar que sea finito (no Infinito ni NaN)
        if not dec_val.is_finite():
            if verbose:
                print(f"  [DEBUG] Fila {row_num}: Importe no finito. Original: '{
                      decimal_str}' -> Final: '{final_str_for_decimal}'. Omitiendo.")
            return None

        # Éxito
        if verbose:
            print(f"  [DEBUG] Importe parseado: Original: '{
                  decimal_str}' -> Decimal: {dec_val}")
        return dec_val

    except InvalidOperation:
        # Manejar el caso específico de 'nan' si sobrevive a la limpieza
        if final_str_for_decimal.lower() == 'nan':
            if verbose:
                print(f"  [DEBUG] Fila {row_num}: Importe explícitamente NaN.")
            return None
        # Error genérico de conversión Decimal
        print(f"  AVISO: Fila {row_num}: No se pudo convertir importe a Decimal. Original: '{
              decimal_str}', Procesado: '{final_str_for_decimal}'. Omitiendo.")
        return None
    except Exception as e:  # Capturar cualquier otro error inesperado durante la conversión
        print(f"  ERROR INESPERADO parseando decimal en fila {row_num}: Original: '{
              decimal_str}', Procesado: '{final_str_for_decimal}'. Error: {e}. Omitiendo.")
        # traceback.print_exc() # Descomentar para traza completa
        return None


def find_header_row(excel_filepath, expected_header, verbose=False):
    """
    Busca la fila de cabecera especificada en las primeras filas de un archivo Excel.

    Lee las primeras 20 filas del archivo Excel como texto plano para evitar
    conversiones de tipos automáticas por pandas. Luego itera sobre estas filas
    buscando una secuencia de celdas que coincida exactamente con la lista
    `expected_header`, permitiendo que la cabecera no empiece necesariamente
    en la primera columna. Limpia posibles sufijos ".0" que pandas añade a
    números leídos como texto.

    Args:
        excel_filepath (str): La ruta al archivo Excel (.xls o .xlsx).
        expected_header (list[str]): Una lista de strings que representa la
                                     secuencia exacta de nombres de columna esperada.
        verbose (bool, optional): Si es True, imprime mensajes de depuración,
                                  incluyendo la fila encontrada y las filas leídas
                                  si no se encuentra la cabecera. Por defecto es False.

    Returns:
        int: El índice (basado en 0) de la fila donde se encontró la cabecera.
             Retorna -1 si la cabecera no se encuentra en las primeras 20 filas.

    Raises:
        FileNotFoundError: Si `excel_filepath` no existe.
        Exception: Puede propagar otras excepciones relacionadas con la lectura
                   del archivo Excel (ej., archivo corrupto, permisos).
    """
    if verbose:
        print(f"Buscando cabecera: {expected_header}")
    header_row_index = -1
    try:
        # Leer como texto para evitar conversiones numéricas automáticas
        df_pre = pd.read_excel(excel_filepath, header=None,
                               keep_default_na=False, nrows=20, dtype=str)
    except FileNotFoundError:
        print(f"Error Fatal: Archivo no encontrado '{excel_filepath}'")
        return -1  # Devolver -1 directamente si no se encuentra el archivo
    except Exception as e:
        print(f"Error Fatal leyendo inicio Excel '{excel_filepath}': {e}")
        # traceback.print_exc() # Descomentar para traza completa
        return -1

    header_found_flag = False
    # Iterar sobre las filas leídas (índice y valores)
    for idx, row_values in enumerate(df_pre.values.tolist()):
        # Convertir todos los valores de la fila a string y limpiar espacios
        row_str = [str(v).strip() for v in row_values]
        # Buscar la secuencia `expected_header` dentro de la fila `row_str`
        for i in range(len(row_str) - len(expected_header) + 1):
            # Obtener el segmento de la fila actual del tamaño de la cabecera esperada
            segment = row_str[i:i+len(expected_header)]
            # Limpiar posible ".0" añadido por pandas a números leídos como texto
            segment_cleaned = [s.replace('.0', '') if isinstance(
                s, str) and s.endswith('.0') else s for s in segment]
            # Comparar el segmento limpio con la cabecera esperada
            if segment_cleaned == expected_header:
                header_found_flag = True
                header_row_index = idx  # Guardar el índice de la fila
                if verbose:
                    print(f"Cabecera detectada en índice {header_row_index} (Fila Excel {
                          header_row_index + 1}), empezando en columna {i+1}.")
                break  # Salir del bucle de búsqueda de segmento
        if header_found_flag:
            break  # Salir del bucle de filas

    if not header_found_flag:
        print(f"Error Fatal: Cabecera {
              expected_header} no encontrada en las primeras 20 filas de '{excel_filepath}'.")
        # Imprimir filas leídas para depuración si está en modo verbose
        if verbose:
            print("--- Filas leídas (primeras 20) ---")
            for i, row in enumerate(df_pre.values.tolist()):
                print(f"Índice {i}: {row}")
            print("--- Fin filas leídas ---")
        return -1  # Indicar que no se encontró

    return header_row_index


def read_excel_data(excel_filepath, header_row_index, verbose=False):
    """
    Lee los datos de transacciones del archivo Excel a partir de la fila de cabecera.

    Utiliza pandas.read_excel especificando la fila de cabecera encontrada.
    Importante: Lee todas las columnas como tipo 'string' (`dtype=str`) para
    tener control total sobre la conversión de tipos posterior (especialmente
    para fechas e importes) y evitar interpretaciones incorrectas de pandas.
    Limpia los nombres de las columnas leídas (quitando espacios). Valida si
    las columnas críticas definidas en `COL_MAP` están presentes.

    Args:
        excel_filepath (str): La ruta al archivo Excel.
        header_row_index (int): El índice (basado en 0) de la fila que
                                contiene los encabezados de columna.
        verbose (bool, optional): Si es True, imprime las columnas leídas y
                                  potencialmente las primeras filas de datos
                                  (comentado por defecto). Por defecto es False.

    Returns:
        pd.DataFrame | None: Un DataFrame de pandas que contiene los datos de las
                            transacciones leídos como strings.
                            Retorna None si ocurre un error durante la lectura
                            o si faltan columnas mapeadas críticas.

    Raises:
        FileNotFoundError: Si `excel_filepath` no existe (aunque find_header_row ya lo chequea).
        Exception: Puede propagar otras excepciones de pandas o de lectura de Excel.
    """
    if verbose:
        print(f"Leyendo datos con cabecera en índice {header_row_index}...")
    try:
        # Forzar lectura de todas las columnas como texto inicialmente
        df_data = pd.read_excel(
            excel_filepath,
            header=header_row_index,
            keep_default_na=False,  # No interpretar 'NA', 'NULL' etc. como NaN automáticamente
            dtype=str              # Leer todo como string
        )
        # Limpiar nombres de columnas leídos (quitar espacios laterales)
        df_data.columns = df_data.columns.map(
            lambda x: str(x).strip() if pd.notna(x) else x)

        if verbose:
            print(f"Columnas leídas (como texto inicialmente): {
                  df_data.columns.tolist()}")
            # Descomentar para ver las primeras filas tal cual se leyeron:
            # try:
            #     print("Primeras 5 filas leídas (como texto):")
            #     print(df_data.head().to_string())
            # except Exception as e_print:
            #     print(f"  No se pudieron imprimir las primeras filas: {e_print}")

        # Validar que las columnas que esperamos mapear realmente existen
        missing_mapped_cols = [col_name for col_key, col_name in COL_MAP.items(
        ) if col_name not in df_data.columns]
        if missing_mapped_cols:
            print(f"ADVERTENCIA: Columnas mapeadas esperadas no encontradas en el archivo: {
                  missing_mapped_cols}")
            # Comprobar si faltan columnas críticas definidas como requeridas
            critical_missing = [mc for mc in missing_mapped_cols if mc in [
                COL_MAP[k] for k in REQUIRED_COLS_INTERNAL]]
            if critical_missing:
                print(f"Error Fatal: Faltan columnas críticas mapeadas necesarias para procesar: {
                      critical_missing}")
                return None  # Fallar si faltan columnas esenciales
        return df_data

    except Exception as e:
        print(f"Error Fatal leyendo datos principales del Excel '{
              excel_filepath}': {e}")
        # traceback.print_exc() # Descomentar para traza completa
        return None


def process_transaction_row(row, original_excel_row, col_map, verbose=False):
    """
    Procesa una única fila de datos (transacción) del DataFrame leído del Excel.

    Extrae los datos relevantes (fecha, importe, concepto, movimiento) de la fila
    (que se espera contenga strings). Realiza las siguientes operaciones:
    - Parsea la fecha valor.
    - Parsea el importe usando `parse_spanish_decimal`.
    - Busca un número de referencia de 16+ dígitos en el campo 'Movimiento' original.
    - Construye el texto para el campo Memo ('M') del QIF combinando 'Concepto' y
      'Movimiento' originales.
    - Prepara un diccionario con los campos formateados para la generación del QIF
      ('date', 'amount', 'memo', 'transaction_ref_num'), dejando 'payee' y
      'category' como None.

    Args:
        row (pd.Series): La fila del DataFrame a procesar. Se asume que los
                         valores son strings debido a la lectura con `dtype=str`.
        original_excel_row (int): El número de fila original en el archivo Excel
                                  (1-based), usado para mensajes de error claros.
        col_map (dict[str, str]): El diccionario que mapea nombres internos
                                  (ej: 'date') a los nombres de columna del Excel
                                  (ej: 'F.Valor').
        verbose (bool, optional): Si es True, imprime información detallada
                                  sobre el procesamiento de la fila, incluyendo
                                  datos extraídos y resultados finales.
                                  Por defecto es False.

    Returns:
        dict | None: Un diccionario listo para ser usado en la generación del QIF,
                     conteniendo claves como 'date' (datetime), 'amount' (Decimal),
                     'memo' (str|None), 'transaction_ref_num' (str|None), 'payee' (None),
                     'category' (None).
                     Retorna None si la fila no contiene datos válidos (ej: fecha
                     o importe incorrectos) y debe ser omitida.
    """
    if verbose:
        print(f"\n--- Procesando Fila Excel {original_excel_row} ---")
        # Descomentar para ver los datos crudos de la fila leída como string
        # try:
        #     print(f"  [RAW DATA] {row.to_dict()}")
        # except Exception as e_print:
        #      print(f"  [RAW DATA] Error al mostrar datos: {e_print}")

    try:
        # --- Fecha ---
        # Acceder como string
        raw_date_val = row.get(col_map['date'], '').strip()
        if not raw_date_val:
            if verbose:
                print(f"  [DEBUG] Fila {
                      original_excel_row}: Fecha Valor vacía. Omitiendo.")
            return None
        tx_date = None
        try:
            # Intentar parsear formato DD/MM/YYYY
            tx_date = datetime.datetime.strptime(raw_date_val, "%d/%m/%Y")
        except ValueError:
            # Si falla, intentar interpretar como número de serie de Excel
            try:
                # Excel guarda fechas como números
                date_num = float(raw_date_val)
                # Época de Excel en Windows/moderno: 30/12/1899 día 0.
                excel_epoch = datetime.datetime(1899, 12, 30)
                delta = datetime.timedelta(days=date_num)
                tx_date = excel_epoch + delta
                if verbose:
                    print(f"  [DEBUG] Fecha parseada desde número Excel '{
                          raw_date_val}' -> {tx_date.strftime('%d/%m/%Y')}")
            except (ValueError, TypeError, OverflowError) as e_num:
                print(f"  OMITIENDO fila {original_excel_row}: Fecha inválida o formato no reconocido '{
                      raw_date_val}'. Error: {e_num}")
                return None

        # Validar rango razonable de fecha
        if not (REASONABLE_START_DATE <= tx_date <= REASONABLE_END_DATE):
            # Solo avisar, no omitir por esto necesariamente
            print(f"  AVISO: Fila {original_excel_row}: Fecha '{
                  tx_date.strftime('%d/%m/%Y')}' fuera del rango razonable.")

        # --- Importe ---
        raw_amount_val = row.get(col_map['amount'], '')  # Acceder como string
        amount = parse_spanish_decimal(
            raw_amount_val, original_excel_row, verbose)
        if amount is None:
            # parse_spanish_decimal ya imprimió el aviso/error si es necesario
            return None  # Omitir fila si el importe no es válido

        # --- Concepto y Movimiento (Leer originales como string) ---
        concept = row.get(col_map['concept'], '').strip()
        movement_original = row.get(col_map['movement'], '').strip()

        if verbose:
            print(f"  [DATA] Concepto: '{concept}'")
            print(f"  [DATA] Movimiento (Original): '{movement_original}'")

        # --- Búsqueda del Número de Referencia (16+ dígitos) para el campo N ---
        transaction_ref_num = None
        if movement_original:  # Solo buscar si hay texto en la columna Movimiento
            match = SIXTEEN_PLUS_DIGIT_PATTERN.search(movement_original)
            if match:
                # Capturar el número encontrado
                transaction_ref_num = match.group(1)
                if verbose:
                    print(f"  [REGEX] Número Referencia (N) ENCONTRADO: '{
                          transaction_ref_num}'")
            elif verbose:
                print(
                    f"  [REGEX] No se encontró número de 16+ dígitos en Movimiento.")
        elif verbose:
            print(f"  [DEBUG] Columna Movimiento vacía, no se busca número.")

        # --- Construcción del Memo (usando concepto y movimiento ORIGINAL) ---
        memo_parts = [concept]
        # Añadir movimiento original si tiene contenido y es diferente del concepto (ignorando mayúsculas/minúsculas)
        if movement_original and movement_original.lower() != concept.lower():
            memo_parts.append(movement_original)

        # Filtrar partes vacías (por si 'concept' o 'movement' fueran vacíos) y unir con " - "
        memo_parts_filtered = [part for part in memo_parts if part]
        memo_text = " - ".join(memo_parts_filtered) if memo_parts_filtered else None
        # Limpiar espacios dobles o laterales en el memo final
        if memo_text:
            memo_text = re.sub(r'\s{2,}', ' ', memo_text).strip()
            # Si después de limpiar queda vacío, ponerlo a None
            if not memo_text:
                memo_text = None

        # Imprimir resultados finales si estamos en modo verbose
        if verbose:
            print(f"  [FINAL] Fecha (D): {tx_date.strftime('%d/%m/%Y')}")
            print(f"  [FINAL] Importe (T): {amount}")
            print(f"  [FINAL] Número Ref (N): '{
                  transaction_ref_num}'")  # Puede ser None
            # Puede ser None
            print(f"  [FINAL] Memo (M): '{memo_text}'")
            print(f"  [FINAL] Payee (P): None")
            print(f"  [FINAL] Categoría (L): None")

        # --- Devolver datos estructurados para QIF ---
        transaction_data = {
            'date': tx_date,                # datetime object
            'amount': amount,               # Decimal object
            'payee': None,                  # Siempre None para este script
            'category': None,               # Siempre None para este script
            'memo': memo_text,              # str | None
            'transaction_ref_num': transaction_ref_num  # str | None
        }
        return transaction_data

    except Exception as e:
        # Captura cualquier error no esperado durante el procesamiento de la fila
        print(
            f"*** Error INESPERADO procesando fila Excel {original_excel_row}: {e} ***")
        # Intentar mostrar los datos crudos de la fila que causó el error
        try:
            print(f"    Datos fila (raw): {row.to_dict()}")
        except Exception as inner_e:
            print(
                f"    No se pudieron mostrar los datos de la fila: {inner_e}")
        # traceback.print_exc() # Descomentar para traza completa del error
        return None  # Omitir esta fila


def generate_qif_file(transactions, qif_filepath, output_encoding, verbose=False):
    """
    Genera y escribe el archivo QIF a partir de una lista de transacciones procesadas.

    Escribe la cabecera '!Type:Bank'. Luego, itera sobre la lista de
    diccionarios de transacciones, escribiendo los campos QIF correspondientes
    (D, T, N, M) para cada una, separados por el carácter '^'.
    Maneja la codificación de salida especificada y los posibles errores de
    escritura. Los campos P (Payee) y L (Category) no se escriben.

    Args:
        transactions (list[dict]): Una lista de diccionarios, donde cada
                                   diccionario representa una transacción
                                   válida devuelta por `process_transaction_row`.
                                   Se espera que la lista esté ordenada por fecha.
        qif_filepath (str): La ruta completa donde se guardará el archivo QIF.
        output_encoding (str): La codificación de texto a usar para el archivo
                               de salida (ej: 'utf-8', 'cp1252').
        verbose (bool, optional): Si es True, imprime mensajes al inicio y
                                  fin de la generación del archivo.
                                  Por defecto es False.

    Returns:
        bool: True si el archivo QIF se escribió correctamente, False si ocurrió
              algún error durante la escritura.
    """
    if verbose:
        print(f"\nGenerando QIF: {
              qif_filepath} (Codificación: {output_encoding})")
    # Usar 'replace' para caracteres no soportados si la codificación no es UTF-8
    write_errors_mode = 'replace' if output_encoding.lower() != 'utf-8' else 'strict'

    try:
        with open(qif_filepath, mode='w', encoding=output_encoding, errors=write_errors_mode) as outfile:
            # Escribir cabecera QIF estándar para cuentas bancarias
            outfile.write("!Type:Bank\n")
            # Iterar sobre cada diccionario de transacción
            for tx in transactions:
                # D - Fecha (formato MM/DD/YYYY requerido por QIF)
                outfile.write(f"D{tx['date'].strftime('%m/%d/%Y')}\n")
                # T - Importe (formato con punto decimal, dos decimales)
                outfile.write(f"T{tx['amount']:.2f}\n")

                # N - Número de Referencia (si existe en el diccionario)
                if tx.get('transaction_ref_num'):
                    # Limpiar posibles saltos de línea en el número (poco probable, pero seguro)
                    ref_num_cleaned = tx['transaction_ref_num'].replace(
                        '\n', '').replace('\r', '')
                    outfile.write(f"N{ref_num_cleaned}\n")

                # P - Payee (Beneficiario) - No se escribe (siempre None en tx)
                # L - Category (Categoría) - No se escribe (siempre None en tx)

                # M - Memo (Nota) (si existe en el diccionario)
                if tx.get('memo'):
                    # Limpiar saltos de línea dentro del memo
                    memo_cleaned = tx['memo'].replace(
                        '\n', ' ').replace('\r', '')
                    outfile.write(f"M{memo_cleaned}\n")

                # ^ - Separador de fin de transacción
                outfile.write("^\n")

        print(f"Archivo QIF creado con éxito: {qif_filepath}")
        return True
    except IOError as e:
        # Error específico de entrada/salida (ej: permisos, disco lleno)
        print(f"Error de E/S escribiendo QIF en '{qif_filepath}': {e}")
        return False
    except Exception as e:
        # Cualquier otro error durante la escritura del archivo
        print(f"Error Fatal escribiendo archivo QIF '{qif_filepath}': {e}")
        # traceback.print_exc() # Descomentar para traza completa
        return False

# --- Función Principal ---


def main():
    """
    Función principal que orquesta todo el proceso de conversión.

    1.  Imprime mensajes iniciales.
    2.  Parsea los argumentos de la línea de comandos.
    3.  Determina el nombre del archivo de salida.
    4.  Llama a `find_header_row` para localizar la cabecera en el Excel.
    5.  Llama a `read_excel_data` para leer los datos de transacciones.
    6.  Valida que las columnas necesarias estén presentes en el DataFrame leído.
    7.  Itera sobre las filas del DataFrame, llamando a `process_transaction_row`
        para procesar cada una.
    8.  Recopila las transacciones procesadas válidas y cuenta las omitidas.
    9.  Imprime un resumen del procesamiento.
    10. Ordena las transacciones válidas por fecha.
    11. Llama a `generate_qif_file` para escribir el archivo QIF final.
    12. Imprime un mensaje final de éxito o error y finaliza con el
        código de salida apropiado (0 para éxito, 1 para error).
    """
    print("--- Conversor Excel (BBVA) a QIF ---")
    print("--- Requiere 'pandas', 'openpyxl', 'xlrd' ---")

    args = parse_arguments()

    # Determinar nombre del archivo de salida
    if args.output:
        output_filename = args.output
    else:
        # Si no se especifica -o, usar el nombre del archivo de entrada con extensión .qif
        base_name = args.excel_file
        output_filename = re.sub(
            r'\.[Xx][Ll][Ss][Xx]?$', '', base_name, flags=re.IGNORECASE) + ".qif"
        print(f"Nombre de archivo de salida no especificado. Usando: '{
              output_filename}'")

    # --- Pasos principales de la conversión ---

    # 1. Encontrar la fila de cabecera
    print(f"\n1. Buscando cabecera en '{args.excel_file}'...")
    header_idx = find_header_row(
        args.excel_file, EXPECTED_HEADER, args.verbose)
    if header_idx == -1:
        # find_header_row ya imprime el mensaje de error detallado
        sys.exit(1)  # Terminar si no se encuentra la cabecera

    # 2. Leer los datos del Excel (forzando tipo string inicialmente)
    print(f"\n2. Leyendo datos del archivo Excel...")
    df_data = read_excel_data(args.excel_file, header_idx, args.verbose)
    if df_data is None:
        # read_excel_data ya imprime el mensaje de error
        sys.exit(1)  # Terminar si la lectura falla o faltan columnas críticas

    # 3. Validar columnas requeridas (Doble chequeo post-lectura)
    # Aunque read_excel_data ya hace una validación, confirmamos aquí.
    current_columns = df_data.columns.tolist()
    missing_req_cols_final = [col_name for col_key, col_name in COL_MAP.items()
                              if col_key in REQUIRED_COLS_INTERNAL and col_name not in current_columns]
    if missing_req_cols_final:
        print(f"Error Fatal: Faltan columnas requeridas en el Excel (después de leer con cabecera en fila {
              header_idx+1}): {missing_req_cols_final}")
        print(f"Columnas encontradas: {current_columns}")
        print(f"Asegúrate de que el archivo '{
              args.excel_file}' tiene las columnas: {EXPECTED_HEADER}")
        sys.exit(1)

    # 4. Procesar cada fila de transacción
    print(f"\n3. Procesando filas de transacciones...")
    processed_transactions = []
    skipped_count = 0
    # Iterar sobre las filas del DataFrame usando iterrows()
    for idx, row in df_data.iterrows():
        # Calcular número de fila original en Excel (1-based)
        # Fila = índice_cabecera(0-based) + 1 (para ajustar a 1-based) + 1 (fila cabecera) + índice_fila_datos(0-based)
        original_excel_row = header_idx + 2 + idx
        # Procesar la fila actual
        processed_data = process_transaction_row(
            row, original_excel_row, COL_MAP, args.verbose)
        # Si el procesamiento fue exitoso (retornó un diccionario)
        if processed_data:
            processed_transactions.append(processed_data)
        else:
            # Si retornó None, la fila fue omitida (ya se imprimió aviso/error)
            skipped_count += 1

    # 5. Mostrar resumen del procesamiento
    print("-" * 30)
    print(f"Procesamiento de filas completado.")
    processed_count = len(processed_transactions)
    print(f"  Transacciones procesadas válidas: {processed_count}")
    print(f"  Filas omitidas/inválidas: {skipped_count}")

    # Verificar si se procesó alguna transacción
    if not processed_transactions:
        print("\nError Fatal: No se procesaron transacciones válidas. Revisa el archivo Excel o los mensajes de error/aviso anteriores.")
        sys.exit(1)  # Terminar si no hay nada que escribir
    print("-" * 30)

    # Ordenar las transacciones procesadas por fecha (ascendente)
    # Es una buena práctica antes de generar el QIF
    processed_transactions.sort(key=lambda tx: tx['date'])
    if args.verbose:
        print("Transacciones ordenadas por fecha.")

    # 6. Generar el archivo QIF final
    print(f"\n4. Generando archivo QIF...")
    success = generate_qif_file(
        processed_transactions, output_filename, args.encoding, args.verbose)

    # Mensaje final basado en el éxito de la escritura del QIF
    print(
        f"\n--- Ejecución Finalizada {'con Éxito' if success else 'con Errores'} ---")
    if not success:
        sys.exit(1)  # Terminar con código de error si la generación falló


# --- Punto de Entrada ---
if __name__ == "__main__":
    # Esta construcción asegura que main() solo se ejecute cuando el script
    # es llamado directamente (no cuando es importado como módulo).
    main()
