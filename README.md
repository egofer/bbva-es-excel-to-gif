# BBVA Excel to QIF Converter 🏦➡️🧾

## Descripción

Este script de Python convierte los archivos de movimientos de cuenta descargados en formato Excel (`.xls` o `.xlsx`) desde la web de **BBVA España** al formato **QIF (Quicken Interchange Format)**. El script extrae los detalles de la transacción, busca un número de referencia largo en la columna "Movimiento" para el campo Número (`N`) del QIF, y combina las columnas "Concepto" y "Movimiento" originales para formar el campo Memo (`M`).

## Motivación

BBVA España permite descargar los movimientos de cuenta en formato Excel, pero muchas aplicaciones populares de finanzas personales como [KMyMoney](https://kmymoney.org/), [GnuCash](https://www.gnucash.org/), [HomeBank](https://www.gethomebank.org) o versiones antiguas de Quicken, funcionan mejor o únicamente con archivos QIF.

Este script automatiza el proceso de conversión, extrayendo la información relevante del Excel de BBVA y formateándola en un archivo QIF listo para importar. Coloca un posible número de referencia en el campo `N` y la descripción completa en el campo `M` para facilitar la identificación y categorización posterior.

## ✨ Características principales

*   **Lee formato Excel BBVA:** Procesa archivos `.xls` y `.xlsx` descargados de BBVA.
*   **Conversión a QIF:** Genera un archivo QIF estándar (`!Type:Bank`) listo para importar.
*   **Extracción de número de referencia (`N`):**
    *   Busca automáticamente una secuencia de 16 o más dígitos en la columna `Movimiento`.
    *   Si la encuentra, coloca este número en el campo **Número (`N`)** del QIF.
*   **Construcción de memo (`M`):**
    *   Combina el contenido de la columna `Concepto` y el contenido **original** de la columna `Movimiento` (incluyendo el número de referencia si lo hubiera).
    *   Este texto combinado se coloca en el campo **Memo (`M`)** del archivo QIF.
*   **Campos QIF vacíos:** Los campos Beneficiario (`P`) y Categoría (`L`) del QIF se dejan **intencionadamente vacíos**, ya que esta información no está estructurada de forma fiable en el extracto de BBVA.
*   **Manejo de formatos españoles:** Parsea correctamente importes con coma decimal y fechas en formato `DD/MM/YYYY`.
*   **Validación de datos:**
    *   Comprueba que las columnas esenciales (`F.Valor`, `Concepto`, `Movimiento`, `Importe`) estén presentes.
    *   Valida que las fechas sean válidas y estén en un rango razonable.
    *   Valida que los importes sean numéricos, omitiendo filas con datos inválidos y manejando correctamente el formato decimal español.
*   **Codificación flexible:** Permite elegir la codificación del archivo QIF de salida (`utf-8` por defecto, recomendado para compatibilidad con acentos).
*   **Modo Verbose:** Incluye una opción `-v` para mostrar información detallada del procesamiento y depuración.
*   **Modular:** El código está estructurado en funciones para facilitar su lectura y mantenimiento.

## ⚙️ Requisitos e instalación

1.  **Python:** Necesitas Python 3.6 o superior.
2.  **Bibliotecas:** Instala las dependencias necesarias usando pip:
    ```bash
    pip install pandas openpyxl xlrd
    ```
    *   `pandas`: Para leer archivos Excel.
    *   `xlrd`: Necesario si trabajas con archivos `.xls` antiguos (aunque `openpyxl` es preferido para `.xlsx`).
    *   `openpyxl`: Necesario para leer archivos `.xlsx` modernos.

## 🚀 Uso

El script se ejecuta desde la línea de comandos:

```bash
python bbva2qif.py [opciones] <archivo_excel_entrada>
```

**Argumentos:**

*   `archivo_excel_entrada`: Ruta obligatoria a tu archivo Excel (`.xls` o `.xlsx`) descargado de BBVA.

**Opciones:**

*   `-o ARCHIVO_SALIDA`, `--output ARCHIVO_SALIDA`: Especifica la ruta y nombre del archivo QIF de salida. Por defecto, se crea un archivo con el mismo nombre que el de entrada pero con extensión `.qif`.
*   `--encoding CODIFICACION`: Especifica la codificación del archivo QIF de salida. Opciones: `utf-8` (recomendado y por defecto), `cp1252`, `iso-8859-1`.
*   `-v`, `--verbose`: Activa el modo detallado, mostrando mensajes de depuración durante el procesamiento (muy útil para verificar la extracción de datos).
*   `-h`, `--help`: Muestra la ayuda con todos los argumentos y opciones.

**Ejemplos:**

*   **Conversión básica (salida por defecto `movimientos_bbva.qif`):**
    ```bash
    python bbva2qif.py movimientos_bbva.xlsx
    ```
*   **Especificando archivo de salida:**
    ```bash
    python bbva2qif.py mis_movimientos_bbva.xls -o extracto_bbva_2025.qif
    ```
*   **Activando modo detallado:**
    ```bash
    python bbva2qif.py extracto_banco_bbva.xlsx -v
    ```

## 📄 Formato del archivo Excel de entrada (esperado)

El script está diseñado para funcionar con la estructura típica de los archivos Excel descargados desde la web de BBVA España. Espera encontrar:

1.  Algunas filas iniciales con información del informe (título, fecha de generación).
2.  **Una fila de cabecera EXACTA** con los siguientes nombres de columna (buscada en las primeras 20 filas):
    ```
    F.Valor, Fecha, Concepto, Movimiento, Importe, Divisa, Disponible, Divisa, Observaciones
    ```
3.  Las filas de datos de transacciones debajo de la cabecera.

**¡Importante!** Si BBVA cambia la estructura o los nombres exactos de estas columnas, el script podría necesitar ajustes (principalmente en las constantes `EXPECTED_HEADER` y `COL_MAP`).

## 🧾 Formato del archivo QIF de salida

El script genera un archivo QIF estándar (`!Type:Bank`). Los campos se mapean de la siguiente manera:

*   `D`: Fecha (Columna `F.Valor` del Excel, Formato `MM/DD/YYYY`)
*   `T`: Importe (Columna `Importe`, con punto decimal, formato correcto)
*   `N`: Número de Referencia (Secuencia de 16+ dígitos encontrada en `Movimiento`, si existe)
*   `P`: **(VACÍO)** - Campo Beneficiario intencionadamente en blanco.
*   `L`: **(VACÍO)** - Campo Categoría intencionadamente en blanco.
*   `M`: Memo/Nota (Combinación de las columnas `Concepto` y `Movimiento` originales del Excel)
*   `^`: Separador de transacción

## 🔧 Configuración y personalización

La lógica principal está definida dentro del script (`bbva2qif.py`):

*   **Nombres de columna esperados:** La constante `EXPECTED_HEADER` define la cabecera buscada.
*   **Mapeo de columnas:** La constante `COL_MAP` asocia nombres internos a las columnas del Excel.
*   **Patrón de número de referencia:** La expresión regular `SIXTEEN_PLUS_DIGIT_PATTERN` define qué se busca para el campo `N`.
*   **Lógica de procesamiento:** La función `process_transaction_row` contiene la lógica de extracción y construcción de los campos QIF.

Para personalizaciones (ej: intentar extraer un beneficiario del concepto, asignar categorías basadas en el memo), sería necesario modificar el código Python.

## ⚠️ Troubleshooting y problemas conocidos

*   **Error "Cabecera no encontrada" / "Faltan columnas":** Verifica que tu archivo Excel descargado de BBVA contenga *exactamente* la fila de cabecera definida en `EXPECTED_HEADER`. Comprueba que no haya espacios extra o variaciones.
*   **Caracteres raros/incorrectos (acentos):** Usa `--encoding utf-8` (opción por defecto). Si tu software financiero no soporta UTF-8, prueba con `cp1252`.
*   **Errores de lectura Excel:** Asegúrate de tener `pandas`, `openpyxl`, y `xlrd` instalados correctamente (`pip install pandas openpyxl xlrd`).
*   **Importes incorrectos:** Verifica que la función `parse_spanish_decimal` esté manejando bien el formato de tu Excel (coma vs punto decimal). La versión actual intenta ser robusta.
*   **Número (`N`) o Memo (`M`) inesperados:** Usa el modo `-v` para ver en detalle cómo se leen las columnas `Concepto` y `Movimiento`, si se detecta el número para `N`, y cómo se construye el `M` final para cada fila.

## 🔮 Posibles mejoras futuras

*   **Archivo de configuración externo:** Para definir `EXPECTED_HEADER`, `COL_MAP`, y patrones sin editar el código.
*   **Reglas de mapeo avanzadas:** Permitir al usuario definir reglas (quizás en un archivo YAML o JSON) para intentar asignar un Beneficiario (`P`) o Categoría (`L`) basado en el contenido del Memo (`M`).
*   **Interfaz Gráfica (GUI).**

## 🤝 Contribuciones

¡Las contribuciones son bienvenidas! Si encuentras un error o tienes una idea para mejorar el script, por favor abre un "Issue" o envía un "Pull Request" en el repositorio donde alojes este código.

