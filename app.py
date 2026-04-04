# app.py
# -*- coding: utf-8 -*-

# ==============================================================================
# PARTE 1: IMPORTACIONES, CONFIGURACIÓN GENERAL DE LA APLICACIÓN Y CONSTANTES GLOBALES
# ==============================================================================
# Esta sección incluye todas las librerías necesarias para el funcionamiento
# de la aplicación, la configuración inicial de la página en Streamlit, y
# la definición de todas las constantes globales que rigen el comportamiento
# del sistema: nombres de archivos, hojas esperadas, tokens a ignorar,
# opciones no productivas, patrones de "no se observa", preguntas que se
# mantienen desglosadas (exactas), etiquetas unificadas por pregunta, y
# refuerzos directos (aliases extra) para preguntas unificadas.
# ==============================================================================

import io
import re
import csv
import unicodedata
from pathlib import Path
from collections import defaultdict, Counter

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Configuración de la página en Streamlit
# --------------------------------------------------------------------------
# Se establece el título que aparece en la pestaña del navegador y se
# habilita el layout ancho para aprovechar mejor el espacio horizontal.
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Conteo de respuestas por pregunta / descriptor",
    layout="wide"
)

# =========================================================
# CONFIGURACIÓN
# =========================================================

# --------------------------------------------------------------------------
# EXCEL_GUIA: Nombre del archivo Excel que sirve como guía para identificar
# las preguntas y sus descriptores correspondientes. Debe estar ubicado en
# la raíz del repositorio.
# --------------------------------------------------------------------------
EXCEL_GUIA = "Guía de Preguntas para paretos 2026.xlsx"

# --------------------------------------------------------------------------
# SHEET_BY_FILETYPE: Diccionario que mapea cada tipo de archivo (según su
# nombre) a la hoja correspondiente dentro del Excel guía. Se incluyen
# variantes ortográficas ("policia" y "policial") para maximizar la
# detección automática del tipo de archivo.
# Nota: "Comunidad " incluye un espacio final porque así está nombrada
# la hoja en el Excel original.
# --------------------------------------------------------------------------
SHEET_BY_FILETYPE = {
    "comunidad": "Comunidad ",
    "comercio": "Comercio",
    "policia": "Policia",
    "policial": "Policia",
}

# --------------------------------------------------------------------------
# TOKENS_IGNORAR: Conjunto de tokens que representan valores vacíos o nulos
# después de la normalización. Cualquier celda cuyo contenido normalizado
# coincida con alguno de estos se considera vacía y no se procesa.
# --------------------------------------------------------------------------
TOKENS_IGNORAR = {"", "nan", "none", "null"}

# --------------------------------------------------------------------------
# OPCIONES_NO_PRODUCTIVAS: Conjunto de tokens que, aunque no están vacíos,
# no representan respuestas útiles para el conteo (ej.: "otro", "especifique",
# "cuales"). Se excluyen del conteo en todas las preguntas, con excepciones
# específicas manejadas en la función is_unproductive_option.
# --------------------------------------------------------------------------
OPCIONES_NO_PRODUCTIVAS = {
    "otro",
    "otros",
    "otra",
    "otras",
    "otro_problema",
    "otro_problema_que_considere_importante",
    "otros_delitos",
    "otros_delitos_cuales",
    "cual",
    "cuales",
    "especifique",
}

# --------------------------------------------------------------------------
# PATRONES_NO_OBSERVA: Lista de patrones (subcadenas) que identifican
# opciones del tipo "no se observa", "no hay", "ninguno", etc. Cualquier
# token que contenga alguno de estos patrones se considera no productivo,
# salvo la excepción explícita de "no_se_observan_delitos_ambientales"
# en la pregunta 27 de comunidad.
# --------------------------------------------------------------------------
PATRONES_NO_OBSERVA = [
    "no_se_observa",
    "no_se_observan",
    "no_se_presenta",
    "no_se_presentan",
    "no_hay",
    "ninguno",
    "ninguna",
]

# --------------------------------------------------------------------------
# PREGUNTAS_EXACTAS: Diccionario que indica, por tipo de archivo, cuáles
# preguntas deben mantenerse desglosadas descriptor por descriptor (modo
# "exacto"). El valor None significa que TODAS las preguntas de ese tipo
# son exactas. Un conjunto con números específicos significa que solo esas
# preguntas son exactas; las demás se unifican bajo una sola etiqueta.
# --------------------------------------------------------------------------
PREGUNTAS_EXACTAS = {
    "comunidad": {"12", "18", "20", "22", "24", "26", "27"},
    "comercio": {"12", "18", "20", "22"},
    "policia": None,
    "policial": None,
}

# --------------------------------------------------------------------------
# UNIFIED_LABELS: Diccionario que define la etiqueta unificada (legible)
# para las preguntas que NO son exactas. La clave es una tupla
# (tipo_archivo, numero_pregunta) y el valor es el texto que se mostrará
# como descriptor único para esa pregunta.
# --------------------------------------------------------------------------
UNIFIED_LABELS = {
    ("comunidad", "13"): "Oferta de servicios y oportunidades",
    ("comunidad", "15"): "Infraestructura vial",
    ("comunidad", "16"): "Espacios de riesgo",
    ("comunidad", "19"): "Venta de drogas",
    ("comunidad", "21"): "Delitos sexuales",
    ("comunidad", "23"): "Estafa",
    ("comunidad", "25"): "Abandono de personas",
    ("comunidad", "28"): "Trata de personas",

    ("comercio", "13"): "Oferta de servicios y oportunidades",
    ("comercio", "14"): "Infraestructura vial",
    ("comercio", "15"): "Inversión social",
    ("comercio", "19"): "Venta de drogas",
    ("comercio", "21"): "Estafa",
}

# --------------------------------------------------------------------------
# UNIFIED_EXTRA_ALIASES: Diccionario de refuerzos directos para preguntas
# unificadas. Contiene tokens adicionales que deben reconocerse como parte
# del grupo unificado, incluso si no aparecen explícitamente como
# descriptores en el Excel guía. Se usa durante la construcción dinámica
# de aliases para ampliar la cobertura de coincidencias.
# --------------------------------------------------------------------------
UNIFIED_EXTRA_ALIASES = {
    ("comercio", "13"): {
        "falta_de_oferta_educativa",
        "falta_de_oferta_laboral",
        "falta_de_oferta_recreativa",
        "falta_de_actividades_culturales",
    },
    ("comunidad", "13"): {
        "falta_de_oferta_educativa",
        "falta_de_oferta_laboral",
        "falta_de_oferta_recreativa",
        "falta_de_actividades_culturales",
    },
    ("comercio", "14"): {
        "calles_en_mal_estado",
        "falta_de_iluminacion",
        "falta_de_senalizacion",
        "falta_o_deterioro_de_aceras",
    },
    ("comunidad", "15"): {
        "calles_en_mal_estado",
        "falta_de_iluminacion",
        "falta_de_senalizacion",
        "falta_o_deterioro_de_aceras",
    },
    ("comercio", "15"): {
        "falta_de_programas_sociales",
        "falta_de_espacios_de_integracion_social",
        "falta_de_apoyo_institucional",
        "falta_de_inversion_social",
    },
    ("comunidad", "16"): {
        "calles_solas_u_oscuras",
        "parques_o_lotes_abandonados",
        "edificaciones_abandonadas",
        "paradas_de_bus_inseguras",
        "puentes_peatonales_inseguros",
        "zonas_sin_vigilancia",
    },
    ("comercio", "19"): {
        "en_via_publica",
        "en_espacios_cerrados_casas_edificaciones_u_otros_inmuebles",
        "de_forma_ocasional_o_movil_modalidad_expres_sin_punto_fijo",
        "venta_de_drogas",
    },
    ("comunidad", "19"): {
        "en_via_publica",
        "en_espacios_cerrados_casas_edificaciones_u_otros_inmuebles",
        "de_forma_ocasional_o_movil_modalidad_expres_sin_punto_fijo",
        "venta_de_drogas",
    },
}
# ==============================================================================
# PARTE 2: UTILIDADES DE NORMALIZACIÓN Y PROCESAMIENTO DE TEXTO
# ==============================================================================
# Esta sección contiene todas las funciones encargadas de transformar textos
# crudos en formas normalizadas aptas para comparación: eliminación de
# acentos, conversión a minúsculas, eliminación de espacios extras, creación
# de slugs (tokens sin caracteres especiales), corrección de errores
# ortográficos comunes, extracción de números de pregunta y funciones
# auxiliares de limpieza de display.
# ==============================================================================

# =========================================================
# UTILIDADES DE TEXTO
# =========================================================

def strip_accents(text: str) -> str:
    """Elimina todos los acentos y diacríticos de un texto.

    Convierte el texto a su forma NFD (decomposición) y filtra los
    caracteres de categoría "Mn" (Mark, Nonspacing), que corresponden
    a los diacríticos (tildes, dieresis, cedillas, etc.).

    Parámetros:
        text (str): Texto de entrada, puede contener acentos.

    Retorna:
        str: Texto sin acentos ni diacríticos.
    """
    text = unicodedata.normalize("NFD", str(text))
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def norm(text) -> str:
    """Normaliza un texto para comparación: minúsculas, sin acentos, sin
    saltos de línea, sin espacios extras.

    Procesamiento aplicado en orden:
    1. Convierte None a cadena vacía.
    2. Elimina el BOM (Byte Order Mark) si está presente.
    3. Reemplaza saltos de línea (\n, \r) por espacios.
    4. Elimina acentos mediante strip_accents.
    5. Convierte a minúsculas.
    6. Colapsa múltiples espacios en uno solo.
    7. Elimina espacios al inicio y al final.

    Parámetros:
        text: Valor de entrada (str, None, etc.).

    Retorna:
        str: Texto normalizado listo para comparación.
    """
    if text is None:
        return ""
    s = str(text).strip().strip("\ufeff")
    s = s.replace("\n", " ").replace("\r", " ")
    s = strip_accents(s).lower()
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def slugify(text) -> str:
    """Convierte un texto en un slug (token) apto para comparación.

    Procesamiento:
    1. Normaliza con norm().
    2. Reemplaza barras (/) por espacios.
    3. Elimina todo carácter que no sea alfanumérico ni espacio.
    4. Reemplaza espacios por guiones bajos (_).
    5. Colapsa múltiples guiones bajos en uno solo.
    6. Elimina guiones bajos al inicio y al final.

    Parámetros:
        text: Valor de entrada.

    Retorna:
        str: Slug limpio (ej.: "falta_de_iluminacion").
    """
    if text is None:
        return ""
    s = norm(text)
    s = s.replace("/", " ")
    s = re.sub(r"[^\w\s]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")


def normalize_option_token(text) -> str:
    """Alias de slugify, utilizado para normalizar tokens de opciones.

    Mantiene consistencia semántica: un "token de opción" es el slug
    de un texto de opción.

    Parámetros:
        text: Valor de entrada.

    Retorna:
        str: Token normalizado.
    """
    return slugify(text)


def extract_question_number(text: str) -> str:
    """Extrae el número de pregunta del inicio de un texto.

    Busca un patrón al inicio del texto que consista en uno o más dígitos,
    opcionalmente seguidos de un punto y más dígitos (para sub-preguntas
    como "12.1"), seguido opcionalmente de un punto o paréntesis cerrado.

    Ejemplos:
        "12. ¿Cuál es...?" -> "12"
        "12.1 Subpregunta" -> "12.1"
        "12) Opción" -> "12"

    Parámetros:
        text (str): Texto que comienza con el número de pregunta.

    Retorna:
        str: Número de pregunta extraído, o cadena vacía si no se encuentra.
    """
    s = str(text).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[\.\)]?", s)
    return m.group(1) if m else ""


def question_sort_key(q):
    """Genera una clave de ordenamiento para números de pregunta.

    Convierte el número de pregunta en una tupla donde cada parte
    separada por punto se convierte a entero (si es numérica) o se
    deja como cadena. Esto permite ordenar correctamente "2" antes
    que "12" (ya que int(2) < int(12)).

    Ejemplos:
        "12" -> (12,)
        "12.1" -> (12, 1)
        "3" -> (3,)

    Parámetros:
        q: Número de pregunta (str u otro convertible a str).

    Retorna:
        tuple: Tupla mixta de enteros y cadenas para ordenamiento.
    """
    s = str(q).strip()
    parts = s.split(".")
    out = []
    for p in parts:
        if str(p).isdigit():
            out.append(int(p))
        else:
            out.append(str(p))
    return tuple(out)


def is_effectively_empty(value) -> bool:
    """Determina si un valor es efectivamente vacío.

    Un valor se considera efectivamente vacío si es None o si su
    normalización (mediante norm()) está dentro del conjunto
    TOKENS_IGNORAR (cadena vacía, "nan", "none", "null").

    Parámetros:
        value: Valor a evaluar (str, None, etc.).

    Retorna:
        bool: True si el valor es efectivamente vacío, False en caso
              contrario.
    """
    if value is None:
        return True
    return norm(value) in TOKENS_IGNORAR


def clean_descriptor_display(text: str) -> str:
    """Limpia un texto de descriptor para su visualización.

    Únicamente colapsa múltiples espacios en uno solo y elimina
    espacios al inicio y al final. No elimina acentos ni convierte
    a minúsculas, preservando el texto legible original.

    Parámetros:
        text (str): Texto del descriptor.

    Retorna:
        str: Texto limpio para mostrar al usuario.
    """
    s = str(text).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def normalize_common_typos(token: str) -> str:
    """Corrije errores ortográficos comunes en tokens normalizados.

    Realiza reemplazos directos de cadena para variantes frecuentes:
    - "ocacional" -> "ocasional"
    - "extorcion"/"extorciones" -> "extorsion"/"extorsiones"
    - "gota_gota" -> "gota_a_gota" (en todas las combinaciones de
      préstamo/cobro, singular/plural)

    Parámetros:
        token (str): Token normalizado que puede contener errores.

    Retorna:
        str: Token con errores corregidos.
    """
    t = str(token)

    typo_map = {
        "ocacional": "ocasional",
        "extorcion": "extorsion",
        "extorciones": "extorsiones",
        "prestamo_gota_gota": "prestamo_gota_a_gota",
        "prestamos_gota_gota": "prestamos_gota_a_gota",
        "cobro_gota_gota": "cobro_gota_a_gota",
        "cobros_gota_gota": "cobros_gota_a_gota",
    }

    for bad, good in typo_map.items():
        t = t.replace(bad, good)

    return t


def normalize_token_for_compare(token: str) -> str:
    """Normaliza un token para comparación robusta.

    Aplica en cadena:
    1. normalize_option_token (slugify: sin acentos, minúsculas, solo
       alfanuméricos y guiones bajos).
    2. Elimina puntos, guiones y guiones bajos sobrantes al inicio/final.
    3. Corrije errores ortográficos comunes.
    4. Colapsa múltiples guiones bajos en uno solo.
    5. Elimina guiones bajos al inicio y al final.

    Este es el principal método de normalización para comparar tokens
    de opciones del CSV con aliases de descriptores.

    Parámetros:
        token (str): Token a normalizar.

    Retorna:
        str: Token completamente normalizado para comparación.
    """
    t = normalize_option_token(token)
    t = t.strip("._- ")
    t = normalize_common_typos(t)
    t = re.sub(r"_+", "_", t).strip("_")
    return t


def normalize_display_for_grouping(text: str) -> str:
    """Normaliza un texto para usarlo como etiqueta de grupo visual.

    A diferencia de normalize_token_for_compare, esta función preserva
    acentos y mayúsculas, pero:
    1. Colapsa múltiples espacios.
    2. Elimina signos de puntuación al final (.,;,:).
    3. Elimina espacios al inicio y al final.

    Se usa para generar las etiquetas legibles de los descriptores en
    los resultados.

    Parámetros:
        text (str): Texto del descriptor.

    Retorna:
        str: Texto limpio pero legible, con acentos.
    """
    s = clean_descriptor_display(text)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"[.;,:]+$", "", s).strip()
    return s
    # ==============================================================================
# PARTE 3: FUNCIONES DE IDENTIFICACIÓN Y CLASIFICACIÓN DE ARCHIVOS
# ==============================================================================
# Esta sección contiene la lógica para determinar automáticamente el tipo
# de archivo CSV (comunidad, comercio, policía o policial) basándose en
# el nombre del archivo. Esta clasificación es fundamental porque determina
# qué hoja del Excel guía se utiliza y qué reglas de conteo se aplican.
# ==============================================================================

# =========================================================
# TIPO DE ARCHIVO
# =========================================================

def infer_file_type(filename: str) -> str:
    """Infiere el tipo de archivo a partir de su nombre.

    Normaliza el nombre del archivo (minúsculas, sin acentos) y busca
    las palabras clave "comunidad", "comercio", "policial" o "policia"
    en él. El orden de evaluación importa: "policial" se evalúa antes
    que "policia" para evitar falsos positivos (ya que "policial" contiene
    "policia").

    Palabras clave y tipos retornados:
        - "comunidad" -> "comunidad"
        - "comercio"  -> "comercio"
        - "policial"  -> "policial"
        - "policia"   -> "policia"

    Parámetros:
        filename (str): Nombre del archivo (ej.: "CSV_Comunidad_Bogota.csv").

    Retorna:
        str: Tipo de archivo inferido, o cadena vacía si no se pudo
             determinar.
    """
    name = norm(filename)
    if "comunidad" in name:
        return "comunidad"
    if "comercio" in name:
        return "comercio"
    if "policial" in name:
        return "policial"
    if "policia" in name:
        return "policia"
    return ""
    # ==============================================================================
# PARTE 4: CARGA Y PROCESAMIENTO DEL EXCEL GUÍA
# ==============================================================================
# Esta sección se encarga de leer el archivo Excel que contiene la estructura
# de preguntas y descriptores para cada tipo de encuesta. El Excel puede tener
# múltiples hojas (Comunidad, Comercio, Policia), y dentro de cada hoja las
# preguntas están identificadas por un número al inicio de la celda, seguido
# de los descriptores asociados. La función parsea esta estructura en una
# lista de diccionarios normalizados, deduplicando entradas repetidas.
# También incluye una función para generar un resumen de lo detectado.
# ==============================================================================

# =========================================================
# EXCEL GUÍA
# =========================================================

def load_guide_excel(path_excel: str):
    """Carga el Excel guía y lo parsea en una estructura de diccionario por hoja.

    Para cada hoja del Excel, recorre las filas e identifica:
    - Filas de pregunta: aquellas cuya primera celda no vacía comienza
      con un número (detectado por extract_question_number).
    - Filas de descriptor: aquellas que no comienzan con número pero
      están debajo de una pregunta ya detectada.

    Cada descriptor se almacena como un diccionario con:
        - pregunta_num: número de pregunta (str)
        - pregunta_texto: texto completo de la pregunta (str)
        - pregunta_slug: slug del texto de la pregunta (str)
        - descriptor_texto: texto del descriptor limpio (str)
        - descriptor_slug: slug del descriptor (str)

    Se deduplican entradas usando como clave única la tupla:
        (pregunta_num, norm(pregunta_texto), norm(descriptor_texto),
         normalize_token_for_compare(descriptor_slug))

    Parámetros:
        path_excel (str): Ruta al archivo Excel guía.

    Retorna:
        dict: Diccionario {nombre_hoja: [lista_de_diccionarios]}.
    """
    xls = pd.ExcelFile(path_excel)
    guide = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(path_excel, sheet_name=sheet, header=None).fillna("")

        rows = []
        current_question_num = ""
        current_question_text = ""

        for _, row in df.iterrows():
            vals = [str(v).strip() for v in row.tolist()]
            vals_nonempty = [v for v in vals if v and norm(v) not in TOKENS_IGNORAR]

            if not vals_nonempty:
                continue

            first = vals_nonempty[0]
            qnum = extract_question_number(first)

            if qnum:
                # Es una fila de pregunta nueva
                current_question_num = qnum
                current_question_text = first

                # Los demás valores de la fila son descriptores de esta pregunta
                for desc in vals_nonempty[1:]:
                    desc_clean = str(desc).strip()
                    if desc_clean:
                        rows.append({
                            "pregunta_num": current_question_num,
                            "pregunta_texto": current_question_text,
                            "pregunta_slug": normalize_option_token(current_question_text),
                            "descriptor_texto": clean_descriptor_display(desc_clean),
                            "descriptor_slug": normalize_option_token(desc_clean),
                        })
                continue

            # No tiene número al inicio: es un descriptor de la pregunta actual
            if current_question_num and current_question_text:
                rows.append({
                    "pregunta_num": current_question_num,
                    "pregunta_texto": current_question_text,
                    "pregunta_slug": normalize_option_token(current_question_text),
                    "descriptor_texto": clean_descriptor_display(first),
                    "descriptor_slug": normalize_option_token(first),
                })

                # Valores adicionales en la misma fila también son descriptores
                for extra in vals_nonempty[1:]:
                    extra_clean = str(extra).strip()
                    if extra_clean:
                        rows.append({
                            "pregunta_num": current_question_num,
                            "pregunta_texto": current_question_text,
                            "pregunta_slug": normalize_option_token(current_question_text),
                            "descriptor_texto": clean_descriptor_display(extra_clean),
                            "descriptor_slug": normalize_option_token(extra_clean),
                        })

        # Deduplicación: eliminar entradas repetidas por clave única
        dedup = []
        seen = set()
        for r in rows:
            key = (
                r["pregunta_num"],
                norm(r["pregunta_texto"]),
                norm(r["descriptor_texto"]),
                normalize_token_for_compare(r["descriptor_slug"]),
            )
            if key not in seen and r["descriptor_slug"]:
                seen.add(key)
                dedup.append(r)

        guide[sheet] = dedup

    return guide


def build_guide_summary(guide: dict) -> pd.DataFrame:
    """Genera un resumen de lo detectado en el Excel guía.

    Para cada hoja del Excel guía, calcula:
        - hoja: nombre de la hoja (sin espacios extras)
        - preguntas_detectadas: cantidad de números de pregunta únicos
        - descriptores_detectados: cantidad total de descriptores

    Parámetros:
        guide (dict): Diccionario retornado por load_guide_excel.

    Retorna:
        pd.DataFrame: DataFrame con una fila por hoja del Excel.
    """
    rows = []
    for sh, items in guide.items():
        if not items:
            rows.append({
                "hoja": sh.strip(),
                "preguntas_detectadas": 0,
                "descriptores_detectados": 0,
            })
            continue

        df_tmp = pd.DataFrame(items)
        rows.append({
            "hoja": sh.strip(),
            "preguntas_detectadas": df_tmp["pregunta_num"].nunique(),
            "descriptores_detectados": len(df_tmp),
        })
    return pd.DataFrame(rows)
    # ==============================================================================
# PARTE 5: LECTURA Y PROCESAMIENTO DE ARCHIVOS CSV
# ==============================================================================
# Esta sección maneja la lectura robusta de archivos CSV que pueden venir en
# diferentes codificaciones (UTF-8 con/sin BOM, Latin-1) y con diferentes
# delimitadores (coma o punto y coma). Implementa dos estrategias de lectura:
# (1) un parser manual basado en el módulo csv de Python para manejar casos
# problemáticos, y (2) un fallback a pd.read_csv con múltiples combinaciones
# de parámetros. También incluye una función para aplanar encabezados
# multilinea del DataFrame resultante.
# ==============================================================================

# =========================================================
# CSV
# =========================================================

def parse_csv_with_python_engine(content: bytes, encoding: str, delimiter: str):
    """Parsea un CSV usando el módulo csv de Python con control total.

    Esta función implementa un parser manual que:
    1. Decodifica los bytes con la codificación dada (reemplazando
       caracteres inválidos).
    2. Usa csv.reader para dividir correctamente campos con comillas
       y delimitadores embebidos.
    3. Normaliza todas las filas a la misma cantidad de columnas
       (rellenando con cadenas vacías o truncando según sea necesario).
    4. Usa la primera fila como encabezado.

    Este enfoque es más robusto que pd.read_csv para archivos con
    formatos irregulares o delimitadores embebidos en campos citados.

    Parámetros:
        content (bytes): Contenido crudo del archivo CSV.
        encoding (str): Codificación a usar (ej.: "utf-8-sig").
        delimiter (str): Delimitador de campos ("," o ";").

    Retorna:
        pd.DataFrame: DataFrame con los datos parseados, o DataFrame
                      vacío si no se pudo leer.
    """
    text = content.decode(encoding, errors="replace")
    rows = []

    reader = csv.reader(io.StringIO(text), delimiter=delimiter, quotechar='"')

    max_cols = 0
    for row in reader:
        rows.append(row)
        max_cols = max(max_cols, len(row))

    if not rows or max_cols <= 1:
        return pd.DataFrame()

    # Normalizar todas las filas a la misma cantidad de columnas
    normalized_rows = []
    for row in rows:
        if len(row) < max_cols:
            row = row + [""] * (max_cols - len(row))
        elif len(row) > max_cols:
            row = row[:max_cols]
        normalized_rows.append(row)

    header = normalized_rows[0]
    data = normalized_rows[1:]

    if not header:
        return pd.DataFrame()

    return pd.DataFrame(data, columns=header).fillna("")


def try_read_csv_bytes(content: bytes) -> pd.DataFrame:
    """Intenta leer un CSV con múltiples estrategias de codificación y delimitador.

    Estrategia 1 - Parser manual (parse_csv_with_python_engine):
    Prueba 6 combinaciones de (encoding, delimiter):
        - (utf-8-sig, ","), (utf-8, ","), (latin-1, ",")
        - (utf-8-sig, ";"), (utf-8, ";"), (latin-1, ";")
    Retorna el primer DataFrame válido (más de 1 columna).

    Estrategia 2 - pd.read_csv como fallback:
    Si la estrategia 1 falla, prueba las mismas 6 combinaciones usando
    pd.read_csv con engine="python" y on_bad_lines="skip".

    Si ninguna estrategia funciona, lanza ValueError con el último error.

    Parámetros:
        content (bytes): Contenido crudo del archivo CSV.

    Retorna:
        pd.DataFrame: DataFrame con los datos del CSV.

    Excepciones:
        ValueError: Si no se pudo leer el CSV con ninguna estrategia.
    """
    attempts = [
        ("utf-8-sig", ","),
        ("utf-8", ","),
        ("latin-1", ","),
        ("utf-8-sig", ";"),
        ("utf-8", ";"),
        ("latin-1", ";"),
    ]

    last_error = None

    # Estrategia 1: Parser manual
    for encoding, delimiter in attempts:
        try:
            df = parse_csv_with_python_engine(content, encoding, delimiter)
            if not df.empty and df.shape[1] > 1:
                return df
        except Exception as e:
            last_error = e

    # Estrategia 2: pd.read_csv como fallback
    for at in [
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "utf-8"},
        {"sep": ",", "encoding": "latin-1"},
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "utf-8"},
        {"sep": ";", "encoding": "latin-1"},
    ]:
        try:
            df = pd.read_csv(
                io.BytesIO(content),
                dtype=str,
                keep_default_na=False,
                engine="python",
                on_bad_lines="skip",
                **at
            )
            if df.shape[1] > 1:
                return df.fillna("")
        except Exception as e:
            last_error = e

    raise ValueError(f"No se pudo leer el CSV. Error: {last_error}")


def flatten_headers(df: pd.DataFrame) -> pd.DataFrame:
    """Aplana encabezados multilinea del DataFrame.

    Reemplaza saltos de línea (\n, \r) dentro de los nombres de columnas
    por espacios simples y colapsa múltiples espacios en uno solo.
    Modifica el DataFrame in-place y lo retorna.

    Parámetros:
        df (pd.DataFrame): DataFrame con posibles encabezados multilinea.

    Retorna:
        pd.DataFrame: El mismo DataFrame con encabezados aplanados.
    """
    new_cols = []
    for c in df.columns:
        s = str(c).replace("\n", " ").replace("\r", " ").strip()
        s = re.sub(r"\s+", " ", s)
        new_cols.append(s)
    df.columns = new_cols
    return df
# ==============================================================================
# PARTE 6: UBICACIÓN DE PREGUNTAS EN EL CSV Y EXTRACCIÓN DE RESPUESTAS DE CELDAS
# ==============================================================================
# Esta sección aborda dos problemas centrales:
# (1) Dado un número y texto de pregunta del Excel guía, encontrar la columna
#     correspondiente en el CSV mediante un sistema de puntuación (scoring)
#     que compara múltiples estrategias de coincidencia.
# (2) Extraer y tokenizar las respuestas dentro de cada celda del CSV, que
#     pueden contener múltiples opciones separadas por comas, filtrando
#     opciones vacías, no productivas y de "no se observa".
# ==============================================================================

# =========================================================
# UBICAR PREGUNTA
# =========================================================

def build_question_groups(guide_sheet_rows: list):
    """Agrupa los descriptores por pregunta (número + texto).

    Toma la lista plana de diccionarios de una hoja del Excel guía y
    la agrupa en un defaultdict donde la clave es la tupla
    (pregunta_num, pregunta_texto) y el valor es la lista de todos los
    descriptores pertenecientes a esa pregunta.

    Parámetros:
        guide_sheet_rows (list): Lista de diccionarios con la estructura
            generada por load_guide_excel.

    Retorna:
        defaultdict: { (pregunta_num, pregunta_texto): [descriptor_dict, ...] }
    """
    grouped = defaultdict(list)
    for r in guide_sheet_rows:
        grouped[(r["pregunta_num"], r["pregunta_texto"])].append(r)
    return grouped


def score_question_column(col_name: str, question_num: str, question_text: str) -> int:
    """Calcula un puntaje de coincidencia entre una columna del CSV y una pregunta.

    Estrategias de coincidencia evaluadas (en orden de peso):
    1. Coincidencia exacta del texto completo normalizado (+120)
    2. Coincidencia exacta del slug completo (+120)
    3. Texto sin número contenido en la columna normalizada (+90)
    4. Slug sin número contenido en el slug de la columna (+90)
    5. Columna normalizada que inicia con el número de pregunta (+100)
    6. Número seguido de punto/paréntesis en la columna (+80)
    7. Número como palabra separada en la columna (+60)
    8. Texto completo contenido como subcadena (+50)
    9. Slug completo contenido como subcadena (+50)

    Los puntajes son acumulativos: una columna que coincida en múltiples
    estrategias obtiene un puntaje mayor.

    Parámetros:
        col_name (str): Nombre de la columna del CSV.
        question_num (str): Número de pregunta (ej.: "12").
        question_text (str): Texto completo de la pregunta del Excel guía.

    Retorna:
        int: Puntaje de coincidencia (0 si no hay coincidencia significativa).
    """
    col_norm = norm(col_name)
    col_slug = normalize_option_token(col_name)

    score = 0
    qnum = str(question_num).strip()
    qtext_norm = norm(question_text)
    qtext_slug = normalize_option_token(question_text)

    # Extraer el texto sin el número inicial para comparaciones parciales
    text_wo_num = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\.\)]?\s*", "", question_text).strip()
    text_wo_num_norm = norm(text_wo_num)
    text_wo_num_slug = normalize_option_token(text_wo_num)

    # Estrategia 5: columna que inicia con el número
    if qnum:
        if col_norm.startswith(qnum):
            score += 100
        if f"{qnum}." in col_norm or f"{qnum})" in col_norm:
            score += 80
        if re.search(rf"(^|\s){re.escape(qnum)}(\.|\)|\s|$)", col_norm):
            score += 60

    # Estrategia 1: coincidencia exacta del texto normalizado
    if qtext_norm and qtext_norm == col_norm:
        score += 120
    # Estrategia 2: coincidencia exacta del slug
    if qtext_slug and qtext_slug == col_slug:
        score += 120

    # Estrategia 3: texto sin número contenido en la columna
    if text_wo_num_norm and text_wo_num_norm in col_norm:
        score += 90
    # Estrategia 4: slug sin número contenido en el slug de la columna
    if text_wo_num_slug and text_wo_num_slug in col_slug:
        score += 90

    # Estrategia 8-9: texto/slug como subcadena
    if qtext_norm and qtext_norm in col_norm:
        score += 50
    if qtext_slug and qtext_slug in col_slug:
        score += 50

    return score


def find_question_column(df: pd.DataFrame, question_num: str, question_text: str):
    """Encuentra la columna del CSV que mejor coincide con una pregunta del Excel guía.

    Evalúa todas las columnas del DataFrame usando score_question_column
    y retorna la de mayor puntaje. Si el mejor puntaje es menor a 50,
    se considera que no se encontró una columna confiable.

    Parámetros:
        df (pd.DataFrame): DataFrame del CSV con encabezados aplanados.
        question_num (str): Número de pregunta.
        question_text (str): Texto completo de la pregunta.

    Retorna:
        tuple: (nombre_columna, puntaje). Si no hay coincidencia suficiente,
               retorna (None, puntaje).
    """
    best_col = None
    best_score = -1

    for col in df.columns:
        score = score_question_column(col, question_num, question_text)
        if score > best_score:
            best_score = score
            best_col = col

    if best_score < 50:
        return None, best_score

    return best_col, best_score


# =========================================================
# RESPUESTAS EN CELDAS
# =========================================================

def split_multiselect_cell(value: str):
    """Divide el contenido de una celda en opciones individuales.

    Las celdas del CSV pueden contener múltiples opciones separadas por
    comas. Esta función:
    1. Retorna lista vacía si el valor es efectivamente vacío.
    2. Divide por coma.
    3. Elimina espacios alrededor de cada parte.
    4. Filtra partes vacías o que normalizadas estén en TOKENS_IGNORAR.

    Parámetros:
        value (str): Contenido de la celda.

    Retorna:
        list: Lista de strings con las opciones individuales.
    """
    if is_effectively_empty(value):
        return []

    raw = str(value).strip()
    if not raw:
        return []

    parts = [p.strip() for p in raw.split(",")]
    parts = [p for p in parts if p and norm(p) not in TOKENS_IGNORAR]
    return parts


def is_no_observa_option(token_norm: str) -> bool:
    """Determina si un token representa una opción de "no se observa".

    Busca cualquiera de los patrones en PATRONES_NO_OBSERVA como
    subcadena del token normalizado.

    Parámetros:
        token_norm (str): Token normalizado.

    Retorna:
        bool: True si contiene un patrón de "no se observa".
    """
    for p in PATRONES_NO_OBSERVA:
        if p in token_norm:
            return True
    return False


def is_unproductive_option(token_norm: str, file_type: str = "", question_num: str = "") -> bool:
    """Determina si un token es no productivo (debe excluirse del conteo).

    Criterios de exclusión (en orden):
    1. Si está en OPCIONES_NO_PRODUCTIVAS -> sí es no productivo.
    2. Si comienza con "otro_" o "otros_" -> sí es no productivo.
    3. Si contiene un patrón de PATRONES_NO_OBSERVA -> sí es no productivo.
       Esto aplica para TODAS las preguntas sin excepción, incluyendo
       "no_se_observan_delitos_ambientales" en comunidad 27.
    4. En caso contrario -> no es no productivo.

    Parámetros:
        token_norm (str): Token normalizado.
        file_type (str): Tipo de archivo (se conserva por compatibilidad
                         de firma, pero ya no se usa para excepciones).
        question_num (str): Número de pregunta (se conserva por
                            compatibilidad de firma, pero ya no se usa
                            para excepciones).

    Retorna:
        bool: True si el token debe excluirse del conteo.
    """
    if token_norm in OPCIONES_NO_PRODUCTIVAS:
        return True
    if token_norm.startswith("otro_") or token_norm.startswith("otros_"):
        return True
    if is_no_observa_option(token_norm):
        return True
    return False


def tokenize_cell_unique(value: str, file_type: str = "", question_num: str = ""):
    """Extrae tokens únicos y productivos de una celda.

    Procesamiento:
    1. Divide la celda en opciones con split_multiselect_cell.
    2. Normaliza cada opción con normalize_token_for_compare.
    3. Filtra tokens vacíos o que estén en TOKENS_IGNORAR.
    4. Filtra tokens no productivos con is_unproductive_option.
    5. Retorna tokens únicos ordenados alfabéticamente.

    Parámetros:
        value (str): Contenido de la celda.
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.

    Retorna:
        list: Lista ordenada de tokens únicos productivos.
    """
    options = split_multiselect_cell(value)
    tokens = []

    for opt in options:
        token = normalize_token_for_compare(opt)
        if not token or token in TOKENS_IGNORAR:
            continue
        if is_unproductive_option(token, file_type=file_type, question_num=str(question_num)):
            continue
        tokens.append(token)

    return sorted(set(tokens))
# ==============================================================================
# PARTE 7: REGLAS DE CONTEO Y AGRUPACIÓN DE DESCRIPTORES
# ==============================================================================
# Esta sección define la lógica central de agrupación y conteo:
# - Determina si una pregunta se cuenta en modo "exacto" (descriptor por
#   descriptor) o "unificado" (una sola etiqueta agregada).
# - Construye los aliases (sinónimos) para cada descriptor, incluyendo
#   un mapa extenso de equivalencias para unir variantes como
#   "consumo_de_drogas" y "consumo_de_drogas_en_espacios_publicos".
# - Define los grupos canónicos para preguntas exactas, unificando
#   descriptores que representan el mismo concepto.
# - Implementa dos modos de conteo: exacto (cuenta cada token que coincide)
#   y unificado (cuenta 1 por fila si al menos un token coincide).
# ==============================================================================

# =========================================================
# REGLAS DE CONTEO
# =========================================================

def is_exact_question(file_type: str, question_num: str) -> bool:
    """Determina si una pregunta debe contarse en modo exacto.

    Consulta PREGUNTAS_EXACTAS para el tipo de archivo dado:
    - Si el valor es None, TODAS las preguntas son exactas.
    - Si es un conjunto, solo las preguntas cuyo número esté en el
      conjunto son exactas; las demás son unificadas.

    Parámetros:
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.

    Retorna:
        bool: True si la pregunta es exacta, False si es unificada.
    """
    exacts = PREGUNTAS_EXACTAS.get(file_type)
    if exacts is None:
        return True
    return question_num in exacts


def get_unified_label(file_type: str, question_num: str, question_text: str) -> str:
    """Obtiene la etiqueta unificada para una pregunta no exacta.

    Busca en UNIFIED_LABELS usando la clave (file_type, question_num).
    Si no encuentra una etiqueta predefinida, genera una a partir del
    texto de la pregunta sin el número inicial.

    Parámetros:
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.
        question_text (str): Texto completo de la pregunta.

    Retorna:
        str: Etiqueta legible para el grupo unificado.
    """
    if (file_type, question_num) in UNIFIED_LABELS:
        return UNIFIED_LABELS[(file_type, question_num)]

    text_wo_num = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\.\)]?\s*", "", question_text).strip()
    return clean_descriptor_display(text_wo_num) if text_wo_num else f"Pregunta {question_num}"


def build_descriptor_aliases(file_type: str, question_num: str, descriptor_text: str):
    """Construye el conjunto de aliases (sinónimos) para un descriptor.

    Proceso:
    1. Genera el token base del descriptor.
    2. Busca el token base en alias_map (mapa extenso de equivalencias
       que incluye consumo de drogas, contaminación sónica, alumbrado,
       prostitución, extorsión, oferta educativa/laboral, infraestructura
       vial, préstamos gota a gota, delitos ambientales, estafa/fraude,
       asaltos por tipo de objetivo, etc.).
    3. Agrega aliases extra de UNIFIED_EXTRA_ALIASES si aplica.
    4. Para preguntas de estafa (comunidad 23, comercio 21), agrega
       un conjunto amplio de variantes de estafa y fraude.
    5. Para comercio 20, agrega variantes de asalto según el tipo
       (persona, comercio, vivienda, transporte).
    6. Filtra aliases que sean no productivos.

    NOTA: "no_se_observan_delitos_ambientales" NO se incluye como alias
    porque es una opción no productiva que debe ser filtrada por
    is_unproductive_option.

    Parámetros:
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.
        descriptor_text (str): Texto del descriptor.

    Retorna:
        set: Conjunto de tokens normalizados que son sinónimos del descriptor.
    """
    base = normalize_option_token(descriptor_text)
    base_norm = normalize_token_for_compare(base)
    aliases = {base_norm}

    # Mapa extenso de equivalencias entre descriptores
    alias_map = {
        "consumo_de_drogas": {
            "consumo_de_drogas",
            "consumo_de_drogas_en_espacios_publicos",
        },
        "consumo_de_drogas_en_espacios_publicos": {
            "consumo_de_drogas_en_espacios_publicos",
            "consumo_de_drogas",
        },
        "contaminacion_sonica": {
            "escandalos_musicales_o_ruidos_excesivos",
            "contaminacion_sonica",
        },
        "escandalos_musicales_o_ruidos_excesivos": {
            "escandalos_musicales_o_ruidos_excesivos",
            "contaminacion_sonica",
        },
        "carencia_o_inexistencia_de_alumbrado_publico": {
            "carencia_o_inexistencia_de_alumbrado_publico",
            "deficiencias_en_el_alumbrado_publico",
        },
        "deficiencias_en_el_alumbrado_publico": {
            "deficiencias_en_el_alumbrado_publico",
            "carencia_o_inexistencia_de_alumbrado_publico",
        },
        "presencia_de_personas_en_situacion_de_calle": {
            "presencia_de_personas_en_situacion_de_calle",
            "presencia_de_personas_en_situacion_de_calle_personas_que_viven_permanentemente_en_la_via_publica",
            "personas_en_situacion_de_calle",
        },
        "presencia_de_personas_en_situacion_de_calle_personas_que_viven_permanentemente_en_la_via_publica": {
            "presencia_de_personas_en_situacion_de_calle_personas_que_viven_permanentemente_en_la_via_publica",
            "presencia_de_personas_en_situacion_de_calle",
            "personas_en_situacion_de_calle",
        },

        "zona_donde_se_ejerce_prostitucion": {
            "zona_donde_se_ejerce_prostitucion",
            "zonas_donde_se_ejerce_prostitucion",
            "zona_de_prostitucion",
            "zonas_de_prostitucion",
            "prostitucion",
            "ejercicio_de_la_prostitucion",
            "zona_donde_hay_prostitucion",
            "sitios_donde_se_ejerce_prostitucion",
        },

        "extorsion": {
            "extorsion",
            "extorcion",
            "extorsiones",
            "extorciones",
            "cobro_ilegal",
            "cobro_ilegal_a_comercios",
            "exigencias_indebidas",
            "exigencias_ilegales",
            "intimidacion_para_exigir_cobro",
            "amenazas_o_intimidacion_para_exigir_cobro_de_dinero_u_otros_beneficios_de_manera_ilegal_a_comercios",
        },

        "falta_de_oferta_educativa": {
            "falta_de_oferta_educativa",
            "falta_de_oferta_laboral",
            "falta_de_oferta_recreativa",
            "falta_de_actividades_culturales",
        },
        "falta_de_oferta_laboral": {
            "falta_de_oferta_educativa",
            "falta_de_oferta_laboral",
            "falta_de_oferta_recreativa",
            "falta_de_actividades_culturales",
        },
        "falta_de_oferta_recreativa": {
            "falta_de_oferta_educativa",
            "falta_de_oferta_laboral",
            "falta_de_oferta_recreativa",
            "falta_de_actividades_culturales",
        },
        "falta_de_actividades_culturales": {
            "falta_de_oferta_educativa",
            "falta_de_oferta_laboral",
            "falta_de_oferta_recreativa",
            "falta_de_actividades_culturales",
        },

        "calles_en_mal_estado": {
            "calles_en_mal_estado",
            "falta_de_iluminacion",
            "falta_de_senalizacion",
            "falta_o_deterioro_de_aceras",
        },
        "falta_de_iluminacion": {
            "calles_en_mal_estado",
            "falta_de_iluminacion",
            "falta_de_senalizacion",
            "falta_o_deterioro_de_aceras",
        },
        "falta_de_senalizacion": {
            "calles_en_mal_estado",
            "falta_de_iluminacion",
            "falta_de_senalizacion",
            "falta_o_deterioro_de_aceras",
        },
        "falta_o_deterioro_de_aceras": {
            "calles_en_mal_estado",
            "falta_de_iluminacion",
            "falta_de_senalizacion",
            "falta_o_deterioro_de_aceras",
        },

        # Policial: préstamos gota a gota
        "prestamos_gota_a_gota": {
            "prestamos_gota_a_gota",
            "prestamo_gota_a_gota",
            "gota_a_gota",
            "prestamos_tipo_gota_a_gota",
            "prestamos_gota_gota",
            "prestamo_gota_gota",
            "cobro_gota_a_gota",
            "cobros_gota_a_gota",
        },
        "prestamo_gota_a_gota": {
            "prestamos_gota_a_gota",
            "prestamo_gota_a_gota",
            "gota_a_gota",
            "prestamos_tipo_gota_a_gota",
            "prestamos_gota_gota",
            "prestamo_gota_gota",
            "cobro_gota_a_gota",
            "cobros_gota_a_gota",
        },
        "gota_a_gota": {
            "prestamos_gota_a_gota",
            "prestamo_gota_a_gota",
            "gota_a_gota",
            "prestamos_tipo_gota_a_gota",
            "prestamos_gota_gota",
            "prestamo_gota_gota",
            "cobro_gota_a_gota",
            "cobros_gota_a_gota",
        },

        # Comunidad 27: delitos ambientales (SOLO variantes de contaminación de aguas,
        # NO se incluye "no_se_observan_delitos_ambientales" porque es opción no productiva)
        "envenenamiento_de_aguas": {
            "envenenamiento_de_aguas",
            "envenenamiento_o_contaminacion_de_aguas",
            "contaminacion_de_aguas",
            "contaminacion_o_envenenamiento_de_aguas",
        },
        "envenenamiento_o_contaminacion_de_aguas": {
            "envenenamiento_de_aguas",
            "envenenamiento_o_contaminacion_de_aguas",
            "contaminacion_de_aguas",
            "contaminacion_o_envenenamiento_de_aguas",
        },
        "contaminacion_de_aguas": {
            "envenenamiento_de_aguas",
            "envenenamiento_o_contaminacion_de_aguas",
            "contaminacion_de_aguas",
            "contaminacion_o_envenenamiento_de_aguas",
        },

        "ventas_informales_ambulantes": {"ventas_informales_ambulantes"},
        "problemas_vecinales_o_conflictos_entre_vecinos": {"problemas_vecinales_o_conflictos_entre_vecinos"},
        "desvinculacion_escolar_desercion_escolar": {"desvinculacion_escolar_desercion_escolar"},
        "perdida_de_espacios_publicos_parques_polideportivos_u_otros": {"perdida_de_espacios_publicos_parques_polideportivos_u_otros"},
        "acumulacion_de_basura_aguas_negras_o_mal_alcantarillado": {"acumulacion_de_basura_aguas_negras_o_mal_alcantarillado"},
        "falta_de_oportunidades_laborales": {"falta_de_oportunidades_laborales"},
        "asentamientos_informales_o_precarios": {"asentamientos_informales_o_precarios"},
        "lotes_baldios": {"lotes_baldios"},
        "cuarterias": {"cuarterias"},
        "consumo_de_alcohol_en_via_publica": {"consumo_de_alcohol_en_via_publica"},
        "hurto": {"hurto"},
        "hurto_simple": {"hurto_simple", "hurto"},
    }

    if base_norm in alias_map:
        aliases.update(alias_map[base_norm])

    # Refuerzos extra para preguntas unificadas
    if (file_type, question_num) in UNIFIED_EXTRA_ALIASES:
        aliases.update(UNIFIED_EXTRA_ALIASES[(file_type, question_num)])

    # Estafa: ampliar aliases para comunidad 23 y comercio 21
    if (
        (file_type == "comunidad" and question_num == "23")
        or (file_type == "comercio" and question_num == "21")
    ):
        if "estafa" in base_norm or "fraude" in base_norm:
            aliases.update({
                "estafa",
                "estafa_informatica",
                "fraude",
                "fraude_informatico",
                "estafa_por_medios_informaticos",
                "estafa_por_internet",
                "estafa_digital",
                "estafa_electronica",
                "estafa_telefonica",
                "estafa_bancaria",
                "estafa_por_redes_sociales",
                "estafa_documental",
                "estafa_simple",
                "estafa_en_compras",
                "estafa_comercial",
                "estafa_por_medio_electronico",
                "fraude_electronico",
                "fraude_digital",
            })

    # Comercio 20: ampliar aliases según tipo de asalto
    if file_type == "comercio" and question_num == "20":
        if "persona" in base_norm:
            aliases.update({
                "asalto_a_personas",
                "asalto_personas",
                "asalto_a_peatones",
                "asalto_peatones",
                "robo_a_persona",
                "robo_a_personas",
                "robo_persona",
                "robo_personas",
            })
        if "comercio" in base_norm:
            aliases.update({
                "asalto_a_comercio",
                "asalto_a_comercios",
                "asalto_comercio",
                "asalto_comercios",
                "asalto_a_locales_comerciales",
                "robo_a_comercio",
                "robo_a_comercios",
                "robo_comercio",
                "robo_comercios",
                "robo_a_comercio_intimidacion",
            })
        if "vivienda" in base_norm or "casa" in base_norm:
            aliases.update({
                "asalto_a_vivienda",
                "asalto_a_viviendas",
                "asalto_vivienda",
                "asalto_viviendas",
                "asalto_a_casa",
                "asalto_a_casas",
                "robo_a_vivienda",
                "robo_a_viviendas",
                "robo_vivienda",
                "robo_viviendas",
                "robo_a_casa",
                "robo_a_casas",
            })
        if "transporte" in base_norm:
            aliases.update({
                "asalto_a_transporte_publico",
                "asalto_transporte_publico",
                "asalto_en_transporte_publico",
                "asalto_bus",
                "asalto_autobus",
                "robo_a_transporte_publico",
                "robo_transporte_publico",
                "robo_a_transporte_publico_con_intimidacion",
                "robo_bus",
                "robo_autobus",
            })

    # Filtrar aliases no productivos
    aliases = {normalize_token_for_compare(a) for a in aliases if a}
    aliases = {
        a for a in aliases
        if not is_unproductive_option(a, file_type=file_type, question_num=str(question_num))
    }
    return aliases


def get_exact_canonical_group(file_type: str, question_num: str, descriptor_text: str):
    """Determina el grupo canónico y modo de agrupación para un descriptor en pregunta exacta.

    Para ciertos descriptores que representan el mismo concepto pero tienen
    textos diferentes en el Excel guía, esta función los une bajo una
    etiqueta común con modo "merged". Para los demás, retorna la etiqueta
    original con modo "exact".

    Reglas de fusión específicas:
    - Comunidad 12: consumo de drogas, contaminación sónica, alumbrado,
      prostitución.
    - Comercio 18: extorsión (variantes con "extors"/"extorc"/"exigencias").
    - Comercio 20: asaltos por tipo (persona, comercio, vivienda, transporte).
    - Comunidad 27: envenenamiento/contaminación de aguas.
      NOTA: "no_se_observan_delitos_ambientales" NO se agrupa aquí porque
      es una opción no productiva que se filtra antes del conteo.
    - Policial/Policía: préstamos gota a gota (cualquier variante).

    Parámetros:
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.
        descriptor_text (str): Texto del descriptor.

    Retorna:
        tuple: (etiqueta_canonica, modo) donde modo es "exact" o "merged".
    """
    base = normalize_token_for_compare(descriptor_text)
    label = normalize_display_for_grouping(descriptor_text)
    group_mode = "exact"

    # Comunidad 12: fusionar variantes
    if file_type == "comunidad" and question_num == "12":
        if base in {"consumo_de_drogas", "consumo_de_drogas_en_espacios_publicos"}:
            return "Consumo de drogas", "merged"
        if base in {"contaminacion_sonica", "escandalos_musicales_o_ruidos_excesivos"}:
            return "Contaminación sónica", "merged"
        if base in {
            "carencia_o_inexistencia_de_alumbrado_publico",
            "deficiencias_en_el_alumbrado_publico"
        }:
            return "Carencia o inexistencia de alumbrado público", "merged"
        if base in {
            "zona_donde_se_ejerce_prostitucion",
            "zonas_donde_se_ejerce_prostitucion",
            "zona_de_prostitucion",
            "zonas_de_prostitucion",
            "ejercicio_de_la_prostitucion",
            "prostitucion",
        }:
            return "Zona donde se ejerce prostitución", "merged"

    # Comercio 18: fusionar variantes de extorsión
    if file_type == "comercio" and question_num == "18":
        if (
            "extors" in base
            or "extorc" in base
            or "exigencias_indebidas" in base
            or "cobro_ilegal" in base
        ):
            return "Extorsión (amenazas o intimidación para exigir cobro de dinero u otros beneficios de manera ilegal a comercios)", "merged"

    # Comercio 20: fusionar asaltos por tipo de objetivo
    if file_type == "comercio" and question_num == "20":
        if ("persona" in base) or ("peaton" in base):
            return "Asalto a personas", "merged"
        if "comerc" in base:
            return "Asalto a comercio", "merged"
        if ("vivienda" in base) or ("casa" in base):
            return "Asalto a vivienda", "merged"
        if ("transporte" in base) or ("bus" in base) or ("autobus" in base):
            return "Asalto a transporte público", "merged"

    # Comunidad 27: unificar variantes de contaminación de aguas.
    # "no_se_observan_delitos_ambientales" NO se incluye aquí:
    # es opción no productiva y se filtra por is_unproductive_option.
    if file_type == "comunidad" and question_num == "27":
        if base in {
            "envenenamiento_de_aguas",
            "envenenamiento_o_contaminacion_de_aguas",
            "contaminacion_de_aguas",
            "contaminacion_o_envenenamiento_de_aguas",
        }:
            return "Envenenamiento o contaminación de aguas", "merged"

    # Policial / Policía: unificar gota a gota y evitar duplicados visuales
    if file_type in {"policial", "policia"}:
        if (
            "gota_a_gota" in base
            or "prestamo_gota_a_gota" in base
            or "prestamos_gota_a_gota" in base
            or "cobro_gota_a_gota" in base
            or "cobros_gota_a_gota" in base
        ):
            return "Préstamos gota a gota", "merged"
        return normalize_display_for_grouping(descriptor_text), "exact"

    return label, group_mode


def build_group_definitions(file_type: str, question_num: str, question_text: str, items: list):
    """Construye las definiciones de grupos para una pregunta.

    Para preguntas unificadas:
    - Crea un único grupo con la etiqueta de UNIFIED_LABELS.
    - Los aliases son la unión de todos los aliases de cada descriptor
      del Excel guía más los aliases extra de UNIFIED_EXTRA_ALIASES.

    Para preguntas exactas:
    - Crea un grupo por cada descriptor (o grupo canónico fusionado).
    - Cada grupo tiene sus propios aliases construidos con
      build_descriptor_aliases.

    Retorna:
        dict: {etiqueta_grupo: {group_label, aliases, source_descriptors, group_mode}}
    """
    if not is_exact_question(file_type, question_num):
        unified_label = get_unified_label(file_type, question_num, question_text)
        aliases = set()

        for item in items:
            aliases.update(build_descriptor_aliases(file_type, question_num, item["descriptor_texto"]))

        aliases.update(UNIFIED_EXTRA_ALIASES.get((file_type, question_num), set()))

        return {
            unified_label: {
                "group_label": unified_label,
                "aliases": aliases,
                "source_descriptors": {item["descriptor_texto"] for item in items},
                "group_mode": "unified",
            }
        }

    groups = {}
    for item in items:
        label, group_mode = get_exact_canonical_group(file_type, question_num, item["descriptor_texto"])

        if label not in groups:
            groups[label] = {
                "group_label": label,
                "aliases": set(),
                "source_descriptors": set(),
                "group_mode": group_mode,
            }

        groups[label]["aliases"].update(
            build_descriptor_aliases(file_type, question_num, item["descriptor_texto"])
        )
        groups[label]["source_descriptors"].add(item["descriptor_texto"])

        if group_mode == "merged":
            groups[label]["group_mode"] = "merged"

    return groups


def count_group_exact(series: pd.Series, aliases: set, file_type: str = "", question_num: str = ""):
    """Cuenta ocurrencias en modo exacto (cada token que coincide suma 1).

    Para cada celda de la serie:
    1. Tokeniza con tokenize_cell_unique.
    2. Identifica tokens que están en el conjunto de aliases.
    3. Suma la cantidad de tokens coincidentes (no 1 por fila).

    Parámetros:
        series (pd.Series): Columna del CSV con las respuestas.
        aliases (set): Conjunto de tokens que definen el grupo.
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.

    Retorna:
        tuple: (total_int, Counter_de_tokens_coincidentes)
    """
    total = 0
    matched_tokens = Counter()

    for val in series:
        row_tokens = tokenize_cell_unique(val, file_type=file_type, question_num=str(question_num))
        row_hit_tokens = set()

        for token in row_tokens:
            if token in aliases:
                row_hit_tokens.add(token)

        total += len(row_hit_tokens)
        for t in row_hit_tokens:
            matched_tokens[t] += 1

    return total, matched_tokens


def count_group_unified(series: pd.Series, aliases: set, file_type: str = "", question_num: str = ""):
    """Cuenta ocurrencias en modo unificado (1 por fila si hay al menos un match).

    Para cada celda de la serie:
    1. Tokeniza con tokenize_cell_unique.
    2. Identifica tokens que están en el conjunto de aliases.
    3. Si al menos un token coincide, suma 1 (no importa cuántos tokens
       coincidan en la misma fila).

    Parámetros:
        series (pd.Series): Columna del CSV con las respuestas.
        aliases (set): Conjunto de tokens que definen el grupo.
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.

    Retorna:
        tuple: (total_int, Counter_de_tokens_coincidentes)
    """
    total = 0
    matched_tokens = Counter()

    for val in series:
        row_tokens = tokenize_cell_unique(val, file_type=file_type, question_num=str(question_num))
        row_hit_tokens = set()

        for token in row_tokens:
            if token in aliases:
                row_hit_tokens.add(token)

        if row_hit_tokens:
            total += 1
            for t in row_hit_tokens:
                matched_tokens[t] += 1

    return total, matched_tokens


def find_unmapped_tokens(series: pd.Series, matched_aliases_union: set, file_type: str = "", question_num: str = ""):
    """Encuentra tokens del CSV que no coincidieron con ningún grupo.

    Para cada celda de la serie:
    1. Tokeniza con tokenize_cell_unique.
    2. Si no hay tokens, cuenta como fila vacía.
    3. Filtra tokens que no estén en matched_aliases_union y que no sean
       no productivos.
    4. Cuenta frecuencia de cada token no ubicado.

    Útil para auditoría: identifica opciones del CSV que no fueron
    cubiertas por ningún descriptor del Excel guía.

    Parámetros:
        series (pd.Series): Columna del CSV con las respuestas.
        matched_aliases_union (set): Unión de todos los aliases de todos
            los grupos de la pregunta.
        file_type (str): Tipo de archivo.
        question_num (str): Número de pregunta.

    Retorna:
        tuple: (Counter_de_tokens_no_ubicados, cantidad_filas_vacias)
    """
    counter = Counter()
    blank_rows = 0

    for val in series:
        row_tokens = tokenize_cell_unique(val, file_type=file_type, question_num=str(question_num))

        if not row_tokens:
            blank_rows += 1
            continue

        for token in row_tokens:
            if token not in matched_aliases_union and not is_unproductive_option(token, file_type=file_type, question_num=str(question_num)):
                counter[token] += 1

    return counter, blank_rows
# ==============================================================================
# PARTE 8: REGLAS ESPECIALES CRUZADAS ENTRE PREGUNTAS Y FUNCIÓN PRINCIPAL DE
#          PROCESAMIENTO
# ==============================================================================
# Esta sección contiene:
# (1) Una regla especial que traslada conteos de extorsión de la pregunta 18
#     de comercio a la pregunta 12 de comercio, cuando los descriptores
#     coinciden con patrones específicos de cobro ilegal/exigencias indebidas.
# (2) La función build_results_for_file, que es el motor principal: orquesta
#     la ubicación de columnas, construcción de grupos, refuerzo dinámico
#     de aliases (basado en lo que realmente aparece en el CSV), conteo y
#     detección de opciones no ubicadas para un archivo completo.
# ==============================================================================

# =========================================================
# REGLA ESPECIAL: COMERCIO 18 -> COMERCIO 12
# =========================================================

def is_comercio_q18_extorsion_descriptor(desc: str) -> bool:
    """Identifica descriptores de extorsión en pregunta 18 de comercio.

    Un descriptor se considera de extorsión si su texto normalizado
    contiene "extorsion", "extorcion", o la combinación de "amenazas"
    con "intimidacion" y "exigir cobro".

    Parámetros:
        desc (str): Texto del descriptor.

    Retorna:
        bool: True si es un descriptor de extorsión.
    """
    d = norm(desc)
    return (
        "extorsion" in d
        or "extorcion" in d
        or (
            "amenazas" in d
            and "intimidacion" in d
            and "exigir cobro" in d
        )
    )


def is_comercio_q12_cobro_ilegal_descriptor(desc: str) -> bool:
    """Identifica descriptores de cobro ilegal en pregunta 12 de comercio.

    Un descriptor coincide si contiene "intentos de cobro ilegal" o
    "cobro ilegal", Y además contiene "exigencias indebidas" o
    "zona comercial" o "comercial".

    Parámetros:
        desc (str): Texto del descriptor.

    Retorna:
        bool: True si es un descriptor de cobro ilegal en zona comercial.
    """
    d = norm(desc)
    return (
        (
            "intentos de cobro ilegal" in d
            or "cobro ilegal" in d
        )
        and (
            "exigencias indebidas" in d
            or "zona comercial" in d
            or "comercial" in d
        )
    )


def apply_special_cross_question_rules(df_results: pd.DataFrame) -> pd.DataFrame:
    """Aplica la regla especial de traslado de comercio 18 a comercio 12.

    Lógica:
    1. Identifica filas de comercio pregunta 18 con descriptores de extorsión.
    2. Calcula el total de respuestas a trasladar.
    3. Busca filas de comercio pregunta 12 con descriptores de cobro ilegal.
    4. Si existe destino: suma el traslado y fusiona los tokens CSV.
    5. Si no existe destino: crea una nueva fila con el traslado.
    6. Elimina las filas origen (pregunta 18 extorsión).

    Parámetros:
        df_results (pd.DataFrame): DataFrame de resultados completos.

    Retorna:
        pd.DataFrame: DataFrame con el traslado aplicado.
    """
    if df_results.empty:
        return df_results

    df_results = df_results.copy()

    mask_src = (
        (df_results["tipo"] == "comercio")
        & (df_results["pregunta_num"].astype(str) == "18")
        & (df_results["descriptor"].astype(str).apply(is_comercio_q18_extorsion_descriptor))
    )

    if not mask_src.any():
        return df_results

    src_rows = df_results[mask_src].copy()
    traslado_total = int(src_rows["cantidad_respuestas"].sum())

    if traslado_total <= 0:
        return df_results

    mask_dst = (
        (df_results["tipo"] == "comercio")
        & (df_results["pregunta_num"].astype(str) == "12")
        & (df_results["descriptor"].astype(str).apply(is_comercio_q12_cobro_ilegal_descriptor))
    )

    # Recopilar tokens CSV de las filas origen
    src_tokens = []
    for txt in src_rows["opciones_csv_que_contaron"].fillna("").astype(str).tolist():
        if txt.strip():
            src_tokens.extend([p.strip() for p in txt.split("|") if p.strip()])
    src_tokens = sorted(set(src_tokens))

    if mask_dst.any():
        # Sumar al destino existente
        dst_idx = df_results[mask_dst].index[0]
        df_results.at[dst_idx, "cantidad_respuestas"] = int(df_results.at[dst_idx, "cantidad_respuestas"]) + traslado_total

        current_tokens = str(df_results.at[dst_idx, "opciones_csv_que_contaron"] or "").strip()
        dst_tokens = [p.strip() for p in current_tokens.split("|") if p.strip()] if current_tokens else []
        merged_tokens = sorted(set(dst_tokens + src_tokens))
        df_results.at[dst_idx, "opciones_csv_que_contaron"] = " | ".join(merged_tokens)
    else:
        # Crear nueva fila destino
        q12_rows = df_results[
            (df_results["tipo"] == "comercio")
            & (df_results["pregunta_num"].astype(str) == "12")
        ].copy()

        if not q12_rows.empty:
            ref = q12_rows.iloc[0].to_dict()
            new_row = ref.copy()
            new_row["descriptor"] = "Intentos de cobro ilegal o exigencias indebidas en la zona comercial"
            new_row["cantidad_respuestas"] = traslado_total
            new_row["opciones_csv_que_contaron"] = " | ".join(src_tokens)
            df_results = pd.concat([df_results, pd.DataFrame([new_row])], ignore_index=True)

    # Eliminar filas origen
    df_results = df_results[~mask_src].copy()

    return df_results


# =========================================================
# PROCESAMIENTO
# =========================================================

def build_results_for_file(df_csv: pd.DataFrame, filename: str, guide: dict):
    """Procesa un archivo CSV completo y genera los DataFrames de resultados.

    Flujo general por cada pregunta del Excel guía:
    1. Ubica la columna correspondiente en el CSV (scoring).
    2. Construye las definiciones de grupos (aliases).
    3. Aplica refuerzos dinámicos de aliases basados en lo que
       realmente aparece en el CSV (para preguntas específicas).
    4. Cuenta respuestas según el modo (exacto o unificado).
    5. Detecta opciones no ubicadas para auditoría.
    6. Aplica la regla especial de traslado comercio 18 -> 12.

    Refuerzos dinámicos por pregunta:
    - Preguntas 13 (comunidad/comercio): oferta educativa/laboral/recreativa/cultural.
    - P14 comercio / P15 comunidad: infraestructura vial.
    - P15 comercio: inversión social.
    - P16 comunidad: espacios de riesgo.
    - P18 comercio: extorsión/cobro ilegal.
    - P19 comercio/comunidad: venta de drogas (modalidades).
    - P20 comercio: asaltos por tipo de objetivo.
    - P21 comercio / P23 comunidad: estafa/fraude.
    - P12 comunidad: prostitución.
    - P27 comunidad: delitos ambientales (SOLO envenenamiento/contaminación
      de aguas; "no_se_observan_delitos_ambientales" se filtra como opción
      no productiva y NO se refuerza).
    - Policial/Policía (cualquier pregunta): préstamos gota a gota.

    Parámetros:
        df_csv (pd.DataFrame): DataFrame del CSV ya leído y con encabezados aplanados.
        filename (str): Nombre del archivo (para inferir tipo).
        guide (dict): Guía completa cargada con load_guide_excel.

    Retorna:
        tuple: (df_results, df_mapping, df_unmapped) donde:
            - df_results: conteos por descriptor
            - df_mapping: información de mapeo pregunta -> columna CSV
            - df_unmapped: opciones del CSV no ubicadas en ningún descriptor

    Excepciones:
        ValueError: Si no se puede inferir el tipo de archivo o si la
                    hoja correspondiente no existe en el Excel guía.
    """
    file_type = infer_file_type(filename)
    if not file_type:
        raise ValueError(f"No pude identificar el tipo del archivo: {filename}")

    sheet_name = SHEET_BY_FILETYPE[file_type]
    if sheet_name not in guide:
        raise ValueError(f"No existe la hoja '{sheet_name}' en el Excel guía.")

    base = guide[sheet_name]
    df_csv = flatten_headers(df_csv.copy())

    results = []
    mapping_info = []
    unmapped_options_rows = []

    grouped_questions = build_question_groups(base)

    for (preg_num, preg_text), items in grouped_questions.items():
        question_col, score = find_question_column(df_csv, preg_num, preg_text)
        group_defs = build_group_definitions(file_type, preg_num, preg_text, items)
        mode_label = "exacto" if is_exact_question(file_type, preg_num) else "unificado"

        # Si no se encontró columna, registrar como sin mapeo
        if not question_col:
            for group_label in group_defs:
                results.append({
                    "archivo": filename,
                    "tipo": file_type,
                    "hoja_excel": sheet_name.strip(),
                    "pregunta_num": preg_num,
                    "pregunta": preg_text,
                    "descriptor": group_label,
                    "columna_pregunta_csv": "",
                    "opciones_csv_que_contaron": "",
                    "modo_conteo": mode_label,
                    "cantidad_respuestas": 0,
                })

                mapping_info.append({
                    "archivo": filename,
                    "tipo": file_type,
                    "pregunta_num": preg_num,
                    "pregunta": preg_text,
                    "descriptor": group_label,
                    "columna_pregunta_csv": "",
                    "puntaje_columna": score,
                    "mapeado": "No",
                    "motivo": "No se encontró columna de la pregunta en el CSV",
                })
            continue

        matched_aliases_union = set()
        series = df_csv[question_col]

        # ------------------------------------------------------------------
        # REFUERZOS DINÁMICOS DE ALIASES
        # ------------------------------------------------------------------
        # Para cada pregunta específica, se escanean los tokens que
        # realmente aparecen en el CSV y se agregan al grupo
        # correspondiente si coinciden con patrones esperados.
        # ------------------------------------------------------------------

        # Preguntas unificadas: agregar tokens observados que coincidan con aliases predefinidos
        if not is_exact_question(file_type, preg_num):
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            known_unified = UNIFIED_EXTRA_ALIASES.get((file_type, preg_num), set())
            if known_unified:
                observed_unified = {tok for tok in csv_tokens if tok in known_unified}
                for _, group_info in group_defs.items():
                    group_info["aliases"].update(observed_unified)

        # Comercio 20: asaltos por tipo de objetivo
        if file_type == "comercio" and preg_num == "20":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                desc_norm = normalize_token_for_compare(group_label)
                extra_matches = set()

                for tok in csv_tokens:
                    if "persona" in desc_norm and (("persona" in tok) or ("peaton" in tok)):
                        extra_matches.add(tok)
                    elif "comercio" in desc_norm and "comerc" in tok:
                        extra_matches.add(tok)
                    elif "vivienda" in desc_norm and (("vivienda" in tok) or ("casa" in tok)):
                        extra_matches.add(tok)
                    elif "transporte" in desc_norm and (("transporte" in tok) or ("bus" in tok) or ("autobus" in tok)):
                        extra_matches.add(tok)

                group_info["aliases"].update(extra_matches)

        # Comunidad 12: prostitución
        if file_type == "comunidad" and preg_num == "12":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                desc_norm = normalize_token_for_compare(group_label)
                if "prostit" in desc_norm:
                    extra_prostitucion = {tok for tok in csv_tokens if "prostit" in tok}
                    group_info["aliases"].update(extra_prostitucion)

        # Comercio 18: extorsión
        if file_type == "comercio" and preg_num == "18":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                desc_norm = normalize_token_for_compare(group_label)

                if (
                    "extors" in desc_norm
                    or "extorc" in desc_norm
                    or "exigir_cobro" in desc_norm
                ):
                    extra_extorsion = {
                        tok for tok in csv_tokens
                        if (
                            "extors" in tok
                            or "extorc" in tok
                            or "cobro_ilegal" in tok
                            or "exigencias_indebidas" in tok
                        )
                    }
                    group_info["aliases"].update(extra_extorsion)

        # Comercio 19: venta de drogas (modalidades)
        if file_type == "comercio" and preg_num == "19":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for _, group_info in group_defs.items():
                extra_drogas = {
                    tok for tok in csv_tokens
                    if (
                        tok == "en_via_publica"
                        or tok == "en_espacios_cerrados_casas_edificaciones_u_otros_inmuebles"
                        or tok == "de_forma_ocasional_o_movil_modalidad_expres_sin_punto_fijo"
                        or tok == "venta_de_drogas"
                        or "forma_ocasional" in tok
                        or "espacios_cerrados" in tok
                        or "via_publica" in tok
                    )
                }
                group_info["aliases"].update(extra_drogas)

        # Pregunta 13 (comunidad y comercio): oferta de servicios
        if preg_num == "13" and file_type in {"comunidad", "comercio"}:
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for _, group_info in group_defs.items():
                extra_oferta = {
                    tok for tok in csv_tokens
                    if (
                        tok == "falta_de_oferta_educativa"
                        or tok == "falta_de_oferta_laboral"
                        or tok == "falta_de_oferta_recreativa"
                        or tok == "falta_de_actividades_culturales"
                    )
                }
                group_info["aliases"].update(extra_oferta)

        # P14 comercio / P15 comunidad: infraestructura vial
        if (
            (file_type == "comercio" and preg_num == "14")
            or (file_type == "comunidad" and preg_num == "15")
        ):
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for _, group_info in group_defs.items():
                extra_infra = {
                    tok for tok in csv_tokens
                    if (
                        tok == "calles_en_mal_estado"
                        or tok == "falta_de_iluminacion"
                        or tok == "falta_de_senalizacion"
                        or tok == "falta_o_deterioro_de_aceras"
                    )
                }
                group_info["aliases"].update(extra_infra)

        # P15 comercio: inversión social
        if file_type == "comercio" and preg_num == "15":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for _, group_info in group_defs.items():
                extra_social = {
                    tok for tok in csv_tokens
                    if (
                        tok == "falta_de_programas_sociales"
                        or tok == "falta_de_espacios_de_integracion_social"
                        or tok == "falta_de_apoyo_institucional"
                        or tok == "falta_de_inversion_social"
                    )
                }
                group_info["aliases"].update(extra_social)

        # Policial/Policía: préstamos gota a gota (cualquier pregunta)
        if file_type in {"policial", "policia"}:
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                desc_norm = normalize_token_for_compare(group_label)
                if (
                    "gota_a_gota" in desc_norm
                    or "prestamo_gota_a_gota" in desc_norm
                    or "prestamos_gota_a_gota" in desc_norm
                ):
                    extra_gota = {
                        tok for tok in csv_tokens
                        if (
                            "gota_a_gota" in tok
                            or "prestamo_gota_a_gota" in tok
                            or "prestamos_gota_a_gota" in tok
                            or "cobro_gota_a_gota" in tok
                            or "cobros_gota_a_gota" in tok
                        )
                    }
                    group_info["aliases"].update(extra_gota)

        # Comunidad 27: delitos ambientales.
        # SOLO se refuerzan las variantes de envenenamiento/contaminación de aguas.
        # "no_se_observan_delitos_ambientales" NO se refuerza aquí porque
        # ya fue filtrado como opción no productiva por is_unproductive_option
        # (contiene "no_se_observan" que está en PATRONES_NO_OBSERVA).
        if file_type == "comunidad" and preg_num == "27":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                desc_norm = normalize_token_for_compare(group_label)

                if "envenenamiento" in desc_norm or "contaminacion_de_aguas" in desc_norm:
                    extra_aguas = {
                        tok for tok in csv_tokens
                        if tok in {
                            "envenenamiento_de_aguas",
                            "envenenamiento_o_contaminacion_de_aguas",
                            "contaminacion_de_aguas",
                            "contaminacion_o_envenenamiento_de_aguas",
                        }
                    }
                    group_info["aliases"].update(extra_aguas)

        # P21 comercio / P23 comunidad: estafa/fraude
        if (
            (file_type == "comunidad" and preg_num == "23")
            or (file_type == "comercio" and preg_num == "21")
        ):
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                if normalize_token_for_compare(group_label) == "estafa":
                    extra_estafa = {
                        tok for tok in csv_tokens
                        if ("estafa" in tok or "fraude" in tok)
                    }
                    group_info["aliases"].update(extra_estafa)

        # ------------------------------------------------------------------
        # CONTEO POR GRUPO
        # ------------------------------------------------------------------
        for group_label, group_info in group_defs.items():
            aliases = group_info["aliases"]
            matched_aliases_union.update(aliases)

            # Seleccionar modo de conteo
            if is_exact_question(file_type, preg_num):
                if group_info.get("group_mode") == "merged":
                    total, matched_tokens = count_group_unified(
                        series, aliases, file_type=file_type, question_num=str(preg_num)
                    )
                else:
                    total, matched_tokens = count_group_exact(
                        series, aliases, file_type=file_type, question_num=str(preg_num)
                    )
            else:
                total, matched_tokens = count_group_unified(
                    series, aliases, file_type=file_type, question_num=str(preg_num)
                )

            results.append({
                "archivo": filename,
                "tipo": file_type,
                "hoja_excel": sheet_name.strip(),
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "descriptor": group_label,
                "columna_pregunta_csv": question_col,
                "opciones_csv_que_contaron": " | ".join(sorted(matched_tokens.keys())),
                "modo_conteo": mode_label,
                "cantidad_respuestas": int(total),
            })

            mapping_info.append({
                "archivo": filename,
                "tipo": file_type,
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "descriptor": group_label,
                "columna_pregunta_csv": question_col,
                "puntaje_columna": score,
                "mapeado": "Sí",
                "motivo": f"Conteo {mode_label} según regla de la pregunta",
            })

        # ------------------------------------------------------------------
        # DETECCIÓN DE OPCIONES NO UBICADAS
        # ------------------------------------------------------------------
        unmapped_counter, blank_rows = find_unmapped_tokens(
            series,
            matched_aliases_union,
            file_type=file_type,
            question_num=str(preg_num),
        )

        for token, cnt in unmapped_counter.items():
            unmapped_options_rows.append({
                "archivo": filename,
                "tipo": file_type,
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "columna_pregunta_csv": question_col,
                "opcion_csv_no_ubicada": token,
                "cantidad": int(cnt),
            })

        unmapped_options_rows.append({
            "archivo": filename,
            "tipo": file_type,
            "pregunta_num": preg_num,
            "pregunta": preg_text,
            "columna_pregunta_csv": question_col,
            "opcion_csv_no_ubicada": "[filas_vacias_ignoradas]",
            "cantidad": int(blank_rows),
        })

    # Construir DataFrames finales
    df_results = pd.DataFrame(results)
    df_mapping = pd.DataFrame(mapping_info)
    df_unmapped = pd.DataFrame(unmapped_options_rows)

    # Aplicar regla especial de traslado comercio 18 -> 12
    df_results = apply_special_cross_question_rules(df_results)

    return df_results, df_mapping, df_unmapped
    # ==============================================================================
# PARTE 9: RESÚMENES, FILTRADO DE CEROS Y FUNCIONES DE EXPORTACIÓN A EXCEL
# ==============================================================================
# Esta sección contiene las funciones que transforman los resultados brutos
# en resúmenes agregados (por pregunta, totales globales por descriptor),
# eliminan filas con conteo cero para limpiar la visualización, generan las
# tablas de ranking para la interfaz y gestionan la exportación de múltiples
# DataFrames a un solo archivo Excel con varias hojas.
# ==============================================================================

# =========================================================
# RESÚMENES
# =========================================================

def render_totals_tables(df: pd.DataFrame):
    if df.empty:
        st.info("No hay resultados con conteos mayores a 0 para los filtros seleccionados.")
        return

    grouped = (
        df.sort_values(
            by=["tipo", "pregunta_num", "cantidad_respuestas", "descriptor"],
            ascending=[True, True, False, True]
        )
        .groupby(["tipo", "pregunta_num", "pregunta", "modo_conteo"], sort=False)
    )

    for (tipo, pregunta_num, pregunta, modo_conteo), subdf in grouped:
        st.markdown(f"## Pregunta {pregunta_num}")
        st.markdown(f"**Tipo:** {tipo}")
        st.markdown(f"**Modo de conteo:** {modo_conteo}")
        st.markdown(f"**Pregunta:** {pregunta}")

        show_df = subdf[["descriptor", "cantidad_respuestas"]].copy()
        show_df = show_df.sort_values(
            by=["cantidad_respuestas", "descriptor"],
            ascending=[False, True]
        ).reset_index(drop=True)

        show_df.insert(0, "Ranking", range(1, len(show_df) + 1))
        show_df = show_df.rename(columns={
            "descriptor": "Descriptor",
            "cantidad_respuestas": "Cantidad"
        })

        # 🔥 AQUI ESTA EL CAMBIO VISUAL (FILAS ALTERNAS)
        def color_rows(row):
            if row.name % 2 == 0:
                return ["background-color: #1e1e1e"] * len(row)  # gris oscuro
            else:
                return ["background-color: #2a2a2a"] * len(row)  # gris claro

        styled_df = show_df.style.apply(color_rows, axis=1)

        st.dataframe(styled_df, use_container_width=True)
        st.divider()
        # ==============================================================================
# PARTE 10: INTERFAZ DE USUARIO PRINCIPAL CON STREAMLIT (SIDEBAR, CARGA DE
#           ARCHIVOS, FILTROS, PESTAÑAS Y DESCARGA)
# ==============================================================================
# Esta sección construye toda la interfaz visual de la aplicación:
# - Título y subtítulo descriptivo.
# - Sidebar con configuración del Excel guía y estado de carga.
# - Área principal con dos columnas: estado del Excel y carga de CSV.
# - Procesamiento de todos los archivos subidos con manejo de errores.
# - Métricas resumen (archivos procesados, preguntas detectadas, etc.).
# - Filtros multinivel por tipo, pregunta y archivo.
# - Cinco pestañas de resultados: totales por descriptor, resumen por
#   pregunta, detalle, mapeo pregunta↔CSV, opciones no ubicadas.
# - Botón de descarga de resultados en formato Excel con múltiples hojas.
# ==============================================================================

# =========================================================
# INTERFAZ
# =========================================================

st.title("Conteo de respuestas por pregunta / descriptor")
st.caption("Usa el Excel como guía, ubica la columna de cada pregunta en el CSV y cuenta las opciones reales dentro de cada celda.")

# --------------------------------------------------------------------------
# Sidebar: configuración del Excel guía
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("Configuración")
    excel_path = st.text_input("Nombre del Excel guía", value=EXCEL_GUIA)
    st.info(
        "El Excel debe estar en la raíz del repositorio.\n\n"
        "Hojas esperadas:\n"
        "- Comunidad\n"
        "- Comercio\n"
        "- Policia"
    )

# --------------------------------------------------------------------------
# Carga del Excel guía
# --------------------------------------------------------------------------
guide = None
guide_error = None

try:
    if Path(excel_path).exists():
        guide = load_guide_excel(excel_path)
    else:
        guide_error = f"No se encontró el archivo Excel guía: {excel_path}"
except Exception as e:
    guide_error = f"Error leyendo el Excel guía: {e}"

# --------------------------------------------------------------------------
# Columnas de estado: Excel guía y carga de CSV
# --------------------------------------------------------------------------
col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Estado del Excel guía")
    if guide_error:
        st.error(guide_error)
    else:
        st.success("Excel guía cargado correctamente.")
        st.dataframe(build_guide_summary(guide), use_container_width=True)

with col2:
    st.subheader("Subir CSV")
    uploaded_files = st.file_uploader(
        "Puedes subir uno o varios CSV",
        type=["csv"],
        accept_multiple_files=True
    )

# --------------------------------------------------------------------------
# Validaciones iniciales
# --------------------------------------------------------------------------
if guide is None:
    st.warning("Primero debe estar disponible el Excel guía en el repositorio.")
    st.stop()

if not uploaded_files:
    st.info("Sube al menos un CSV para procesar.")
    st.stop()

# --------------------------------------------------------------------------
# Procesamiento de archivos CSV
# --------------------------------------------------------------------------
all_results = []
all_mapping = []
all_unmapped = []
read_errors = []

for file in uploaded_files:
    try:
        content = file.read()
        df_csv = try_read_csv_bytes(content)
        df_csv = flatten_headers(df_csv)

        df_results, df_mapping, df_unmapped = build_results_for_file(df_csv, file.name, guide)

        all_results.append(df_results)
        all_mapping.append(df_mapping)
        all_unmapped.append(df_unmapped)

    except Exception as e:
        read_errors.append({
            "archivo": file.name,
            "error": str(e)
        })

# --------------------------------------------------------------------------
# Mostrar errores de lectura
# --------------------------------------------------------------------------
if read_errors:
    st.subheader("Errores detectados")
    st.dataframe(pd.DataFrame(read_errors), use_container_width=True)

if not all_results:
    st.error("No se pudo procesar ningún archivo.")
    st.stop()

# --------------------------------------------------------------------------
# Consolidar resultados de todos los archivos
# --------------------------------------------------------------------------
df_results_all = pd.concat(all_results, ignore_index=True)
df_mapping_all = pd.concat(all_mapping, ignore_index=True)
df_unmapped_all = pd.concat(all_unmapped, ignore_index=True) if all_unmapped else pd.DataFrame()

df_summary = summarize_results(df_results_all)
df_totals = build_global_totals(df_results_all)

# Eliminar filas con conteo cero
df_results_all = remove_zero_rows(df_results_all, "cantidad_respuestas")
df_summary = remove_zero_rows(df_summary, "cantidad_respuestas")
df_totals = remove_zero_rows(df_totals, "cantidad_respuestas")

# Limpiar opciones no ubicadas: eliminar ceros y la fila de filas vacías
if not df_unmapped_all.empty:
    df_unmapped_all = remove_zero_rows(df_unmapped_all, "cantidad")
    df_unmapped_all = df_unmapped_all[df_unmapped_all["opcion_csv_no_ubicada"] != "[filas_vacias_ignoradas]"].copy()

# --------------------------------------------------------------------------
# Filtros multinivel
# --------------------------------------------------------------------------
st.markdown("## Filtros")

colf1, colf2, colf3 = st.columns(3)

tipos = sorted(df_results_all["tipo"].dropna().unique().tolist()) if not df_results_all.empty else []
preguntas = sorted(df_results_all["pregunta_num"].dropna().unique().tolist(), key=question_sort_key) if not df_results_all.empty else []
archivos = sorted(df_results_all["archivo"].dropna().unique().tolist()) if not df_results_all.empty else []

with colf1:
    filtro_tipo = st.multiselect("Tipo", options=tipos, default=tipos)

with colf2:
    filtro_pregunta = st.multiselect("Pregunta", options=preguntas, default=preguntas)

with colf3:
    filtro_archivo = st.multiselect("Archivo", options=archivos, default=archivos)

# Aplicar filtros a todos los DataFrames
df_summary_f = df_summary[
    df_summary["tipo"].isin(filtro_tipo) &
    df_summary["pregunta_num"].isin(filtro_pregunta) &
    df_summary["archivo"].isin(filtro_archivo)
].copy()

df_results_f = df_results_all[
    df_results_all["tipo"].isin(filtro_tipo) &
    df_results_all["pregunta_num"].isin(filtro_pregunta) &
    df_results_all["archivo"].isin(filtro_archivo)
].copy()

df_mapping_f = df_mapping_all[
    df_mapping_all["tipo"].isin(filtro_tipo) &
    df_mapping_all["pregunta_num"].isin(filtro_pregunta) &
    df_mapping_all["archivo"].isin(filtro_archivo)
].copy()

if not df_unmapped_all.empty:
    df_unmapped_f = df_unmapped_all[
        df_unmapped_all["tipo"].isin(filtro_tipo) &
        df_unmapped_all["pregunta_num"].isin(filtro_pregunta) &
        df_unmapped_all["archivo"].isin(filtro_archivo)
    ].copy()
else:
    df_unmapped_f = pd.DataFrame()

df_totals_f = df_totals[
    df_totals["tipo"].isin(filtro_tipo) &
    df_totals["pregunta_num"].isin(filtro_pregunta)
].copy()

# --------------------------------------------------------------------------
# Métricas resumen
# --------------------------------------------------------------------------
st.markdown("## Resumen general")

m1, m2, m3, m4 = st.columns(4)

m1.metric("Archivos procesados", len(df_results_all["archivo"].unique()) if not df_results_all.empty else 0)
m2.metric("Preguntas detectadas", len(df_results_all["pregunta_num"].unique()) if not df_results_all.empty else 0)
m3.metric("Resultados mostrados", len(df_results_all) if not df_results_all.empty else 0)
m4.metric("Respuestas contabilizadas", int(df_results_all["cantidad_respuestas"].sum()) if not df_results_all.empty else 0)

# --------------------------------------------------------------------------
# Pestañas de resultados
# --------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Totales por descriptor",
    "Resumen por pregunta",
    "Detalle",
    "Mapeo pregunta ↔ CSV",
    "Opciones no ubicadas"
])

with tab1:
    st.subheader("Totales por descriptor")
    render_totals_tables(df_totals_f)

with tab2:
    st.subheader("Resumen por pregunta")
    if df_summary_f.empty:
        st.info("No hay datos para mostrar.")
    else:
        st.dataframe(df_summary_f, use_container_width=True)

with tab3:
    st.subheader("Detalle")
    if df_results_f.empty:
        st.info("No hay datos para mostrar.")
    else:
        st.dataframe(df_results_f, use_container_width=True)

with tab4:
    st.subheader("Mapeo pregunta ↔ CSV")
    if df_mapping_f.empty:
        st.info("No hay datos para mostrar.")
    else:
        st.dataframe(df_mapping_f, use_container_width=True)

with tab5:
    st.subheader("Opciones no ubicadas")
    if df_unmapped_f.empty:
        st.info("No hay opciones no ubicadas.")
    else:
        st.dataframe(df_unmapped_f, use_container_width=True)

# --------------------------------------------------------------------------
# Descarga de resultados en Excel
# --------------------------------------------------------------------------
dfs_export = {
    "totales_descriptor": df_totals_f,
    "resumen_pregunta": df_summary_f,
    "detalle": df_results_f,
    "mapeo": df_mapping_f,
}

if not df_unmapped_f.empty:
    dfs_export["no_ubicadas"] = df_unmapped_f

excel_bytes = to_excel_bytes(dfs_export)

st.download_button(
    label="Descargar resultados en Excel",
    data=excel_bytes,
    file_name="conteo_respuestas_preguntas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
    
    
        
    
