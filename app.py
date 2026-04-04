# app.py
# -*- coding: utf-8 -*-

import io
import re
import csv
import unicodedata
from pathlib import Path
from collections import defaultdict, Counter

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Conteo de respuestas por pregunta / descriptor",
    layout="wide"
)

# =========================================================
# CONFIGURACIÓN
# =========================================================
EXCEL_GUIA = "Guía de Preguntas para paretos 2026.xlsx"

SHEET_BY_FILETYPE = {
    "comunidad": "Comunidad ",
    "comercio": "Comercio",
    "policia": "Policia",
    "policial": "Policia",
}

TOKENS_IGNORAR = {"", "nan", "none", "null"}

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

PATRONES_NO_OBSERVA = [
    "no_se_observa",
    "no_se_observan",
    "no_se_presenta",
    "no_se_presentan",
    "no_hay",
    "ninguno",
    "ninguna",
]

# Preguntas que se mantienen desglosadas
PREGUNTAS_EXACTAS = {
    "comunidad": {"12", "18", "20", "22", "24", "26", "27"},
    "comercio": {"12", "18", "20", "22"},
    "policia": None,
    "policial": None,
}

# Etiquetas unificadas por pregunta
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

# Refuerzos directos para preguntas unificadas
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

# =========================================================
# UTILIDADES DE TEXTO
# =========================================================
def strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFD", str(text))
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def norm(text) -> str:
    if text is None:
        return ""
    s = str(text).strip().strip("\ufeff")
    s = s.replace("\n", " ").replace("\r", " ")
    s = strip_accents(s).lower()
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def slugify(text) -> str:
    if text is None:
        return ""
    s = norm(text)
    s = s.replace("/", " ")
    s = re.sub(r"[^\w\s]", "", s, flags=re.UNICODE)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_")


def normalize_option_token(text) -> str:
    return slugify(text)


def extract_question_number(text: str) -> str:
    s = str(text).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[\.\)]?", s)
    return m.group(1) if m else ""


def question_sort_key(q):
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
    if value is None:
        return True
    return norm(value) in TOKENS_IGNORAR


def clean_descriptor_display(text: str) -> str:
    s = str(text).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def normalize_common_typos(token: str) -> str:
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
    t = normalize_option_token(token)
    t = t.strip("._- ")
    t = normalize_common_typos(t)
    t = re.sub(r"_+", "_", t).strip("_")
    return t


def normalize_display_for_grouping(text: str) -> str:
    s = clean_descriptor_display(text)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"[.;,:]+$", "", s).strip()
    return s


# =========================================================
# TIPO DE ARCHIVO
# =========================================================
def infer_file_type(filename: str) -> str:
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


# =========================================================
# EXCEL GUÍA
# =========================================================
def load_guide_excel(path_excel: str):
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
                current_question_num = qnum
                current_question_text = first

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

            if current_question_num and current_question_text:
                rows.append({
                    "pregunta_num": current_question_num,
                    "pregunta_texto": current_question_text,
                    "pregunta_slug": normalize_option_token(current_question_text),
                    "descriptor_texto": clean_descriptor_display(first),
                    "descriptor_slug": normalize_option_token(first),
                })

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


# =========================================================
# CSV
# =========================================================
def parse_csv_with_python_engine(content: bytes, encoding: str, delimiter: str):
    text = content.decode(encoding, errors="replace")
    rows = []

    reader = csv.reader(io.StringIO(text), delimiter=delimiter, quotechar='"')

    max_cols = 0
    for row in reader:
        rows.append(row)
        max_cols = max(max_cols, len(row))

    if not rows or max_cols <= 1:
        return pd.DataFrame()

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
    attempts = [
        ("utf-8-sig", ","),
        ("utf-8", ","),
        ("latin-1", ","),
        ("utf-8-sig", ";"),
        ("utf-8", ";"),
        ("latin-1", ";"),
    ]

    last_error = None

    for encoding, delimiter in attempts:
        try:
            df = parse_csv_with_python_engine(content, encoding, delimiter)
            if not df.empty and df.shape[1] > 1:
                return df
        except Exception as e:
            last_error = e

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
    new_cols = []
    for c in df.columns:
        s = str(c).replace("\n", " ").replace("\r", " ").strip()
        s = re.sub(r"\s+", " ", s)
        new_cols.append(s)
    df.columns = new_cols
    return df


# =========================================================
# UBICAR PREGUNTA
# =========================================================
def build_question_groups(guide_sheet_rows: list):
    grouped = defaultdict(list)
    for r in guide_sheet_rows:
        grouped[(r["pregunta_num"], r["pregunta_texto"])].append(r)
    return grouped


def score_question_column(col_name: str, question_num: str, question_text: str) -> int:
    col_norm = norm(col_name)
    col_slug = normalize_option_token(col_name)

    score = 0
    qnum = str(question_num).strip()
    qtext_norm = norm(question_text)
    qtext_slug = normalize_option_token(question_text)

    text_wo_num = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\.\)]?\s*", "", question_text).strip()
    text_wo_num_norm = norm(text_wo_num)
    text_wo_num_slug = normalize_option_token(text_wo_num)

    if qnum:
        if col_norm.startswith(qnum):
            score += 100
        if f"{qnum}." in col_norm or f"{qnum})" in col_norm:
            score += 80
        if re.search(rf"(^|\s){re.escape(qnum)}(\.|\)|\s|$)", col_norm):
            score += 60

    if qtext_norm and qtext_norm == col_norm:
        score += 120
    if qtext_slug and qtext_slug == col_slug:
        score += 120

    if text_wo_num_norm and text_wo_num_norm in col_norm:
        score += 90
    if text_wo_num_slug and text_wo_num_slug in col_slug:
        score += 90

    if qtext_norm and qtext_norm in col_norm:
        score += 50
    if qtext_slug and qtext_slug in col_slug:
        score += 50

    return score


def find_question_column(df: pd.DataFrame, question_num: str, question_text: str):
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
    if is_effectively_empty(value):
        return []

    raw = str(value).strip()
    if not raw:
        return []

    parts = [p.strip() for p in raw.split(",")]
    parts = [p for p in parts if p and norm(p) not in TOKENS_IGNORAR]
    return parts


def is_q27_no_observa_ambiental(token_norm: str, file_type: str = "", question_num: str = "") -> bool:
    return (
        file_type == "comunidad"
        and str(question_num) == "27"
        and token_norm == "no_se_observan_delitos_ambientales"
    )


def is_no_observa_option(token_norm: str) -> bool:
    for p in PATRONES_NO_OBSERVA:
        if p in token_norm:
            return True
    return False


def is_unproductive_option(token_norm: str, file_type: str = "", question_num: str = "") -> bool:
    if is_q27_no_observa_ambiental(token_norm, file_type, question_num):
        return False

    if token_norm in OPCIONES_NO_PRODUCTIVAS:
        return True
    if token_norm.startswith("otro_") or token_norm.startswith("otros_"):
        return True
    if is_no_observa_option(token_norm):
        return True
    return False


def tokenize_cell_unique(value: str, file_type: str = "", question_num: str = ""):
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


# =========================================================
# REGLAS DE CONTEO
# =========================================================
def is_exact_question(file_type: str, question_num: str) -> bool:
    exacts = PREGUNTAS_EXACTAS.get(file_type)
    if exacts is None:
        return True
    return question_num in exacts


def get_unified_label(file_type: str, question_num: str, question_text: str) -> str:
    if (file_type, question_num) in UNIFIED_LABELS:
        return UNIFIED_LABELS[(file_type, question_num)]

    text_wo_num = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\.\)]?\s*", "", question_text).strip()
    return clean_descriptor_display(text_wo_num) if text_wo_num else f"Pregunta {question_num}"


def build_descriptor_aliases(file_type: str, question_num: str, descriptor_text: str):
    base = normalize_option_token(descriptor_text)
    base_norm = normalize_token_for_compare(base)
    aliases = {base_norm}

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

        # Comunidad 27: delitos ambientales
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
        "no_se_observan_delitos_ambientales": {
            "no_se_observan_delitos_ambientales",
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

    if (file_type, question_num) in UNIFIED_EXTRA_ALIASES:
        aliases.update(UNIFIED_EXTRA_ALIASES[(file_type, question_num)])

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

    aliases = {normalize_token_for_compare(a) for a in aliases if a}
    aliases = {
        a for a in aliases
        if not is_unproductive_option(a, file_type=file_type, question_num=str(question_num))
    }
    return aliases


def get_exact_canonical_group(file_type: str, question_num: str, descriptor_text: str):
    base = normalize_token_for_compare(descriptor_text)
    label = normalize_display_for_grouping(descriptor_text)
    group_mode = "exact"

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

    if file_type == "comercio" and question_num == "18":
        if (
            "extors" in base
            or "extorc" in base
            or "exigencias_indebidas" in base
            or "cobro_ilegal" in base
        ):
            return "Extorsión (amenazas o intimidación para exigir cobro de dinero u otros beneficios de manera ilegal a comercios)", "merged"

    if file_type == "comercio" and question_num == "20":
        if ("persona" in base) or ("peaton" in base):
            return "Asalto a personas", "merged"
        if "comerc" in base:
            return "Asalto a comercio", "merged"
        if ("vivienda" in base) or ("casa" in base):
            return "Asalto a vivienda", "merged"
        if ("transporte" in base) or ("bus" in base) or ("autobus" in base):
            return "Asalto a transporte público", "merged"

    # Comunidad 27: unificar ambiental aguas + no observan
    if file_type == "comunidad" and question_num == "27":
        if base in {
            "envenenamiento_de_aguas",
            "envenenamiento_o_contaminacion_de_aguas",
            "contaminacion_de_aguas",
            "contaminacion_o_envenenamiento_de_aguas",
        }:
            return "Envenenamiento o contaminación de aguas", "merged"
        if base == "no_se_observan_delitos_ambientales":
            return "No se observan delitos ambientales", "merged"

    # Policial / Policía: gota a gota + evitar duplicados visuales
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


# =========================================================
# REGLA ESPECIAL: COMERCIO 18 -> COMERCIO 12
# =========================================================
def is_comercio_q18_extorsion_descriptor(desc: str) -> bool:
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

    src_tokens = []
    for txt in src_rows["opciones_csv_que_contaron"].fillna("").astype(str).tolist():
        if txt.strip():
            src_tokens.extend([p.strip() for p in txt.split("|") if p.strip()])
    src_tokens = sorted(set(src_tokens))

    if mask_dst.any():
        dst_idx = df_results[mask_dst].index[0]
        df_results.at[dst_idx, "cantidad_respuestas"] = int(df_results.at[dst_idx, "cantidad_respuestas"]) + traslado_total

        current_tokens = str(df_results.at[dst_idx, "opciones_csv_que_contaron"] or "").strip()
        dst_tokens = [p.strip() for p in current_tokens.split("|") if p.strip()] if current_tokens else []
        merged_tokens = sorted(set(dst_tokens + src_tokens))
        df_results.at[dst_idx, "opciones_csv_que_contaron"] = " | ".join(merged_tokens)
    else:
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

    df_results = df_results[~mask_src].copy()

    return df_results


# =========================================================
# PROCESAMIENTO
# =========================================================
def build_results_for_file(df_csv: pd.DataFrame, filename: str, guide: dict):
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

        if not is_exact_question(file_type, preg_num):
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            known_unified = UNIFIED_EXTRA_ALIASES.get((file_type, preg_num), set())
            if known_unified:
                observed_unified = {tok for tok in csv_tokens if tok in known_unified}
                for _, group_info in group_defs.items():
                    group_info["aliases"].update(observed_unified)

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

        if file_type == "comunidad" and preg_num == "12":
            csv_tokens = set()
            for val in series:
                csv_tokens.update(tokenize_cell_unique(val, file_type=file_type, question_num=str(preg_num)))

            for group_label, group_info in group_defs.items():
                desc_norm = normalize_token_for_compare(group_label)
                if "prostit" in desc_norm:
                    extra_prostitucion = {tok for tok in csv_tokens if "prostit" in tok}
                    group_info["aliases"].update(extra_prostitucion)

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

        # Policial: refuerzo dinámico para préstamos gota a gota
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

        # Comunidad 27: refuerzo dinámico ambiental
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

                if desc_norm == "no_se_observan_delitos_ambientales":
                    extra_no = {
                        tok for tok in csv_tokens
                        if tok == "no_se_observan_delitos_ambientales"
                    }
                    group_info["aliases"].update(extra_no)

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

        for group_label, group_info in group_defs.items():
            aliases = group_info["aliases"]
            matched_aliases_union.update(aliases)

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

    df_results = pd.DataFrame(results)
    df_mapping = pd.DataFrame(mapping_info)
    df_unmapped = pd.DataFrame(unmapped_options_rows)

    df_results = apply_special_cross_question_rules(df_results)

    return df_results, df_mapping, df_unmapped


# =========================================================
# RESÚMENES
# =========================================================
def summarize_results(df_results: pd.DataFrame):
    if df_results.empty:
        return pd.DataFrame()

    summary = (
        df_results
        .groupby(
            ["archivo", "tipo", "pregunta_num", "pregunta", "modo_conteo"],
            as_index=False
        )["cantidad_respuestas"]
        .sum()
    )

    summary["sort_key"] = summary["pregunta_num"].apply(question_sort_key)
    summary = summary.sort_values(
        by=["archivo", "sort_key", "pregunta"],
        kind="stable"
    ).drop(columns=["sort_key"])

    return summary


def build_global_totals(df_results_all: pd.DataFrame) -> pd.DataFrame:
    if df_results_all.empty:
        return pd.DataFrame()

    totals = (
        df_results_all
        .groupby(
            ["tipo", "pregunta_num", "pregunta", "descriptor", "modo_conteo"],
            as_index=False
        )["cantidad_respuestas"]
        .sum()
    )

    totals["sort_key"] = totals["pregunta_num"].apply(question_sort_key)
    totals = totals.sort_values(
        by=["tipo", "sort_key", "cantidad_respuestas", "descriptor"],
        ascending=[True, True, False, True],
        kind="stable"
    ).drop(columns=["sort_key"])

    return totals


def remove_zero_rows(df: pd.DataFrame, count_col: str):
    if df.empty or count_col not in df.columns:
        return df.copy()
    return df[df[count_col] > 0].copy()


# =========================================================
# EXPORTAR
# =========================================================
def to_excel_bytes(dfs: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


# =========================================================
# VISTA
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

        st.table(show_df)
        st.divider()


# =========================================================
# INTERFAZ
# =========================================================
st.title("Conteo de respuestas por pregunta / descriptor")
st.caption("Usa el Excel como guía, ubica la columna de cada pregunta en el CSV y cuenta las opciones reales dentro de cada celda.")

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

guide = None
guide_error = None

try:
    if Path(excel_path).exists():
        guide = load_guide_excel(excel_path)
    else:
        guide_error = f"No se encontró el archivo Excel guía: {excel_path}"
except Exception as e:
    guide_error = f"Error leyendo el Excel guía: {e}"

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

if guide is None:
    st.warning("Primero debe estar disponible el Excel guía en el repositorio.")
    st.stop()

if not uploaded_files:
    st.info("Sube al menos un CSV para procesar.")
    st.stop()

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

if read_errors:
    st.subheader("Errores detectados")
    st.dataframe(pd.DataFrame(read_errors), use_container_width=True)

if not all_results:
    st.error("No se pudo procesar ningún archivo.")
    st.stop()

df_results_all = pd.concat(all_results, ignore_index=True)
df_mapping_all = pd.concat(all_mapping, ignore_index=True)
df_unmapped_all = pd.concat(all_unmapped, ignore_index=True) if all_unmapped else pd.DataFrame()

df_summary = summarize_results(df_results_all)
df_totals = build_global_totals(df_results_all)

df_results_all = remove_zero_rows(df_results_all, "cantidad_respuestas")
df_summary = remove_zero_rows(df_summary, "cantidad_respuestas")
df_totals = remove_zero_rows(df_totals, "cantidad_respuestas")

if not df_unmapped_all.empty:
    df_unmapped_all = remove_zero_rows(df_unmapped_all, "cantidad")
    df_unmapped_all = df_unmapped_all[df_unmapped_all["opcion_csv_no_ubicada"] != "[filas_vacias_ignoradas]"].copy()

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

st.markdown("## Resumen general")

m1, m2, m3, m4 = st.columns(4)

m1.metric("Archivos procesados", len(df_results_all["archivo"].unique()) if not df_results_all.empty else 0)
m2.metric("Preguntas detectadas", len(df_results_all["pregunta_num"].unique()) if not df_results_all.empty else 0)
m3.metric("Resultados mostrados", len(df_results_all) if not df_results_all.empty else 0)
m4.metric("Respuestas contabilizadas", int(df_results_all["cantidad_respuestas"].sum()) if not df_results_all.empty else 0)

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
