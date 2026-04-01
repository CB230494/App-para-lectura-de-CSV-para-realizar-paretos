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
    "comercio": {"12", "18", "20"},
    "policia": None,
    "policial": None,
}

# Preguntas con regla especial de estafas
PREGUNTAS_ESTAFA_ESPECIAL = {
    ("comunidad", "23"),
    ("comercio", "21"),
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
    ("comercio", "15"): "Infraestructura vial",
    ("comercio", "16"): "Espacios de riesgo",
    ("comercio", "19"): "Venta de drogas",
    ("comercio", "21"): "Estafa",
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


def normalize_token_for_compare(token: str) -> str:
    t = normalize_option_token(token)
    t = t.strip("._- ")
    return t


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
                r["descriptor_slug"],
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
# RESPUESTAS DENTRO DE CELDAS
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


def is_no_observa_option(token_norm: str) -> bool:
    for p in PATRONES_NO_OBSERVA:
        if p in token_norm:
            return True
    return False


def is_unproductive_option(token_norm: str) -> bool:
    if token_norm in OPCIONES_NO_PRODUCTIVAS:
        return True
    if token_norm.startswith("otro_") or token_norm.startswith("otros_"):
        return True
    if is_no_observa_option(token_norm):
        return True
    return False


def tokenize_cell_unique(value: str):
    options = split_multiselect_cell(value)
    tokens = []

    for opt in options:
        token = normalize_token_for_compare(opt)
        if not token or token in TOKENS_IGNORAR:
            continue
        if is_unproductive_option(token):
            continue
        tokens.append(token)

    return sorted(set(tokens))


# =========================================================
# REGLAS POR PREGUNTA
# =========================================================
def is_exact_question(file_type: str, question_num: str) -> bool:
    exacts = PREGUNTAS_EXACTAS.get(file_type)
    if exacts is None:
        return True
    return question_num in exacts


def is_estafa_special(file_type: str, question_num: str) -> bool:
    return (file_type, question_num) in PREGUNTAS_ESTAFA_ESPECIAL


def get_unified_label(file_type: str, question_num: str, question_text: str) -> str:
    if (file_type, question_num) in UNIFIED_LABELS:
        return UNIFIED_LABELS[(file_type, question_num)]

    text_wo_num = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\.\)]?\s*", "", question_text).strip()
    return clean_descriptor_display(text_wo_num) if text_wo_num else f"Pregunta {question_num}"


def build_descriptor_aliases(file_type: str, question_num: str, descriptor_text: str):
    base = normalize_option_token(descriptor_text)
    aliases = {base}

    alias_map = {
        "consumo_de_drogas": {
            "consumo_de_drogas",
            "consumo_de_drogas_en_espacios_publicos",
        },
        "consumo_de_drogas_en_espacios_publicos": {
            "consumo_de_drogas_en_espacios_publicos",
        },
        "contaminacion_sonica": {
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

    if base in alias_map:
        aliases.update(alias_map[base])

    if is_estafa_special(file_type, question_num):
        if "estafa" in base and "informatica" in base:
            aliases = {
                "estafa_informatica",
                "fraude_informatico",
                "estafa_por_medios_informaticos",
            }
        elif "estafa" in base:
            aliases = {
                "estafa",
                "estafa_telefonica",
                "estafa_bancaria",
                "estafa_por_redes_sociales",
                "estafa_por_medio_electronico",
                "estafa_documental",
                "estafa_simple",
                "estafa_en_compras",
                "estafa_comercial",
            }

    aliases = {normalize_token_for_compare(a) for a in aliases if a}
    aliases = {a for a in aliases if not is_unproductive_option(a)}
    return aliases


def build_group_definitions(file_type: str, question_num: str, question_text: str, items: list):
    """
    Para preguntas exactas:
      devuelve grupos por descriptor.
    Para preguntas unificadas:
      devuelve un solo grupo con el nombre de la pregunta.
    """
    if not is_exact_question(file_type, question_num):
        unified_label = get_unified_label(file_type, question_num, question_text)
        aliases = set()

        for item in items:
            aliases.update(build_descriptor_aliases(file_type, question_num, item["descriptor_texto"]))

        return {
            unified_label: {
                "group_label": unified_label,
                "aliases": aliases,
                "source_descriptors": {item["descriptor_texto"] for item in items},
            }
        }

    groups = {}
    for item in items:
        label = clean_descriptor_display(item["descriptor_texto"])
        if label not in groups:
            groups[label] = {
                "group_label": label,
                "aliases": set(),
                "source_descriptors": set(),
            }
        groups[label]["aliases"].update(
            build_descriptor_aliases(file_type, question_num, item["descriptor_texto"])
        )
        groups[label]["source_descriptors"].add(item["descriptor_texto"])
    return groups


def count_group_exact(series: pd.Series, aliases: set):
    total = 0
    matched_tokens = Counter()

    for val in series:
        row_tokens = tokenize_cell_unique(val)
        row_hit_tokens = set()

        for token in row_tokens:
            if token in aliases:
                row_hit_tokens.add(token)

        total += len(row_hit_tokens)
        for t in row_hit_tokens:
            matched_tokens[t] += 1

    return total, matched_tokens


def count_group_unified(series: pd.Series, aliases: set):
    total = 0
    matched_tokens = Counter()

    for val in series:
        row_tokens = tokenize_cell_unique(val)
        row_hit_tokens = set()

        for token in row_tokens:
            if token in aliases:
                row_hit_tokens.add(token)

        if row_hit_tokens:
            total += 1
            for t in row_hit_tokens:
                matched_tokens[t] += 1

    return total, matched_tokens


def find_unmapped_tokens(series: pd.Series, matched_aliases_union: set):
    counter = Counter()
    blank_rows = 0

    for val in series:
        row_tokens = tokenize_cell_unique(val)

        if not row_tokens:
            blank_rows += 1
            continue

        for token in row_tokens:
            if token not in matched_aliases_union and not is_unproductive_option(token):
                counter[token] += 1

    return counter, blank_rows


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
        exact_mode = is_exact_question(file_type, preg_num)
        mode_label = "exacto" if exact_mode else "unificado"

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

        for group_label, group_info in group_defs.items():
            aliases = group_info["aliases"]
            matched_aliases_union.update(aliases)

            if exact_mode:
                total, matched_tokens = count_group_exact(series, aliases)
            else:
                total, matched_tokens = count_group_unified(series, aliases)

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

        unmapped_counter, blank_rows = find_unmapped_tokens(series, matched_aliases_union)

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

# quitar ceros
df_results_all = remove_zero_rows(df_results_all, "cantidad_respuestas")
df_summary = remove_zero_rows(df_summary, "cantidad_respuestas")
df_totals = remove_zero_rows(df_totals, "cantidad_respuestas")

if not df_unmapped_all.empty:
    df_unmapped_all = remove_zero_rows(df_unmapped_all, "cantidad")
    df_unmapped_all = df_unmapped_all[df_unmapped_all["opcion_csv_no_ubicada"] != "[filas_vacias_ignoradas]"].copy()

# filtros
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

# métricas
st.markdown("## Resumen general")

m1, m2, m3, m4 = st.columns(4)

m1.metric("Archivos procesados", len(df_results_all["archivo"].unique()) if not df_results_all.empty else 0)
m2.metric("Preguntas detectadas", len(df_results_all["pregunta_num"].unique()) if not df_results_all.empty else 0)
m3.metric("Resultados mostrados", len(df_results_all) if not df_results_all.empty else 0)
m4.metric("Respuestas contabilizadas", int(df_results_all["cantidad_respuestas"].sum()) if not df_results_all.empty else 0)

# tabs
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

# descarga
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
