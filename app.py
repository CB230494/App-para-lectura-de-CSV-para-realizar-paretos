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

TOKENS_IGNORAR = {
    "",
    "nan",
    "none",
    "null",
}

OPCIONES_NO_PRODUCTIVAS = {
    "otro",
    "otros",
    "otra",
    "otras",
    "otro problema",
    "otro problema que considere importante",
    "otros delitos",
    "otros delitos cuales",
    "cual",
    "cuales",
    "especifique",
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
    s = re.sub(r"[^\w\s]", "", s, flags=re.UNICODE)
    s = s.replace("/", " ")
    s = re.sub(r"\s+", "_", s)
    s = s.strip("_")
    return s


def normalize_option_token(text) -> str:
    """
    Normaliza opciones del CSV y del Excel para compararlas.
    """
    s = slugify(text)
    s = re.sub(r"_+", "_", s).strip("_")
    return s


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
    s = norm(value)
    return s in TOKENS_IGNORAR


# =========================================================
# DETECCIÓN DE TIPO DE ARCHIVO
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
# LECTURA DEL EXCEL GUÍA
# =========================================================
def load_guide_excel(path_excel: str):
    """
    Devuelve:
    {
      hoja: [
        {
          pregunta_num,
          pregunta_texto,
          pregunta_slug,
          descriptor_texto,
          descriptor_slug
        }
      ]
    }
    """
    xls = pd.ExcelFile(path_excel)
    guide = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(path_excel, sheet_name=sheet, header=None)
        df = df.fillna("")

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

                possible_descriptors = vals_nonempty[1:]
                for desc in possible_descriptors:
                    desc_clean = str(desc).strip()
                    if desc_clean:
                        rows.append({
                            "pregunta_num": current_question_num,
                            "pregunta_texto": current_question_text,
                            "pregunta_slug": normalize_option_token(current_question_text),
                            "descriptor_texto": desc_clean,
                            "descriptor_slug": normalize_option_token(desc_clean),
                        })
                continue

            if current_question_num and current_question_text:
                rows.append({
                    "pregunta_num": current_question_num,
                    "pregunta_texto": current_question_text,
                    "pregunta_slug": normalize_option_token(current_question_text),
                    "descriptor_texto": first,
                    "descriptor_slug": normalize_option_token(first),
                })

                for extra in vals_nonempty[1:]:
                    extra_clean = str(extra).strip()
                    if extra_clean:
                        rows.append({
                            "pregunta_num": current_question_num,
                            "pregunta_texto": current_question_text,
                            "pregunta_slug": normalize_option_token(current_question_text),
                            "descriptor_texto": extra_clean,
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
# LECTURA ROBUSTA DE CSV
# =========================================================
def parse_csv_with_python_engine(content: bytes, encoding: str, delimiter: str):
    text = content.decode(encoding, errors="replace")
    rows = []

    reader = csv.reader(
        io.StringIO(text),
        delimiter=delimiter,
        quotechar='"'
    )

    max_cols = 0
    for row in reader:
        rows.append(row)
        if len(row) > max_cols:
            max_cols = len(row)

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

    df = pd.DataFrame(data, columns=header)
    return df.fillna("")


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
                return df.fillna("")
        except Exception as e:
            last_error = e

    try:
        attempts_pd = [
            {"sep": ",", "encoding": "utf-8-sig"},
            {"sep": ",", "encoding": "utf-8"},
            {"sep": ",", "encoding": "latin-1"},
            {"sep": ";", "encoding": "utf-8-sig"},
            {"sep": ";", "encoding": "utf-8"},
            {"sep": ";", "encoding": "latin-1"},
        ]

        for at in attempts_pd:
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
# UBICACIÓN DE PREGUNTAS EN EL CSV
# =========================================================
def build_question_groups(guide_sheet_rows: list):
    grouped = defaultdict(list)
    for r in guide_sheet_rows:
        grouped[(r["pregunta_num"], r["pregunta_texto"])].append(r)
    return grouped


def question_header_candidates(question_num: str, question_text: str):
    """
    Genera candidatos para encontrar la columna de la pregunta en el CSV.
    """
    qnorm = norm(question_text)
    qslug = normalize_option_token(question_text)

    candidates = set()

    if question_num:
        candidates.add(question_num)
        candidates.add(f"{question_num}.")
        candidates.add(f"{question_num})")

    candidates.add(qnorm)
    candidates.add(qslug)

    # Agrega versión del texto sin el número
    text_wo_num = re.sub(r"^\s*\d+(?:\.\d+)?\s*[\.\)]?\s*", "", question_text).strip()
    if text_wo_num:
        candidates.add(norm(text_wo_num))
        candidates.add(normalize_option_token(text_wo_num))

    return [c for c in candidates if c]


def score_question_column(col_name: str, question_num: str, question_text: str) -> int:
    """
    Puntaje para escoger la mejor columna del CSV para una pregunta.
    """
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
# CONTEO DE OPCIONES DENTRO DE UNA CELDA
# =========================================================
def split_multiselect_cell(value: str):
    """
    Divide respuestas múltiples separadas por coma.
    """
    if is_effectively_empty(value):
        return []

    raw = str(value).strip()
    if not raw:
        return []

    parts = [p.strip() for p in raw.split(",")]
    parts = [p for p in parts if p and norm(p) not in TOKENS_IGNORAR]
    return parts


def is_unproductive_option(token_norm: str) -> bool:
    if token_norm in OPCIONES_NO_PRODUCTIVAS:
        return True

    # Variantes que empiecen con "otro_"
    if token_norm.startswith("otro_") or token_norm.startswith("otros_"):
        return True

    return False


def count_descriptor_occurrences_in_question_column(series: pd.Series):
    """
    Cuenta las opciones reales seleccionadas dentro de una columna de pregunta.
    Devuelve:
      - contador por token normalizado
      - tokens no ubicados
      - filas vacías
    """
    counter = Counter()
    unmapped_counter = Counter()
    blank_rows = 0

    for val in series:
        options = split_multiselect_cell(val)

        if not options:
            blank_rows += 1
            continue

        for opt in options:
            token = normalize_option_token(opt)

            if not token or token in TOKENS_IGNORAR:
                continue

            if is_unproductive_option(token):
                continue

            counter[token] += 1

    return counter, unmapped_counter, blank_rows


# =========================================================
# MAPEO EXCEL <-> RESPUESTAS DEL CSV
# =========================================================
def build_descriptor_aliases(descriptor_text: str):
    """
    Genera alias para mejorar la coincidencia entre Excel y CSV.
    """
    base = normalize_option_token(descriptor_text)
    aliases = {base}

    # Ajustes puntuales frecuentes
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
            "presencia_de_personas_en_situacion_de_calle_personas_que_viven_permanentemente_en_la_via_publica",
            "presencia_de_personas_en_situacion_de_calle",
        },
        "ventas_informales_ambulantes": {
            "ventas_informales_ambulantes",
        },
        "problemas_vecinales_o_conflictos_entre_vecinos": {
            "problemas_vecinales_o_conflictos_entre_vecinos",
        },
        "desvinculacion_escolar_desercion_escolar": {
            "desvinculacion_escolar_desercion_escolar",
        },
        "perdida_de_espacios_publicos_parques_polideportivos_u_otros": {
            "perdida_de_espacios_publicos_parques_polideportivos_u_otros",
        },
        "acumulacion_de_basura_aguas_negras_o_mal_alcantarillado": {
            "acumulacion_de_basura_aguas_negras_o_mal_alcantarillado",
        },
        "falta_de_oportunidades_laborales": {
            "falta_de_oportunidades_laborales",
        },
        "asentamientos_informales_o_precarios": {
            "asentamientos_informales_o_precarios",
        },
        "lotes_baldios": {
            "lotes_baldios",
        },
        "cuarterias": {
            "cuarterias",
        },
        "consumo_de_alcohol_en_via_publica": {
            "consumo_de_alcohol_en_via_publica",
        },
        "no_se_observan_estas_problematicas_en_el_distrito": {
            "no_se_observan_estas_problematicas_en_el_distrito",
        },
    }

    if base in alias_map:
        aliases.update(alias_map[base])

    return aliases


def match_descriptor_count(descriptor_text: str, option_counter: Counter):
    """
    Toma un descriptor del Excel y devuelve cuántas veces aparece
    en las opciones reales del CSV.
    """
    aliases = build_descriptor_aliases(descriptor_text)

    total = 0
    matched_tokens = []

    for token, cnt in option_counter.items():
        if token in aliases:
            total += cnt
            matched_tokens.append(token)

    return total, sorted(set(matched_tokens))


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

    grouped = build_question_groups(base)

    for (preg_num, preg_text), items in grouped.items():
        question_col, score = find_question_column(df_csv, preg_num, preg_text)

        if not question_col:
            for item in items:
                results.append({
                    "archivo": filename,
                    "tipo": file_type,
                    "hoja_excel": sheet_name.strip(),
                    "pregunta_num": preg_num,
                    "pregunta": preg_text,
                    "descriptor": item["descriptor_texto"],
                    "columna_pregunta_csv": "",
                    "opciones_csv_que_contaron": "",
                    "cantidad_respuestas": 0,
                })

                mapping_info.append({
                    "archivo": filename,
                    "tipo": file_type,
                    "pregunta_num": preg_num,
                    "pregunta": preg_text,
                    "descriptor": item["descriptor_texto"],
                    "columna_pregunta_csv": "",
                    "puntaje_columna": score,
                    "mapeado": "No",
                    "motivo": "No se encontró columna de la pregunta en el CSV",
                })
            continue

        option_counter, _, blank_rows = count_descriptor_occurrences_in_question_column(df_csv[question_col])

        matched_any_token = set()

        for item in items:
            descriptor = item["descriptor_texto"]
            total, matched_tokens = match_descriptor_count(descriptor, option_counter)

            for mt in matched_tokens:
                matched_any_token.add(mt)

            results.append({
                "archivo": filename,
                "tipo": file_type,
                "hoja_excel": sheet_name.strip(),
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "descriptor": descriptor,
                "columna_pregunta_csv": question_col,
                "opciones_csv_que_contaron": " | ".join(matched_tokens),
                "cantidad_respuestas": int(total),
            })

            mapping_info.append({
                "archivo": filename,
                "tipo": file_type,
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "descriptor": descriptor,
                "columna_pregunta_csv": question_col,
                "puntaje_columna": score,
                "mapeado": "Sí" if matched_tokens or total == 0 else "Sí",
                "motivo": "Conteo por opciones dentro de la celda",
            })

        # Opciones del CSV que no se lograron ubicar en el Excel
        for token, cnt in option_counter.items():
            if token not in matched_any_token and not is_unproductive_option(token):
                unmapped_options_rows.append({
                    "archivo": filename,
                    "tipo": file_type,
                    "pregunta_num": preg_num,
                    "pregunta": preg_text,
                    "columna_pregunta_csv": question_col,
                    "opcion_csv_no_ubicada": token,
                    "cantidad": int(cnt),
                })

        # Registro informativo de filas vacías por pregunta
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
        .groupby(["archivo", "tipo", "pregunta_num", "pregunta"], as_index=False)["cantidad_respuestas"]
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
        .groupby(["tipo", "pregunta_num", "pregunta", "descriptor"], as_index=False)["cantidad_respuestas"]
        .sum()
    )

    totals["sort_key"] = totals["pregunta_num"].apply(question_sort_key)
    totals = totals.sort_values(
        by=["tipo", "sort_key", "descriptor"],
        kind="stable"
    ).drop(columns=["sort_key"])

    return totals


# =========================================================
# EXPORTAR A EXCEL
# =========================================================
def to_excel_bytes(dfs: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


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

# =========================================================
# FILTROS
# =========================================================
st.markdown("## Filtros")

colf1, colf2, colf3 = st.columns(3)

tipos = sorted(df_results_all["tipo"].dropna().unique().tolist())
preguntas = sorted(df_results_all["pregunta_num"].dropna().unique().tolist(), key=question_sort_key)
archivos = sorted(df_results_all["archivo"].dropna().unique().tolist())

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

# =========================================================
# MÉTRICAS
# =========================================================
st.markdown("## Resumen general")

m1, m2, m3, m4 = st.columns(4)

m1.metric("Archivos procesados", len(df_results_all["archivo"].unique()))
m2.metric("Preguntas detectadas", len(df_results_all["pregunta_num"].unique()))
m3.metric("Descriptores evaluados", len(df_results_all))
m4.metric("Respuestas contabilizadas", int(df_results_all["cantidad_respuestas"].sum()))

# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Totales por descriptor",
    "Resumen por pregunta",
    "Detalle",
    "Mapeo pregunta ↔ CSV",
    "Opciones no ubicadas"
])

with tab1:
    st.subheader("Totales por descriptor")
    st.dataframe(df_totals_f, use_container_width=True)

with tab2:
    st.subheader("Resumen por pregunta")
    st.dataframe(df_summary_f, use_container_width=True)

with tab3:
    st.subheader("Detalle por archivo")
    st.dataframe(df_results_f, use_container_width=True)

with tab4:
    st.subheader("Mapeo pregunta ↔ CSV")
    st.dataframe(df_mapping_f, use_container_width=True)

with tab5:
    st.subheader("Opciones del CSV no ubicadas en el Excel")
    if df_unmapped_f.empty:
        st.info("No hay opciones no ubicadas.")
    else:
        st.dataframe(df_unmapped_f, use_container_width=True)

# =========================================================
# DESCARGA
# =========================================================
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
