# app.py
# -*- coding: utf-8 -*-

import io
import re
import csv
import unicodedata
from pathlib import Path
from collections import defaultdict

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


def extract_question_number(text: str) -> str:
    """
    Extrae numeración tipo:
    12.
    20.
    31.4
    """
    s = str(text).strip()
    m = re.match(r"^\s*(\d+(?:\.\d+)?)\s*[\.\)]?", s)
    return m.group(1) if m else ""


def question_sort_key(q):
    """
    Orden natural para preguntas tipo 12, 12.1, 31.4.
    Devuelve una tupla para que pandas pueda ordenar sin error.
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
    if value is None:
        return True
    s = str(value).strip()
    if s == "":
        return True
    if norm(s) in {"nan", "none", "null"}:
        return True
    return False


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
    Lee el Excel guía y devuelve una estructura:
    {
      "Comunidad ": [
         {
            "pregunta_num": "12",
            "pregunta_texto": "...",
            "descriptor_texto": "...",
            "descriptor_slug": "..."
         },
         ...
      ],
      ...
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
            vals_nonempty = [v for v in vals if v and norm(v) not in {"nan", "none"}]

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
                            "descriptor_texto": desc_clean,
                            "descriptor_slug": slugify(desc_clean),
                        })
                continue

            if current_question_num and current_question_text:
                rows.append({
                    "pregunta_num": current_question_num,
                    "pregunta_texto": current_question_text,
                    "descriptor_texto": first,
                    "descriptor_slug": slugify(first),
                })

                for extra in vals_nonempty[1:]:
                    extra_clean = str(extra).strip()
                    if extra_clean:
                        rows.append({
                            "pregunta_num": current_question_num,
                            "pregunta_texto": current_question_text,
                            "descriptor_texto": extra_clean,
                            "descriptor_slug": slugify(extra_clean),
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
    """
    Lectura robusta para CSV de Survey123.
    Intenta con varios separadores y encodings.
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
# MAPEO DESCRIPTOR -> COLUMNAS CSV
# =========================================================
def find_matching_columns(df: pd.DataFrame, descriptor_slug: str):
    matches = []

    for col in df.columns:
        cslug = slugify(col)

        if cslug == descriptor_slug:
            matches.append(col)
            continue

        if descriptor_slug and descriptor_slug in cslug:
            matches.append(col)
            continue

        if cslug and cslug in descriptor_slug:
            matches.append(col)
            continue

    return matches


def count_answers_in_columns(df: pd.DataFrame, cols: list) -> int:
    """
    Si una fila tiene valor en cualquiera de las columnas mapeadas, cuenta 1.
    """
    if not cols:
        return 0

    subset = df[cols].copy()

    def row_has_answer(row):
        for val in row:
            if not is_effectively_empty(val):
                sval = norm(val)
                if sval not in {"0", "false", "no", "n"}:
                    return True
        return False

    return int(subset.apply(row_has_answer, axis=1).sum())


# =========================================================
# PROCESAMIENTO PRINCIPAL
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

    grouped = defaultdict(list)
    for r in base:
        grouped[(r["pregunta_num"], r["pregunta_texto"])].append(r)

    for (preg_num, preg_text), items in grouped.items():
        for item in items:
            descriptor = item["descriptor_texto"]
            descriptor_slug = item["descriptor_slug"]

            matches = find_matching_columns(df_csv, descriptor_slug)
            count = count_answers_in_columns(df_csv, matches)

            results.append({
                "archivo": filename,
                "tipo": file_type,
                "hoja_excel": sheet_name.strip(),
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "descriptor": descriptor,
                "columnas_encontradas": " | ".join(matches) if matches else "",
                "cantidad_respuestas": count,
            })

            mapping_info.append({
                "archivo": filename,
                "tipo": file_type,
                "pregunta_num": preg_num,
                "pregunta": preg_text,
                "descriptor": descriptor,
                "descriptor_slug": descriptor_slug,
                "columnas_encontradas": " | ".join(matches) if matches else "",
                "mapeado": "Sí" if matches else "No",
            })

    df_results = pd.DataFrame(results)
    df_mapping = pd.DataFrame(mapping_info)

    return df_results, df_mapping


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
        preguntas = df_tmp["pregunta_num"].nunique()
        descriptores = len(df_tmp)

        rows.append({
            "hoja": sh.strip(),
            "preguntas_detectadas": preguntas,
            "descriptores_detectados": descriptores,
        })

    return pd.DataFrame(rows)


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
st.caption("Carga el Excel guía desde el repositorio y luego sube los CSV de Comunidad, Comercio y Policial.")

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
read_errors = []

for file in uploaded_files:
    try:
        content = file.read()
        df_csv = try_read_csv_bytes(content)
        df_csv = flatten_headers(df_csv)

        df_results, df_mapping = build_results_for_file(df_csv, file.name, guide)

        all_results.append(df_results)
        all_mapping.append(df_mapping)

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
m4.metric("Descriptores mapeados", int((df_mapping_all["mapeado"] == "Sí").sum()))

# =========================================================
# TABS
# =========================================================
tab1, tab2, tab3, tab4 = st.tabs([
    "Resumen por pregunta",
    "Totales por descriptor",
    "Detalle por archivo",
    "Mapeo Excel ↔ CSV"
])

with tab1:
    st.subheader("Resumen por pregunta")
    st.dataframe(df_summary_f, use_container_width=True)

with tab2:
    st.subheader("Total global por descriptor")
    st.dataframe(df_totals_f, use_container_width=True)

with tab3:
    st.subheader("Detalle por archivo / descriptor")
    st.dataframe(df_results_f, use_container_width=True)

with tab4:
    st.subheader("Mapeo Excel ↔ CSV")
    st.dataframe(df_mapping_f, use_container_width=True)

# =========================================================
# DESCARGA
# =========================================================
excel_bytes = to_excel_bytes({
    "resumen_pregunta": df_summary_f,
    "totales_descriptor": df_totals_f,
    "detalle_archivo": df_results_f,
    "mapeo_excel_csv": df_mapping_f,
})

st.download_button(
    label="Descargar resultados en Excel",
    data=excel_bytes,
    file_name="conteo_respuestas_preguntas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
