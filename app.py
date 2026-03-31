# app.py
# -*- coding: utf-8 -*-

import io
import re
import unicodedata
from pathlib import Path
from collections import defaultdict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Conteo de respuestas por pregunta", layout="wide")


# =========================================================
# CONFIG
# =========================================================
EXCEL_GUIA = "Guía de Preguntas para paretos 2026.xlsx"

# Mapeo nombre hoja Excel según tipo de archivo
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
# DETECCIÓN TIPO DE ARCHIVO
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

            # Tomamos el primer valor no vacío como texto principal
            first = vals_nonempty[0]

            # Si parece una pregunta numerada, actualizamos contexto
            qnum = extract_question_number(first)
            if qnum:
                current_question_num = qnum
                current_question_text = first

                # Puede venir descriptor en otra columna de esa misma fila
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

            # Si no es pregunta, lo tomamos como descriptor de la pregunta actual
            if current_question_num and current_question_text:
                rows.append({
                    "pregunta_num": current_question_num,
                    "pregunta_texto": current_question_text,
                    "descriptor_texto": first,
                    "descriptor_slug": slugify(first),
                })

                # Si hay más columnas con texto, también las agregamos como posibles descriptores
                for extra in vals_nonempty[1:]:
                    extra_clean = str(extra).strip()
                    if extra_clean:
                        rows.append({
                            "pregunta_num": current_question_num,
                            "pregunta_texto": current_question_text,
                            "descriptor_texto": extra_clean,
                            "descriptor_slug": slugify(extra_clean),
                        })

        # Eliminar duplicados exactos
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
# LECTURA ROBUSTA CSV SURVEY123
# =========================================================
def try_read_csv_bytes(content: bytes) -> pd.DataFrame:
    """
    Intenta varias combinaciones de lectura porque los CSV pueden venir
    con encabezado irregular o formatos complicados.
    """
    attempts = [
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "utf-8"},
        {"sep": ",", "encoding": "latin-1"},
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "utf-8"},
        {"sep": ";", "encoding": "latin-1"},
    ]

    last_error = None

    for at in attempts:
        try:
            df = pd.read_csv(io.BytesIO(content), dtype=str, keep_default_na=False, **at)
            if df.shape[1] > 1:
                return df.fillna("")
        except Exception as e:
            last_error = e

    raise ValueError(f"No se pudo leer el CSV. Error: {last_error}")


def flatten_multiline_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    En algunos CSV los nombres vienen raros o con textos largos.
    Aquí solo normalizamos encabezados.
    """
    new_cols = []
    for c in df.columns:
        s = str(c).replace("\n", " ").replace("\r", " ").strip()
        s = re.sub(r"\s+", " ", s)
        new_cols.append(s)
    df.columns = new_cols
    return df


def find_matching_columns(df: pd.DataFrame, descriptor_slug: str):
    """
    Busca columnas del CSV que coincidan razonablemente con el slug del descriptor.
    """
    matches = []

    for col in df.columns:
        cslug = slugify(col)

        # coincidencia exacta
        if cslug == descriptor_slug:
            matches.append(col)
            continue

        # descriptor dentro del nombre de columna
        if descriptor_slug and descriptor_slug in cslug:
            matches.append(col)
            continue

        # nombre columna dentro del descriptor
        if cslug and cslug in descriptor_slug:
            matches.append(col)
            continue

    return matches


def count_answers_in_columns(df: pd.DataFrame, cols: list) -> int:
    """
    Cuenta respuestas efectivas en las columnas ubicadas para un descriptor.
    Si una fila marcó valor en cualquiera de esas columnas, cuenta 1.
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
    df_csv = flatten_multiline_headers(df_csv.copy())

    results = []
    mapping_info = []

    # Agrupar descriptores por pregunta
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
                "hoja_excel": sheet_name,
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
        .sort_values(["archivo", "pregunta_num", "pregunta"])
    )
    return summary


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
# UI
# =========================================================
st.title("Conteo de respuestas por pregunta / descriptor")
st.caption("Carga el Excel guía desde el repositorio y luego sube los CSV de Comunidad, Comercio y Policial.")

with st.sidebar:
    st.header("Configuración")
    excel_path = st.text_input("Nombre del Excel guía", value=EXCEL_GUIA)
    st.info(
        "El Excel debe estar en la raíz del repositorio.\n\n"
        "Hojas esperadas:\n"
        "- Comunidad \n"
        "- Comercio\n"
        "- Policia"
    )

# Cargar guía
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
        meta_rows = []
        for sh, rows in guide.items():
            df_tmp = pd.DataFrame(rows)
            preguntas = df_tmp["pregunta_num"].nunique() if not df_tmp.empty else 0
            descriptores = len(df_tmp)
            meta_rows.append({
                "hoja": sh,
                "preguntas_detectadas": preguntas,
                "descriptores_detectados": descriptores,
            })
        st.dataframe(pd.DataFrame(meta_rows), use_container_width=True)

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
        df_results, df_mapping = build_results_for_file(df_csv, file.name, guide)
        all_results.append(df_results)
        all_mapping.append(df_mapping)
    except Exception as e:
        read_errors.append({"archivo": file.name, "error": str(e)})

if read_errors:
    st.subheader("Errores detectados")
    st.dataframe(pd.DataFrame(read_errors), use_container_width=True)

if not all_results:
    st.error("No se pudo procesar ningún archivo.")
    st.stop()

df_results_all = pd.concat(all_results, ignore_index=True)
df_mapping_all = pd.concat(all_mapping, ignore_index=True)
df_summary = summarize_results(df_results_all)

# =========================================================
# VISTAS
# =========================================================
tab1, tab2, tab3 = st.tabs(["Resumen por pregunta", "Detalle por descriptor", "Mapeo Excel ↔ CSV"])

with tab1:
    st.subheader("Resumen por pregunta")
    st.dataframe(df_summary, use_container_width=True)

    st.markdown("### Filtros")
    tipos = sorted(df_summary["tipo"].dropna().unique().tolist())
    preguntas = sorted(df_summary["pregunta_num"].dropna().unique().tolist(), key=lambda x: [int(p) if p.isdigit() else p for p in x.split(".")])

    colf1, colf2 = st.columns(2)
    with colf1:
        filtro_tipo = st.multiselect("Tipo", options=tipos, default=tipos)
    with colf2:
        filtro_preg = st.multiselect("Pregunta", options=preguntas, default=preguntas)

    df_filtrado = df_summary[
        df_summary["tipo"].isin(filtro_tipo) &
        df_summary["pregunta_num"].isin(filtro_preg)
    ].copy()

    st.dataframe(df_filtrado, use_container_width=True)

with tab2:
    st.subheader("Detalle por descriptor")
    st.dataframe(df_results_all, use_container_width=True)

with tab3:
    st.subheader("Mapeo Excel ↔ CSV")
    st.dataframe(df_mapping_all, use_container_width=True)

# =========================================================
# DESCARGA
# =========================================================
excel_bytes = to_excel_bytes({
    "resumen_pregunta": df_summary,
    "detalle_descriptor": df_results_all,
    "mapeo_excel_csv": df_mapping_all,
})

st.download_button(
    label="Descargar resultados en Excel",
    data=excel_bytes,
    file_name="conteo_respuestas_preguntas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

