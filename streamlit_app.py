import streamlit as st
import pandas as pd
import numpy as np
import xml.etree.ElementTree as ET
from bs4 import BeautifulSoup
import pdfplumber
from io import BytesIO
import re

# Настройка страницы
st.set_page_config(page_title="Анализ смет", layout="wide")
st.title("🔍 Сравнение сметы и ВОР")

# === Функция: парсинг PDF ===
def parse_pdf(uploaded_file):
    tables = []
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df_page = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df_page)
        return pd.concat(tables, ignore_index=True) if tables else None
    except Exception as e:
        st.error(f"Ошибка при чтении PDF: {e}")
        return None

# === Функция: парсинг XML ===
def parse_xml(uploaded_file):
    try:
        tree = ET.parse(uploaded_file)
        root = tree.getroot()
        data = []

        for work in root.findall(".//Work") + root.findall(".//Row") + root.findall(".//Работа"):
            kod = work.find("Code") or work.find("Number") or work.find("Код") or work.find("Шифр")
            name = work.find("Name") or work.find("Description") or work.find("Наименование")
            volume = work.find("Volume") or work.find("Amount") or work.find("Объем") or work.find("Количество")

            data.append({
                "Шифр": kod.text.strip() if kod is not None else "",
                "Наименование": name.text.strip() if name is not None else "",
                "Объём": float(re.sub(r"[^\d.]", "", volume.text.replace(",", "."))) 
                         if volume is not None and volume.text and re.search(r"\d", volume.text) else np.nan
            })
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Ошибка при чтении XML: {e}")
        return None

# === Функция: парсинг HTML ===
def parse_html(uploaded_file):
    try:
        content = uploaded_file.read().decode("utf-8", errors="ignore")
        dfs = pd.read_html(content, encoding="utf-8")
        return dfs[0] if len(dfs) > 0 else None
    except Exception as e:
        st.error(f"Ошибка при чтении HTML: {e}")
        return None

# === Функция: загрузка файла ===
def load_file(uploaded_file):
    if uploaded_file.name.endswith(".xlsx"):
        return pd.read_excel(uploaded_file)
    elif uploaded_file.name.endswith(".pdf"):
        return parse_pdf(uploaded_file)
    elif uploaded_file.name.endswith(".xml"):
        uploaded_file.seek(0)
        return parse_xml(uploaded_file)
    elif uploaded_file.name.endswith((".html", ".htm")):
        uploaded_file.seek(0)
        return parse_html(uploaded_file)
    else:
        st.error("Формат не поддерживается")
        return None

# === Автоопределение столбцов ===
def auto_detect_columns(df):
    columns_lower = [str(col).strip().lower() for col in df.columns]
    mapping = {}

    # Поиск шифра
    for i, col in enumerate(columns_lower):
        if any(k in col for k in ["шифр", "код", "расценка", "number", "code", "id", "номер"]):
            mapping["key"] = df.columns[i]
            break

    # Поиск наименования
    for i, col in enumerate(columns_lower):
        if any(k in col for k in ["наименование", "работа", "описание", "name", "description", "наим"]):
            mapping["name"] = df.columns[i]
            break

    # Поиск объёма
    for i, col in enumerate(columns_lower):
        if any(k in col for k in ["объём", "объем", "количество", "кол-во", "qty", "amount", "volume"]):
            mapping["volume"] = df.columns[i]
            break

    return mapping

# === Интерфейс: загрузка файлов ===
st.header("1. Загрузите файлы")

col1, col2 = st.columns(2)
with col1:
    file_smeta = st.file_uploader("Сметный расчёт", type=["xlsx", "pdf", "xml", "html", "htm"], key="smeta")
with col2:
    file_vor = st.file_uploader("ВОР (объёмы работ)", type=["xlsx", "pdf", "xml", "html", "htm"], key="vor")

# === Настройки ===
st.header("2. Настройки сравнения")
match_by = st.radio("Сопоставлять позиции по:", ["Шифр расценки", "Наименование работ"])
threshold = st.slider("Порог расхождения объёмов (%)", 0, 100, 5)

# === Обработка файлов ===
if file_smeta and file_vor:
    with st.spinner("Чтение файлов..."):
        df_smeta = load_file(file_smeta)
        df_vor = load_file(file_vor)

    if df_smeta is None or df_vor is None:
        st.error("Не удалось прочитать один из файлов")
    else:
        st.success("Файлы загружены!")

        # Автоопределение
        detected_smeta = auto_detect_columns(df_smeta)
        detected_vor = auto_detect_columns(df_vor)

        st.header("3. Настройка столбцов (можно изменить)")

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Смета")
            options_smeta = list(df_smeta.columns)
            key_smeta = st.selectbox(
                "Столбец для сопоставления",
                options_smeta,
                index=options_smeta.index(detected_smeta.get("key" if match_by == "Шифр расценки" else "name")) 
                if (match_by == "Шифр расценки" and "key" in detected_smeta) or 
                   (match_by != "Шифр расценки" and "name" in detected_smeta) 
                else 0
            )
            vol_smeta = st.selectbox(
                "Столбец с объёмом",
                options_smeta,
                index=options_smeta.index(detected_smeta["volume"]) 
                if "volume" in detected_smeta else 0
            )

        with col2:
            st.subheader("ВОР")
            options_vor = list(df_vor.columns)
            key_vor = st.selectbox(
                "Столбец для сопоставления",
                options_vor,
                index=options_vor.index(detected_vor.get("key" if match_by == "Шифр расценки" else "name")) 
                if (match_by == "Шифр расценки" and "key" in detected_vor) or 
                   (match_by != "Шифр расценки" and "name" in detected_vor) 
                else 0
            )
            vol_vor = st.selectbox(
                "Столбец с объёмом",
                options_vor,
                index=options_vor.index(detected_vor["volume"]) 
                if "volume" in detected_vor else 0
            )

        # Приведение данных
        df_smeta[key_smeta] = df_smeta[key_smeta].astype(str).str.strip()
        df_vor[key_vor] = df_vor[key_vor].astype(str).str.strip()

        if match_by == "Наименование работ":
            df_smeta[key_smeta] = df_smeta[key_smeta].str.lower()
            df_vor[key_vor] = df_vor[key_vor].str.lower()

        df_smeta[vol_smeta] = pd.to_numeric(
            df_smeta[vol_smeta].astype(str).str.replace(",", ".").str.replace(r"[^\d.-]", "", regex=True),
            errors='coerce'
        )
        df_vor[vol_vor] = pd.to_numeric(
            df_vor[vol_vor].astype(str).str.replace(",", ".").str.replace(r"[^\d.-]", "", regex=True),
            errors='coerce'
        )

        # === Слияние и сравнение ===
        merged = pd.merge(
            df_smeta, df_vor,
            left_on=key_smeta, right_on=key_vor,
            how="outer", indicator=True,
            suffixes=("_смета", "_ВОР")
        )

        merge_col = "Шифр" if match_by == "Шифр расценки" else "Наименование"
        merged[merge_col] = merged[key_smeta].fillna(merged[key_vor])

        merged["Объём_смета"] = merged[vol_smeta + "_смета"]
        merged["Объём_ВОР"] = merged[vol_vor + "_ВОР"]

        # Разница в %
        valid_max = merged[["Объём_смета", "Объём_ВОР"]].max(axis=1).replace(0, np.nan)
        merged["Разница_%"] = (
            (merged["Объём_смета"] - merged["Объём_ВОР"]).abs() / valid_max * 100
        ).round(2)

        # Статус
        merged["Статус"] = np.where(
            merged["Разница_%"] > threshold, "Расхождение",
            np.where(merged["_merge"] == "both", "Совпадает",
            np.where(merged["_merge"] == "left_only", "Нет в ВОР", "Нет в смете"))
        )

        # === Результат ===
        st.header("4. Результат сравнения")
        result = merged[[merge_col, "Объём_смета", "Объём_ВОР", "Разница_%", "Статус"]].copy()

        def highlight_status(row):
            if row["Статус"] == "Расхождение":
                return ["background-color: #ffcccc"] * len(row)
            elif "Нет" in row["Статус"]:
                return ["background-color: #ffffcc"] * len(row)
            else:
                return ["background-color: #ccffcc"] * len(row)

        st.dataframe(
            result.style.format({
                "Объём_смета": "{:.3f}", "Объём_ВОР": "{:.3f}", "Разница_%": "{}%"
            }).apply(highlight_status, axis=1),
            use_container_width=True
        )

        # Экспорт
        @st.cache_data
        def convert_df_to_excel(df):
            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            return buffer.getvalue()

        output = convert_df_to_excel(merged)
        st.download_button(
            "📥 Скачать полный отчёт (Excel)",
            output,
            "сравнение_сметы_и_ВОР.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Загрузите оба файла для анализа.")
