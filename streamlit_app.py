import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# Настройка страницы
st.set_page_config(page_title="Анализ смет", layout="wide")
st.title("🔍 Анализ сметных расчётов и ВОР")

# === Загрузка файлов ===
st.header("1. Загрузите файлы")

col1, col2 = st.columns(2)
with col1:
    file_smeta = st.file_uploader("Сметный расчёт (Excel)", type=["xlsx"], key="smeta")
with col2:
    file_vor = st.file_uploader("Ведомость объёмов работ (ВОР)", type=["xlsx"], key="vor")

# === Настройки сравнения ===
st.header("2. Настройки сравнения")

match_by = st.radio(
    "Сопоставлять позиции по:",
    options=["Шифр расценки", "Наименование работ"],
    help="Выберите ключ для сопоставления строк"
)

threshold = st.slider(
    "Порог расхождения объёмов (%)", 
    min_value=0, max_value=100, value=5,
    help="Расхождения больше этого значения будут выделены"
)

# === Обработка данных ===
if file_smeta and file_vor:
    try:
        # Читаем файлы
        df_smeta = pd.read_excel(file_smeta)
        df_vor = pd.read_excel(file_vor)

        st.success("Файлы успешно загружены!")

        # Показываем первые строки
        with st.expander("Посмотреть данные сметы"):
            st.dataframe(df_smeta.head())

        with st.expander("Посмотреть данные ВОР"):
            st.dataframe(df_vor.head())

        # === Подготовка данных ===
        df_smeta.columns = [str(col).strip().lower() for col in df_smeta.columns]
        df_vor.columns = [str(col).strip().lower() for col in df_vor.columns]

        # === Выбор ключа и объёмов ===
        if match_by == "Шифр расценки":
            key_smeta = st.selectbox("Выберите столбец 'Шифр' в смете", df_smeta.columns, key="key_smeta")
            key_vor = st.selectbox("Выберите столбец 'Шифр' в ВОР", df_vor.columns, key="key_vor")
            smeta_key_col = df_smeta[key_smeta].astype(str).str.strip()
            vor_key_col = df_vor[key_vor].astype(str).str.strip()
        else:
            key_smeta = st.selectbox("Выберите столбец 'Наименование' в смете", df_smeta.columns, key="name_smeta")
            key_vor = st.selectbox("Выберите столбец 'Наименование' в ВОР", df_vor.columns, key="name_vor")
            smeta_key_col = df_smeta[key_smeta].astype(str).str.lower().str.strip()
            vor_key_col = df_vor[key_vor].astype(str).str.lower().str.strip()

        # Выбор столбца с объёмами
        vol_smeta = st.selectbox("Столбец 'Объём' в смете", df_smeta.columns, key="vol_smeta")
        vol_vor = st.selectbox("Столбец 'Объём' в ВОР", df_vor.columns, key="vol_vor")

        # Приводим объёмы к числу
        df_smeta[vol_smeta] = pd.to_numeric(df_smeta[vol_smeta], errors='coerce')
        df_vor[vol_vor] = pd.to_numeric(df_vor[vol_vor], errors='coerce')

        # === Слияние данных ===
        merged = pd.DataFrame()
        if match_by == "Шифр расценки":
            merged = pd.merge(
                df_smeta, df_vor,
                left_on=key_smeta, right_on=key_vor,
                suffixes=('_смета', '_ВОР'),
                how='outer', indicator=True
            )
            merge_col = "Шифр"
            left_key = f"{key_smeta}_смета"
            right_key = key_vor
        else:
            merged = pd.merge(
                df_smeta, df_vor,
                left_on=key_smeta, right_on=key_vor,
                suffixes=('_смета', '_ВОР'),
                how='outer', indicator=True
            )
            merge_col = "Наименование"
            left_key = f"{key_smeta}_смета"
            right_key = key_vor

        # Формируем ключ
        merged[merge_col] = merged[left_key].fillna(merged[right_key])

        # Объёмы
        merged["Объём_смета"] = merged[vol_smeta + "_смета"]
        merged["Объём_ВОР"] = merged[vol_vor + "_ВОР"]

        # Расчёт разницы
        max_vol = merged[["Объём_смета", "Объём_ВОР"]].max(axis=1).replace(0, np.nan)
        merged["Разница_объёмов"] = (
            (merged["Объём_смета"] - merged["Объём_ВОР"]).abs() / max_vol * 100
        ).round(2)

        # Статус
        merged["Статус"] = "Совпадает"
        merged.loc[merged["Разница_объёмов"] > threshold, "Статус"] = "Расхождение"
        merged.loc[merged["Объём_смета"].isna(), "Статус"] = "Нет в смете"
        merged.loc[merged["Объём_ВОР"].isna(), "Статус"] = "Нет в ВОР"

        # === Отображение результата ===
        st.header("3. Результат сравнения")

        filter_status = st.multiselect(
            "Фильтр по статусу",
            options=merged["Статус"].unique(),
            default=merged["Статус"].unique()
        )

        result = merged[merged["Статус"].isin(filter_status)][
            [merge_col, "Объём_смета", "Объём_ВОР", "Разница_объёмов", "Статус"]
        ].copy()

        # Цветовая подсветка
        def color_status(val):
            if val == "Расхождение":
                return "background-color: #ffcccc"
            elif val == "Нет в смете" or val == "Нет в ВОР":
                return "background-color: #ffffcc"
            else:
                return "background-color: #ccffcc"

        st.dataframe(
            result.style.format({
                "Объём_смета": "{:.3f}",
                "Объём_ВОР": "{:.3f}",
                "Разница_объёмов": "{}%"
            }).applymap(color_status, subset=["Статус"]),
            use_container_width=True
        )

        # === Экспорт отчёта ===
        @st.cache_data
        def convert_df_to_excel(df):
            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine='openpyxl')
            buffer.seek(0)
            return buffer.getvalue()

        output = convert_df_to_excel(merged)
        st.download_button(
            label="📥 Скачать отчёт (Excel)",
            data=output,
            file_name="сравнение_сметы_и_ВОР.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ошибка при обработке файлов: {e}")
        st.exception(e)
else:
    st.info("Загрузите оба файла для начала анализа.")