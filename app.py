"""Веб-интерфейс нормализатора городов Краснодарского края."""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from city_normalizer import (
    DEFAULT_THRESHOLD,
    detect_city_column,
    process_file,
    read_file,
    to_csv_bytes,
    to_excel_bytes,
)

st.set_page_config(
    page_title="Нормализатор городов Краснодарского края",
    page_icon="🗺️",
    layout="wide",
)

st.title("Нормализатор городов Краснодарского края")
st.caption(
    "Приводит столбец «город» к официальным названиям муниципальных образований."
)

uploaded = st.file_uploader(
    "Загрузите таблицу Excel или CSV",
    type=["xlsx", "xls", "xlsm", "csv"],
)

if uploaded is not None and st.session_state.get("uploaded_name") != uploaded.name:
    st.session_state.pop("result", None)
    st.session_state["uploaded_name"] = uploaded.name

if uploaded is None:
    st.info("Выберите файл, чтобы начать. Поддерживаются xlsx, xls, xlsm и csv.")
    st.stop()


@st.cache_data(show_spinner=False)
def _load_table(data: bytes, name: str) -> pd.DataFrame:
    return read_file(io.BytesIO(data), filename=name)


try:
    df = _load_table(uploaded.getvalue(), uploaded.name)
except Exception as exc:  # noqa: BLE001
    st.error(f"Не удалось прочитать файл: {exc}")
    st.stop()

if df.empty:
    st.warning("В файле нет данных.")
    st.stop()

columns = [str(col) for col in df.columns]
detected = detect_city_column(columns)
default_index = columns.index(detected) if detected in columns else 0

city_column = st.selectbox(
    "Столбец с городом / населённым пунктом",
    options=columns,
    index=default_index,
    help="По умолчанию выбран столбец, определённый автоматически.",
)

add_original = st.checkbox("Добавить столбец с оригинальным значением", value=True)

threshold = st.slider(
    "Порог нечёткого совпадения",
    min_value=50,
    max_value=100,
    value=DEFAULT_THRESHOLD,
    help="Ниже порог — больше совпадений, выше риск ложных срабатываний.",
)

if st.button("Обработать", type="primary"):
    with st.spinner("Нормализация…"):
        try:
            result = process_file(
                uploaded,
                filename=uploaded.name,
                column=city_column,
                add_original=add_original,
                dry_run=True,
                threshold=threshold,
                df=df,
            )
        except Exception as exc:  # noqa: BLE001
            st.error(f"Ошибка обработки: {exc}")
            st.stop()
    st.session_state["result"] = result

result = st.session_state.get("result")
if result is None:
    st.dataframe(df.head(50), use_container_width=True)
    st.stop()

col_total, col_changed, col_unmatched = st.columns(3)
col_total.metric("Всего строк", result.total)
col_changed.metric("Изменено", result.changed)
col_unmatched.metric("Не сопоставлено", result.unmatched)

st.subheader("Предпросмотр")
st.dataframe(result.df.head(200), use_container_width=True)

with st.expander(f"Несопоставленные строки ({result.unmatched})"):
    if result.unmatched_rows.empty:
        st.success("Все заполненные значения сопоставлены с официальными МО.")
    else:
        st.dataframe(result.unmatched_rows, use_container_width=True)

stem = uploaded.name.rsplit(".", 1)[0]
dl_xlsx, dl_csv = st.columns(2)
with dl_xlsx:
    st.download_button(
        "Скачать xlsx",
        data=to_excel_bytes(result.df),
        file_name=f"{stem}_normalized.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with dl_csv:
    st.download_button(
        "Скачать csv",
        data=to_csv_bytes(result.df),
        file_name=f"{stem}_normalized.csv",
        mime="text/csv",
    )
