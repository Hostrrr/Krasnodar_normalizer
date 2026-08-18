"""Веб-интерфейс нормализатора городов Краснодарского края."""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

from city_normalizer import (
    DEFAULT_THRESHOLD,
    detect_city_column,
    delete_user_locality,
    load_user_localities,
    process_file,
    read_file,
    style_preview,
    to_csv_bytes,
    to_excel_bytes,
    unique_unresolved_from_matches,
    upsert_user_locality,
)
from krasnodar_districts import SHORT_DISTRICT_LIST

st.set_page_config(
    page_title="Нормализатор городов Краснодарского края",
    page_icon="🗺️",
    layout="wide",
)


@st.cache_data(show_spinner=False)
def _load_table(data: bytes, name: str) -> pd.DataFrame:
    return read_file(io.BytesIO(data), filename=name)


def _reprocess(uploaded_file, source_df: pd.DataFrame):
    opts = st.session_state.get("process_opts", {})
    st.session_state["result"] = process_file(
        uploaded_file,
        filename=uploaded_file.name,
        column=opts.get("column"),
        add_original=opts.get("add_original", True),
        dry_run=True,
        threshold=opts.get("threshold", DEFAULT_THRESHOLD),
        df=source_df,
    )
    st.session_state["editor_nonce"] = st.session_state.get("editor_nonce", 0) + 1


def _save_place(place: str, district: str, uploaded_file, source_df) -> None:
    saved_name = upsert_user_locality(place, district)
    if uploaded_file is not None and source_df is not None and "result" in st.session_state:
        _reprocess(uploaded_file, source_df)
    st.success(f"Сохранено: «{saved_name}» → {district}. При следующей обработке будет этот район.")
    st.rerun()


st.title("Нормализатор городов Краснодарского края")
st.caption(
    "Приводит столбец «город» к короткому названию района. "
    "Жёлтые строки — несколько районов, красные — пункт не найден: назначьте район ниже."
)

uploaded = st.file_uploader(
    "Загрузите таблицу Excel или CSV",
    type=["xlsx", "xls", "xlsm", "csv"],
)

if uploaded is not None and st.session_state.get("uploaded_name") != uploaded.name:
    st.session_state.pop("result", None)
    st.session_state["uploaded_name"] = uploaded.name

df = None
if uploaded is None:
    st.info("Выберите файл, чтобы начать. Поддерживаются xlsx, xls, xlsm и csv.")
else:
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
    )
    add_original = st.checkbox("Добавить столбец с оригинальным значением", value=True)
    with st.expander("Дополнительно: насколько строго искать похожие названия"):
        st.markdown(
            "Порог — насколько название может отличаться от справочника "
            "(опечатки вроде «Соччи» → Сочи). Обычно **82**, трогать не нужно."
        )
        threshold = st.slider("Порог похожести", 50, 100, DEFAULT_THRESHOLD)

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
        st.session_state["process_opts"] = {
            "column": city_column,
            "add_original": add_original,
            "threshold": threshold,
        }
        st.session_state["editor_nonce"] = st.session_state.get("editor_nonce", 0) + 1

    result = st.session_state.get("result")
    if result is None:
        st.dataframe(df.head(50), use_container_width=True)
    else:
        col_total, col_certain, col_uncertain, col_unmatched = st.columns(4)
        col_total.metric("Всего строк", result.total)
        col_certain.metric("Уверен", result.certain)
        col_uncertain.metric("Не уверен", result.uncertain)
        col_unmatched.metric("Не сопоставлено", result.unmatched)

        unresolved = unique_unresolved_from_matches(result.originals, result.matches)

        st.divider()
        st.subheader("1. Назначить район пункту, которого нет в справочнике")
        st.markdown(
            "Красная строка = программа **не знает** этот хутор/станицу. "
            "Выберите его в списке, укажите район **из всех 44** и нажмите сохранение. "
            "Так вы добавляете пункт в справочник."
        )

        if unresolved.empty:
            st.success("Все названия из файла сопоставлены.")
        else:
            labels = []
            for _, row in unresolved.iterrows():
                extra = f", варианты: {row['варианты']}" if str(row["варианты"]).strip() else ""
                labels.append(
                    f"{row['населённый пункт']}  —  {row['статус']}  ({row['строк']} строк{extra})"
                )
            chosen = st.selectbox(
                "Какой населённый пункт уточнить",
                options=list(range(len(unresolved))),
                format_func=lambda i: labels[i],
            )
            chosen_row = unresolved.iloc[chosen]
            place_name = str(chosen_row["населённый пункт"])
            variants_raw = str(chosen_row["варианты"]).strip()
            variant_list = [v.strip() for v in variants_raw.split("/") if v.strip()] if variants_raw else []

            if variant_list:
                st.info("Программа предлагает такие районы: **" + ", ".join(variant_list) + "**. Можно выбрать любой из 44.")
            else:
                st.warning(
                    f"«{place_name}» в справочнике нет. Выберите район вручную — после сохранения программа будет его знать."
                )

            default_district = variant_list[0] if variant_list and variant_list[0] in SHORT_DISTRICT_LIST else SHORT_DISTRICT_LIST[0]
            district = st.selectbox(
                "Район",
                options=SHORT_DISTRICT_LIST,
                index=SHORT_DISTRICT_LIST.index(default_district),
                key="assign_district",
            )
            if st.button("Сохранить этот пункт в справочник", type="primary"):
                try:
                    _save_place(place_name, district, uploaded, df)
                except ValueError as exc:
                    st.error(str(exc))

        st.markdown("**Пункта нет даже в списке выше? Впишите название сами:**")
        man_col1, man_col2, man_col3 = st.columns([2, 2, 1])
        with man_col1:
            typed_place = st.text_input("Название станицы / хутора / села", placeholder="например Цыпка")
        with man_col2:
            typed_district = st.selectbox("Район для нового пункта", options=SHORT_DISTRICT_LIST, key="typed_district")
        with man_col3:
            st.write("")
            st.write("")
            typed_save = st.button("Добавить", key="typed_save")
        if typed_save:
            try:
                _save_place(typed_place, typed_district, uploaded, df)
            except ValueError as exc:
                st.error(str(exc))

        st.divider()
        st.subheader("2. Предпросмотр таблицы")
        st.caption("Жёлтым сверху — не уверена (несколько районов). Красным — не найдено.")
        st.dataframe(style_preview(result.df), use_container_width=True, height=480)

        st.subheader("3. Скачать результат")
        stem = uploaded.name.rsplit(".", 1)[0]
        dl_xlsx, dl_csv = st.columns(2)
        try:
            xlsx_data = to_excel_bytes(result.df)
        except Exception as exc:  # noqa: BLE001
            xlsx_data = None
            st.error(f"Не удалось собрать Excel: {exc}")
        with dl_xlsx:
            if xlsx_data is not None:
                st.download_button(
                    "Скачать xlsx",
                    data=xlsx_data,
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

st.divider()
st.subheader("Мой справочник")
st.caption("Любой неизвестный пункт можно добавить здесь — даже без файла.")
ref_col1, ref_col2, ref_col3 = st.columns([2, 2, 1])
with ref_col1:
    dict_place = st.text_input("Новый населённый пункт", key="dict_place")
with ref_col2:
    dict_district = st.selectbox("Его район", options=SHORT_DISTRICT_LIST, key="dict_district")
with ref_col3:
    st.write("")
    st.write("")
    dict_save = st.button("Добавить в справочник", key="dict_save")
if dict_save:
    try:
        _save_place(dict_place, dict_district, uploaded, df)
    except ValueError as exc:
        st.error(str(exc))

user_map = load_user_localities()
if not user_map:
    st.caption("Пока пусто. Когда сохраните пункт — он появится здесь.")
else:
    user_df = pd.DataFrame(
        [
            {"населённый пункт": name, "район": district}
            for name, district in sorted(user_map.items())
        ]
    )
    st.dataframe(user_df, hide_index=True, use_container_width=True)
    to_delete = st.selectbox("Удалить запись", options=["—"] + sorted(user_map.keys()))
    if to_delete != "—" and st.button("Удалить из справочника"):
        delete_user_locality(to_delete)
        if uploaded is not None and df is not None and "result" in st.session_state:
            _reprocess(uploaded, df)
        st.rerun()
