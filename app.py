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


st.title("Нормализатор городов Краснодарского края")
st.caption(
    "Приводит столбец «город» к короткому названию района. "
    "Неоднозначные строки — сверху, жёлтым; неизвестные — красным."
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
        help="По умолчанию выбран столбец, определённый автоматически.",
    )
    add_original = st.checkbox("Добавить столбец с оригинальным значением", value=True)
    with st.expander("Дополнительно: насколько строго искать похожие названия"):
        st.markdown(
            "Порог — насколько название в таблице может **отличаться** от справочника "
            "(опечатки вроде «Соччи» → Сочи).\n\n"
            "- **82** — обычный режим, так и оставляйте.\n"
            "- **Ниже** (70–80) — программа угадывает смелее, больше ошибок.\n"
            "- **Выше** (90–100) — только почти точные совпадения."
        )
        threshold = st.slider(
            "Порог похожести",
            min_value=50,
            max_value=100,
            value=DEFAULT_THRESHOLD,
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

        st.caption(
            "Жёлтым сверху — несколько районов или нечёткое совпадение. "
            "Красным — населённый пункт не найден в справочнике."
        )
        st.subheader("Предпросмотр")
        st.dataframe(style_preview(result.df), use_container_width=True, height=520)

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

        unresolved = unique_unresolved_from_matches(result.originals, result.matches)
        st.divider()
        st.subheader("Уточнить станицы и населённые пункты")
        st.caption(
            "Выберите район для неизвестных или неоднозначных пунктов. "
            "Запись сохранится в справочнике и будет использоваться дальше."
        )
        if unresolved.empty:
            st.success("Неизвестных и неоднозначных названий в этом файле нет.")
        else:
            edited = st.data_editor(
                unresolved,
                column_config={
                    "населённый пункт": st.column_config.TextColumn(disabled=True),
                    "статус": st.column_config.TextColumn(disabled=True),
                    "варианты": st.column_config.TextColumn(
                        "Возможные районы",
                        disabled=True,
                    ),
                    "район": st.column_config.SelectboxColumn(
                        "Назначить район",
                        options=["не выбран"] + SHORT_DISTRICT_LIST,
                    ),
                    "строк": st.column_config.NumberColumn(disabled=True),
                },
                hide_index=True,
                use_container_width=True,
                key=f"unresolved_editor_{st.session_state.get('editor_nonce', 0)}",
            )
            if st.button("Сохранить выбранные районы", type="primary"):
                saved = 0
                for _, row in edited.iterrows():
                    district = str(row.get("район", "")).strip()
                    place = str(row.get("населённый пункт", "")).strip()
                    if place and district and district not in {"", "не выбран", "nan"}:
                        upsert_user_locality(place, district)
                        saved += 1
                if saved:
                    _reprocess(uploaded, df)
                    st.success(f"Сохранено записей: {saved}. Таблица пересчитана.")
                    st.rerun()
                else:
                    st.warning("Выберите район хотя бы для одной строки.")

st.divider()
st.subheader("Справочник населённых пунктов")
st.caption("Сюда можно добавить станицу, хутор или село, которого нет в программе.")

with st.form("add_place_form", clear_on_submit=True):
    add_col1, add_col2 = st.columns(2)
    with add_col1:
        new_place = st.text_input("Название (станица, хутор, село, город)")
    with add_col2:
        new_district = st.selectbox("Район", options=SHORT_DISTRICT_LIST)
    submitted = st.form_submit_button("Добавить в справочник")
    if submitted:
        try:
            saved_name = upsert_user_locality(new_place, new_district)
        except ValueError as exc:
            st.error(str(exc))
        else:
            if uploaded is not None and df is not None and "result" in st.session_state:
                _reprocess(uploaded, df)
            st.success(f"«{saved_name}» → {new_district}")
            st.rerun()

user_map = load_user_localities()
if not user_map:
    st.caption("Пока пусто — добавленные пункты появятся в этом списке.")
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
