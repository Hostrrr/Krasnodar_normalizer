"""Нормализация столбца «город» к официальным МО Краснодарского края."""

from __future__ import annotations

import argparse
import io
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import BinaryIO, Iterable

import pandas as pd
from rapidfuzz import fuzz, process

from krasnodar_districts import LOCALITY_TO_DISTRICT, OFFICIAL_DISTRICTS

DEFAULT_THRESHOLD = 82
SUPPORTED_EXCEL = {".xlsx", ".xls", ".xlsm"}
SUPPORTED_CSV = {".csv"}
CSV_ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "windows-1251")

COLUMN_ALIASES = (
    "населённый пункт",
    "населенный пункт",
    "город",
    "city",
    "адрес",
    "address",
    "местность",
    "район",
    "нп",
)

_OFFICIAL_FOLD = {
    name.lower().replace("ё", "е"): name for name in OFFICIAL_DISTRICTS
}
_LOCALITY_KEYS = list(LOCALITY_TO_DISTRICT.keys())

# Длинные составные префиксы — первыми, короткие аббревиатуры — в конце.
_PREFIX_RE = re.compile(
    r"""^
    (?:
        муниципальный\s+округ\s+город(?:-курорт|-герой)?
        |городской\s+округ\s+город(?:-курорт|-герой)?
        |муниципальный\s+округ
        |городской\s+округ
        |муниципальный\s+район
        |город-курорт
        |город-герой
        |г\.-к\.?
        |ст\.-ца\.?
        |р\.\s*п\.
        |к/п
        |г/о
        |м\.р\.
        |пгт\.?
        |пос[её]лок
        |станица
        |деревня
        |хутор
        |село
        |аул
        |пос\.
        |рп\.?
        |кп\.?
        |город
        |г\.
        |п\.
        |с\.
        |ст\.
        |х\.
        |а\.
        |д\.
    )
    [\s.:\-]*
    """,
    re.IGNORECASE | re.VERBOSE,
)


def _fold(text: str) -> str:
    return " ".join(text.lower().replace("ё", "е").split())


def _is_empty(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def strip_prefix(value: str) -> str:
    """Убирает административные префиксы (г., ст., пгт, городской округ …)."""
    text = str(value).strip().strip("«»\"'`")
    if not text:
        return ""
    previous = None
    while previous != text:
        previous = text
        text = _PREFIX_RE.sub("", text, count=1).strip(" .,;-")
    return " ".join(text.split())


def fuzzy_lookup(name: str, threshold: int = DEFAULT_THRESHOLD) -> str | None:
    """Нечёткий поиск среди ключей LOCALITY_TO_DISTRICT. Возвращает официальное МО."""
    query = _fold(name)
    if not query:
        return None
    hit = process.extractOne(
        query,
        _LOCALITY_KEYS,
        scorer=fuzz.WRatio,
        score_cutoff=threshold,
    )
    if hit is None:
        return None
    return LOCALITY_TO_DISTRICT[hit[0]]


def normalize_city(value: object, threshold: int = DEFAULT_THRESHOLD) -> object:
    """Приводит значение города к официальному названию муниципального образования.

    Если сопоставить не удалось, возвращает очищенное название (данные не теряются).
    """
    if _is_empty(value):
        return value

    original = str(value).strip()
    if original.lower() in {"nan", "none", "null", "-"}:
        return original

    folded_raw = _fold(original)
    if folded_raw in _OFFICIAL_FOLD:
        return _OFFICIAL_FOLD[folded_raw]
    if folded_raw in LOCALITY_TO_DISTRICT:
        return LOCALITY_TO_DISTRICT[folded_raw]

    cleaned = strip_prefix(original)
    folded = _fold(cleaned)
    if not folded:
        return cleaned or original

    if folded in _OFFICIAL_FOLD:
        return _OFFICIAL_FOLD[folded]
    if folded in LOCALITY_TO_DISTRICT:
        return LOCALITY_TO_DISTRICT[folded]

    fuzzy = fuzzy_lookup(folded, threshold=threshold)
    if fuzzy:
        return fuzzy

    return cleaned


def detect_city_column(columns: Iterable[object]) -> str | None:
    """Автоопределение столбца с городом/населённым пунктом."""
    names = [str(col) for col in columns]
    folded = [_fold(name) for name in names]

    for alias in COLUMN_ALIASES:
        alias_fold = _fold(alias)
        for original, current in zip(names, folded):
            if current == alias_fold:
                return original

    for alias in COLUMN_ALIASES:
        alias_fold = _fold(alias)
        for original, current in zip(names, folded):
            if alias_fold in current.split() or alias_fold == current:
                return original
            if alias_fold in current and alias_fold != "нп":
                return original
    return names[0] if names else None


def _suffix(path: str | Path) -> str:
    return Path(path).suffix.lower()


def read_file(
    source: str | Path | BinaryIO,
    filename: str | None = None,
) -> pd.DataFrame:
    """Читает xlsx / xls / xlsm / csv (utf-8, utf-8-sig, windows-1251)."""
    name = filename
    if name is None and isinstance(source, (str, Path)):
        name = str(source)
    if not name:
        raise ValueError("Не удалось определить имя файла.")

    ext = _suffix(name)
    if ext in SUPPORTED_EXCEL:
        return _read_excel(source, ext)
    if ext in SUPPORTED_CSV:
        return _read_csv(source)
    raise ValueError(
        f"Неподдерживаемый формат «{ext}». Допустимы: xlsx, xls, xlsm, csv."
    )


def _as_buffer(source: str | Path | BinaryIO) -> tuple[bool, object]:
    if isinstance(source, (str, Path)):
        return True, source
    if hasattr(source, "seek"):
        source.seek(0)
    return False, source


def _read_excel(source: str | Path | BinaryIO, ext: str) -> pd.DataFrame:
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    _, handle = _as_buffer(source)
    return pd.read_excel(handle, engine=engine)


def _read_csv(source: str | Path | BinaryIO) -> pd.DataFrame:
    is_path, handle = _as_buffer(source)
    raw: bytes
    if is_path:
        raw = Path(handle).read_bytes()
    else:
        raw = handle.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8")

    last_error: Exception | None = None
    for encoding in CSV_ENCODINGS:
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(
                    io.BytesIO(raw),
                    encoding=encoding,
                    sep=sep,
                    dtype=str,
                    keep_default_na=False,
                    na_values=["", "NA", "NaN", "nan", "None"],
                )
            except (UnicodeDecodeError, pd.errors.ParserError) as exc:
                last_error = exc
                continue
            if df.shape[1] == 1 and sep == "," and ";" in raw.decode(
                encoding, errors="ignore"
            ):
                continue
            return df
    raise ValueError(
        "Не удалось прочитать CSV. Проверьте кодировку и разделитель."
    ) from last_error


def write_file(df: pd.DataFrame, path: str | Path) -> Path:
    """Сохраняет таблицу в xlsx / xlsm / csv. Для .xls пишет .xlsx."""
    out = Path(path)
    ext = out.suffix.lower()
    out.parent.mkdir(parents=True, exist_ok=True)

    if ext == ".csv":
        df.to_csv(out, index=False, encoding="utf-8-sig")
        return out
    if ext in {".xlsx", ".xlsm"}:
        df.to_excel(out, index=False, engine="openpyxl")
        return out
    if ext == ".xls":
        out = out.with_suffix(".xlsx")
        df.to_excel(out, index=False, engine="openpyxl")
        return out
    raise ValueError(
        f"Неподдерживаемый формат записи «{ext}». Допустимы: xlsx, xlsm, csv."
    )


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    return buffer.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


@dataclass
class ProcessResult:
    df: pd.DataFrame
    column: str
    total: int
    changed: int
    unmatched: int
    exact: int
    fuzzy: int
    unmatched_rows: pd.DataFrame = field(default_factory=pd.DataFrame)
    output_path: Path | None = None


def process_file(
    source: str | Path | BinaryIO,
    *,
    filename: str | None = None,
    output: str | Path | None = None,
    column: str | None = None,
    add_original: bool = False,
    dry_run: bool = False,
    threshold: int = DEFAULT_THRESHOLD,
    df: pd.DataFrame | None = None,
) -> ProcessResult:
    """Полный цикл: чтение → нормализация столбца → статистика → запись."""
    table = df.copy() if df is not None else read_file(source, filename=filename)
    if table.empty:
        raise ValueError("Файл не содержит строк.")

    city_col = column or detect_city_column(table.columns)
    if not city_col or city_col not in table.columns:
        raise ValueError(
            "Не удалось определить столбец с городом. "
            "Укажите его явно через --column / выбор в интерфейсе."
        )

    originals = table[city_col]
    normalized: list[object] = []
    exact = 0
    fuzzy = 0
    changed = 0
    unmatched_mask: list[bool] = []

    for value in originals:
        if _is_empty(value):
            normalized.append(value)
            unmatched_mask.append(False)
            continue

        cleaned = strip_prefix(str(value))
        folded = _fold(cleaned)
        folded_raw = _fold(str(value))
        result = normalize_city(value, threshold=threshold)
        normalized.append(result)

        original_text = str(value).strip()
        result_text = "" if _is_empty(result) else str(result)
        if result_text != original_text:
            changed += 1

        matched = result_text in OFFICIAL_DISTRICTS
        unmatched_mask.append(not matched)

        if matched:
            is_exact = (
                folded in LOCALITY_TO_DISTRICT
                or folded in _OFFICIAL_FOLD
                or folded_raw in LOCALITY_TO_DISTRICT
                or folded_raw in _OFFICIAL_FOLD
            )
            if is_exact:
                exact += 1
            else:
                fuzzy += 1

    result_df = table.copy()
    if add_original:
        orig_name = f"{city_col}_оригинал"
        insert_at = list(result_df.columns).index(city_col) + 1
        result_df.insert(insert_at, orig_name, originals.to_list())
    result_df[city_col] = normalized

    unmatched_rows = result_df.loc[unmatched_mask].copy()
    total = len(result_df)

    written: Path | None = None
    if not dry_run and output is not None:
        written = write_file(result_df, output)

    return ProcessResult(
        df=result_df,
        column=city_col,
        total=total,
        changed=changed,
        unmatched=int(sum(unmatched_mask)),
        exact=exact,
        fuzzy=fuzzy,
        unmatched_rows=unmatched_rows,
        output_path=written,
    )


def _default_output(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_normalized{input_path.suffix}")


def _print_stats(stats: ProcessResult, output: Path | None, dry_run: bool) -> None:
    print(f"Столбец: {stats.column}")
    print(f"Всего строк: {stats.total}")
    print(f"Изменено: {stats.changed}")
    print(f"Точных совпадений: {stats.exact}")
    print(f"Нечётких совпадений: {stats.fuzzy}")
    print(f"Не сопоставлено: {stats.unmatched}")
    if dry_run:
        print("Режим --dry-run: файл не записан.")
    elif output is not None:
        print(f"Записано: {output}")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Нормализация столбца «город» к официальным названиям "
            "муниципальных образований Краснодарского края."
        )
    )
    parser.add_argument("input", help="Входной файл: xlsx, xls, xlsm или csv")
    parser.add_argument(
        "--output",
        "-o",
        help="Путь для сохранения. По умолчанию: <имя>_normalized.<расширение>",
    )
    parser.add_argument(
        "--column",
        "-c",
        help="Имя столбца с городом (если автоопределение не сработало)",
    )
    parser.add_argument(
        "--add-original",
        action="store_true",
        help="Добавить столбец с исходным значением",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Только статистика, без записи файла",
    )
    parser.add_argument(
        "--threshold",
        type=int,
        default=DEFAULT_THRESHOLD,
        help=f"Порог нечёткого совпадения 0–100 (по умолчанию {DEFAULT_THRESHOLD})",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    if not 0 <= args.threshold <= 100:
        parser.error("--threshold должен быть в диапазоне 0–100")

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Файл не найден: {input_path}", file=sys.stderr)
        return 1

    output_path = None if args.dry_run else Path(args.output) if args.output else _default_output(input_path)

    try:
        stats = process_file(
            input_path,
            output=output_path,
            column=args.column,
            add_original=args.add_original,
            dry_run=args.dry_run,
            threshold=args.threshold,
        )
    except Exception as exc:  # noqa: BLE001 — пользовательский CLI
        print(f"Ошибка: {exc}", file=sys.stderr)
        return 1

    _print_stats(stats, output_path, args.dry_run)
    return 0


if __name__ == "__main__":
    sys.exit(main())
