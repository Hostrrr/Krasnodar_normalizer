# Krasnodar_normalizer

Обработчик Excel/CSV: приводит столбец «город» к официальным названиям муниципальных образований Краснодарского края.

## Запуск

Windows: двойной клик по `run.bat` (установит зависимости при первом запуске и откроет веб-интерфейс).

Или вручную:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Командная строка

```bash
python city_normalizer.py input.xlsx --output result.xlsx --add-original
python city_normalizer.py input.csv --column Город --dry-run --threshold 82
```
