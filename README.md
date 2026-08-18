# Krasnodar_normalizer

Обработчик Excel/CSV: приводит столбец «город» к официальным названиям муниципальных образований Краснодарского края.

## Запуск

Windows: распакуйте ZIP в обычную папку и дважды щёлкните `run.bat`
(установит зависимости при первом запуске и откроет браузер).

Если окно сразу вспыхивает и закрывается — запускайте `run.bat` не из архива,
а из распакованной папки. Если Python не установлен, скачайте его с
https://www.python.org/downloads/ и отметьте **Add python.exe to PATH**.

Или вручную:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Неоднозначные и неизвестные строки поднимаются вверх и выделяются цветом.
На странице можно назначить район станице или хутору и сохранить его в справочник.

## Командная строка

```bash
python city_normalizer.py input.xlsx --output result.xlsx --add-original
python city_normalizer.py input.csv --column Город --dry-run --threshold 82
```
