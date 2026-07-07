# FillPercent

Python/Photoshop utility для оценки заполненности PDF-макетов и изображений. Проект сочетает локальную обработку PDF, экспорт результатов в Excel и JSX-автоматизацию Adobe Photoshop.

## Статус аудита

`fillpercent.py` — актуальная версия. Старые варианты `fillpercent1.py`, `fillpercent2.py`, `fillpercent3s.py`, `fillpercent4s.py` удалены из текущей ветки.

Проект пока не готов к public release: в текущем `fillpercent.py` и `blackfill.jsx` остаются absolute local paths к Photoshop, Poppler и рабочим папкам. Это не credential, но это внутренняя локальная конфигурация. Перед публикацией нужно вынести пути в `.env`, JSON/YAML config или GUI settings и добавить safe example config.

## Возможности

- выбор PDF-файлов или папки через GUI;
- запуск Photoshop JSX preprocessing;
- извлечение страниц PDF через `pdf2image`;
- расчёт процента заполнения по тёмным пикселям;
- определение размера изображения в миллиметрах;
- сохранение результатов в `fill_results.xlsx`;
- очистка имени файла от служебных слов.

## Стек

- Python;
- Tkinter;
- pdf2image / Poppler;
- Pillow;
- NumPy;
- openpyxl;
- Adobe Photoshop JSX.

## Запуск

```bash
python fillpercent.py
```

## Перед публикацией

1. Вынести `PHOTOSHOP_EXE`, `SCRIPT_PATH`, `OUTPUT_FOLDER`, `POPPER_PATH` из кода.
2. Заменить absolute paths в `blackfill.jsx` на аргументы/config.
3. Добавить `.gitignore` для generated Excel/log/Prints/raw outputs.
4. Выполнить full-history secret/data/path scan.
5. Добавить synthetic sample PDF и screenshot/GIF.

## License recommendation

MIT для собственного Python/JSX-кода после cleanup. Adobe Photoshop и Poppler используются как внешние зависимости и лицензируются отдельно.
