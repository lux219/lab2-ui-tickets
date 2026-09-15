# Лабораторная работа 2

Десктопное приложение с графическим интерфейсом для генерации билетов студентам.

## Стек

- Python 3
- Tkinter
- openpyxl
- python-docx
- pytest

## Файлы

- `app.py` - запуск GUI.
- `student_store.py` - чтение групп и студентов из `students.xlsx`.
- `ticket_parser.py` - парсинг билетов из `tickets.docx`.
- `results_store.py` - создание и пополнение `results.xlsx`.
- `sample_data.py` - создание демонстрационных `students.xlsx` и `tickets.docx`.
- `copilot_prompt_en.md` - детальное ТЗ на английском для Copilot Chat.
- `tests/` - unit-тесты.

## Запуск

```bash
pip install -r requirements.txt
python sample_data.py
python app.py
```

Приложение ищет `students.xlsx` и `tickets.docx` в этой же папке. `results.xlsx` создается автоматически при первой генерации.

## Тесты

```bash
pytest
```

## Формат входных файлов

`students.xlsx`: один лист на группу, имя листа равно имени группы. В строке 1 заголовки, данные с строки 2:

- колонка A - фамилия;
- колонка B - имя.

`tickets.docx`: билеты в формате:

```text
Билет 1
1. Вопрос первый
2. Вопрос второй
3. Вопрос третий
```

Билет пропускается, если заголовок не распознан или найдено меньше трех вопросов.
