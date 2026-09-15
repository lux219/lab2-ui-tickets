# Prompt for Copilot Chat

Build a desktop GUI application for Lab 2: a ticket generator for students.

Use Python with Tkinter. Keep the GUI layer separate from the business logic so the parser, student reader, and result journal can be covered by unit tests.

Input files:

- `students.xlsx`: one worksheet per group. The worksheet name is the group name. Row 1 contains headers. Student data starts at row 2. Column A is last name, column B is first name.
- `tickets.docx`: a Word document with tickets in this format:

```text
Билет 1
1. First question
2. Second question
3. Third question

Билет 2
...
```

Ticket headers are paragraphs matching `Билет N`. Questions are numbered paragraphs matching `N. text` or `N) text`.

Output file:

- `results.xlsx`: create it automatically if it does not exist. Append one row immediately after each generation. Never overwrite or edit existing rows.

`results.xlsx` columns:

1. Группа
2. Фамилия
3. Имя
4. Номер билета
5. Дата и время
6. Повтор (да/нет)

GUI behavior:

- On startup, load group names from worksheet names in `students.xlsx`.
- The first dropdown contains group names.
- The second dropdown is empty until a group is selected.
- After selecting a group, fill the student dropdown with students from that worksheet in the format `LastName FirstName`.
- If the selected group has no students, keep the student dropdown empty and keep the button disabled.
- The “Сгенерировать билет” button must be disabled until both a group and a student are selected.
- Clicking the button opens a result window showing `Ваш билет № N` and exactly three questions.
- Pressing Esc closes only the result window and returns to the selection window. The app must remain open.

Ticket assignment rules:

- If the same student, identified by group + last name + first name, already exists in `results.xlsx`, use the ticket number from that student's first journal row. Append a new row with the same ticket number and `Повтор (да/нет)` = `да`.
- If the student is not in the journal, choose a random valid ticket from the parsed DOCX tickets. Append a row with `Повтор (да/нет)` = `нет`.
- Save immediately after generation, before returning control to the user.

Validation and error handling:

- If `students.xlsx` or `tickets.docx` is missing on startup, show a clear error message and do not crash.
- If `results.xlsx` is open in Excel or cannot be written, show a clear message and allow the user to retry after closing the file. Do not crash.
- If a ticket has fewer than three questions or has an invalid header, skip that ticket, log the problem, and continue with all remaining valid tickets.
- If no valid tickets remain, show a clear error and keep generation disabled.

Tests:

- Add unit tests for parsing valid and invalid tickets from Word.
- Add unit tests for reading students by group from Excel, including an empty group.
- Add unit tests proving idempotency: the same student receives the same ticket on repeated generation, while a new row is appended and marked as repeat.

Acceptance criteria:

- Group selection fills the student list from the correct worksheet.
- The generate button is enabled only when both selections are present.
- The result window shows the selected/generated ticket number and three questions.
- Re-selecting the same student returns the first assigned ticket number.
- Esc closes the result window only.
- `results.xlsx` receives a new row immediately after each generation.
- Existing journal rows are not modified.
- Missing input files and locked result files do not crash the app.
- The journal preserves the full chronological history of all generations.
