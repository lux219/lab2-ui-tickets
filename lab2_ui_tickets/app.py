from __future__ import annotations

import logging
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from results_store import ResultsStore, ResultsWriteError
from student_store import Student, StudentStore
from ticket_parser import Ticket, parse_tickets


APP_DIR = Path(__file__).resolve().parent
STUDENTS_FILE = APP_DIR / "students.xlsx"
TICKETS_FILE = APP_DIR / "tickets.docx"
RESULTS_FILE = APP_DIR / "results.xlsx"
LOG_FILE = APP_DIR / "app.log"

BG_COLOR = "#f3f6fb"
CARD_COLOR = "#ffffff"
TEXT_COLOR = "#172033"
MUTED_COLOR = "#5d6b82"
PRIMARY_COLOR = "#2563eb"
PRIMARY_DARK = "#1d4ed8"
SUCCESS_COLOR = "#0f766e"
WARNING_COLOR = "#b45309"
ERROR_COLOR = "#b91c1c"


class TicketGeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Генератор билетов")
        self.geometry("720x460")
        self.minsize(650, 420)
        self.configure(bg=BG_COLOR)

        self.logger = _configure_logger()
        self.student_store: StudentStore | None = None
        self.results_store = ResultsStore(RESULTS_FILE)
        self.tickets: dict[int, Ticket] = {}
        self.current_students: list[Student] = []

        self.group_var = tk.StringVar()
        self.student_var = tk.StringVar()

        self._configure_styles()
        self._build_ui()
        self._load_sources()

    def _configure_styles(self) -> None:
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("App.TFrame", background=BG_COLOR)
        self.style.configure("Card.TFrame", background=CARD_COLOR)
        self.style.configure("Title.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 22, "bold"))
        self.style.configure("Subtitle.TLabel", background=CARD_COLOR, foreground=MUTED_COLOR, font=("Segoe UI", 10))
        self.style.configure("Field.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 10, "bold"))
        self.style.configure("Status.TLabel", background=CARD_COLOR, foreground=MUTED_COLOR, font=("Segoe UI", 10))
        self.style.configure("Result.TFrame", background=CARD_COLOR)
        self.style.configure("Question.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 10))
        self.style.configure("TCombobox", padding=8, arrowsize=14)
        self.style.configure(
            "Primary.TButton",
            background=PRIMARY_COLOR,
            foreground="#ffffff",
            borderwidth=0,
            focusthickness=0,
            padding=(18, 12),
            font=("Segoe UI", 10, "bold"),
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", PRIMARY_DARK), ("disabled", "#cbd5e1")],
            foreground=[("disabled", "#64748b")],
        )

    def _build_ui(self) -> None:
        shell = ttk.Frame(self, style="App.TFrame", padding=28)
        shell.pack(fill=tk.BOTH, expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(0, weight=1)

        card = ttk.Frame(shell, style="Card.TFrame", padding=30)
        card.grid(row=0, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(3, weight=1)

        ttk.Label(card, text="Генератор билетов", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            card,
            text="Выберите группу и студента. Результат сразу сохраняется в журнал.",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 26))

        form = ttk.Frame(card, style="Card.TFrame")
        form.grid(row=2, column=0, sticky="ew")
        form.columnconfigure(0, weight=1)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Группа", style="Field.TLabel").grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Label(form, text="Студент", style="Field.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))

        self.group_combo = ttk.Combobox(form, textvariable=self.group_var, state="readonly", font=("Segoe UI", 10))
        self.group_combo.grid(row=1, column=0, sticky="ew", padx=(0, 12), pady=(8, 0), ipady=4)
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_selected)

        self.student_combo = ttk.Combobox(form, textvariable=self.student_var, state="readonly", font=("Segoe UI", 10))
        self.student_combo.grid(row=1, column=1, sticky="ew", padx=(12, 0), pady=(8, 0), ipady=4)
        self.student_combo.bind("<<ComboboxSelected>>", self._on_student_selected)

        preview = ttk.Frame(card, style="Card.TFrame")
        preview.grid(row=3, column=0, sticky="nsew", pady=(28, 20))
        preview.columnconfigure(0, weight=1)

        self.status_accent = tk.Frame(preview, bg=PRIMARY_COLOR, height=4)
        self.status_accent.grid(row=0, column=0, sticky="ew", pady=(0, 14))

        self.status_title = ttk.Label(preview, text="Ожидание данных", style="Field.TLabel")
        self.status_title.grid(row=1, column=0, sticky="w")

        self.status_label = ttk.Label(preview, text="", style="Status.TLabel", wraplength=590)
        self.status_label.grid(row=2, column=0, sticky="w", pady=(6, 0))

        self.generate_button = ttk.Button(
            card,
            text="Сгенерировать билет",
            command=self._generate_ticket,
            state=tk.DISABLED,
            style="Primary.TButton",
        )
        self.generate_button.grid(row=4, column=0, sticky="ew")

    def _load_sources(self) -> None:
        missing_files = [str(path.name) for path in (STUDENTS_FILE, TICKETS_FILE) if not path.exists()]
        if missing_files:
            self._set_unavailable(f"Не найдены файлы: {', '.join(missing_files)}.")
            return

        try:
            self.student_store = StudentStore(STUDENTS_FILE)
            self.tickets = parse_tickets(TICKETS_FILE, self.logger)
        except Exception as error:
            self._set_unavailable(str(error))
            return

        if not self.tickets:
            self._set_unavailable("В tickets.docx не найдено ни одного корректного билета.")
            return

        groups = self.student_store.group_names()
        self.group_combo["values"] = groups
        self._set_status("Данные загружены", f"Корректных билетов: {len(self.tickets)}. Выберите группу и студента.", PRIMARY_COLOR)

    def _set_unavailable(self, message: str) -> None:
        self._set_status("Ошибка загрузки", message, ERROR_COLOR)
        self.group_combo.config(state=tk.DISABLED)
        self.student_combo.config(state=tk.DISABLED)
        self.generate_button.config(state=tk.DISABLED)
        messagebox.showerror("Ошибка загрузки", message)

    def _set_status(self, title: str, message: str, color: str = PRIMARY_COLOR) -> None:
        self.status_title.config(text=title)
        self.status_label.config(text=message)
        self.status_accent.config(bg=color)

    def _on_group_selected(self, _event: object | None = None) -> None:
        if self.student_store is None:
            return
        group_name = self.group_var.get()
        self.current_students = self.student_store.students_for_group(group_name)
        self.student_var.set("")
        self.student_combo["values"] = [student.full_name for student in self.current_students]
        self.generate_button.config(state=tk.DISABLED)

        if self.current_students:
            self._set_status(
                "Группа выбрана",
                f"В группе найдено студентов: {len(self.current_students)}. Теперь выберите студента.",
                SUCCESS_COLOR,
            )
        else:
            self._set_status("Пустая группа", "В выбранной группе нет студентов. Кнопка генерации остается выключенной.", WARNING_COLOR)

    def _on_student_selected(self, _event: object | None = None) -> None:
        student = self._selected_student()
        self.generate_button.config(state=tk.NORMAL if student else tk.DISABLED)
        if student:
            self._set_status("Готово к генерации", f"Выбран студент: {student.full_name}.", SUCCESS_COLOR)

    def _selected_student(self) -> Student | None:
        full_name = self.student_var.get()
        for student in self.current_students:
            if student.full_name == full_name:
                return student
        return None

    def _generate_ticket(self) -> None:
        group_name = self.group_var.get()
        student = self._selected_student()
        if not group_name or student is None:
            self.generate_button.config(state=tk.DISABLED)
            return

        while True:
            try:
                assignment = self.results_store.record_assignment(group_name, student, sorted(self.tickets))
                break
            except ResultsWriteError as error:
                retry = messagebox.askretrycancel("Журнал недоступен", f"{error}\n\nЗакройте файл и повторите попытку.")
                if not retry:
                    return

        ticket = self.tickets.get(assignment.ticket_number)
        if ticket is None:
            messagebox.showerror(
                "Ошибка билета",
                f"В журнале указан билет № {assignment.ticket_number}, но его нет среди корректных билетов.",
            )
            return

        self._show_result(ticket, assignment.is_repeat)

    def _show_result(self, ticket: Ticket, is_repeat: bool) -> None:
        window = tk.Toplevel(self)
        window.title(f"Билет № {ticket.number}")
        window.geometry("640x430")
        window.minsize(560, 360)
        window.configure(bg=BG_COLOR)
        window.transient(self)
        window.grab_set()

        shell = ttk.Frame(window, style="App.TFrame", padding=24)
        shell.pack(fill=tk.BOTH, expand=True)

        frame = ttk.Frame(shell, style="Result.TFrame", padding=28)
        frame.pack(fill=tk.BOTH, expand=True)

        badge_text = "Повторная выдача" if is_repeat else "Новая выдача"
        badge_color = WARNING_COLOR if is_repeat else SUCCESS_COLOR
        badge = tk.Label(
            frame,
            text=badge_text,
            bg=badge_color,
            fg="#ffffff",
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=4,
        )
        badge.pack(anchor="w")

        ttk.Label(frame, text=f"Ваш билет № {ticket.number}", style="Title.TLabel").pack(anchor="w", pady=(14, 4))
        if is_repeat:
            ttk.Label(frame, text="Номер взят из первой записи журнала для этого студента.", style="Subtitle.TLabel").pack(
                anchor="w",
                pady=(0, 18),
            )
        else:
            ttk.Label(frame, text="Билет выбран случайно из корректных билетов.", style="Subtitle.TLabel").pack(anchor="w", pady=(0, 18))

        for index, question in enumerate(ticket.questions, start=1):
            question_frame = ttk.Frame(frame, style="Result.TFrame")
            question_frame.pack(fill=tk.X, pady=6)
            number = tk.Label(
                question_frame,
                text=str(index),
                bg=PRIMARY_COLOR,
                fg="#ffffff",
                font=("Segoe UI", 10, "bold"),
                width=3,
                height=1,
            )
            number.pack(side=tk.LEFT, anchor="n", padx=(0, 12))
            ttk.Label(
                question_frame,
                text=question,
                wraplength=500,
                justify=tk.LEFT,
                style="Question.TLabel",
            ).pack(side=tk.LEFT, fill=tk.X, expand=True, anchor="w")

        ttk.Label(frame, text="Esc - закрыть окно результата и вернуться к выбору.", style="Subtitle.TLabel").pack(anchor="w", pady=(22, 0))
        window.bind("<Escape>", lambda _event: window.destroy())
        window.focus_set()


def _configure_logger() -> logging.Logger:
    logger = logging.getLogger("lab2_ui_tickets")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
    return logger


if __name__ == "__main__":
    TicketGeneratorApp().mainloop()
