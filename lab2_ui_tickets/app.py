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


class TicketGeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Генератор билетов")
        self.geometry("520x260")
        self.minsize(520, 260)

        self.logger = _configure_logger()
        self.student_store: StudentStore | None = None
        self.results_store = ResultsStore(RESULTS_FILE)
        self.tickets: dict[int, Ticket] = {}
        self.current_students: list[Student] = []

        self.group_var = tk.StringVar()
        self.student_var = tk.StringVar()

        self._build_ui()
        self._load_sources()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self, padding=24)
        frame.pack(fill=tk.BOTH, expand=True)
        frame.columnconfigure(1, weight=1)

        title = ttk.Label(frame, text="Генератор билетов", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 18))

        ttk.Label(frame, text="Группа").grid(row=1, column=0, sticky="w", padx=(0, 14), pady=8)
        self.group_combo = ttk.Combobox(frame, textvariable=self.group_var, state="readonly")
        self.group_combo.grid(row=1, column=1, sticky="ew", pady=8)
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_selected)

        ttk.Label(frame, text="Студент").grid(row=2, column=0, sticky="w", padx=(0, 14), pady=8)
        self.student_combo = ttk.Combobox(frame, textvariable=self.student_var, state="readonly")
        self.student_combo.grid(row=2, column=1, sticky="ew", pady=8)
        self.student_combo.bind("<<ComboboxSelected>>", self._on_student_selected)

        self.generate_button = ttk.Button(
            frame,
            text="Сгенерировать билет",
            command=self._generate_ticket,
            state=tk.DISABLED,
        )
        self.generate_button.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(22, 8))

        self.status_label = ttk.Label(frame, text="", foreground="#555555", wraplength=450)
        self.status_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

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
        self.status_label.config(text=f"Загружено билетов: {len(self.tickets)}. Выберите группу и студента.")

    def _set_unavailable(self, message: str) -> None:
        self.status_label.config(text=message)
        self.group_combo.config(state=tk.DISABLED)
        self.student_combo.config(state=tk.DISABLED)
        self.generate_button.config(state=tk.DISABLED)
        messagebox.showerror("Ошибка загрузки", message)

    def _on_group_selected(self, _event: object | None = None) -> None:
        if self.student_store is None:
            return
        group_name = self.group_var.get()
        self.current_students = self.student_store.students_for_group(group_name)
        self.student_var.set("")
        self.student_combo["values"] = [student.full_name for student in self.current_students]
        self.generate_button.config(state=tk.DISABLED)

        if self.current_students:
            self.status_label.config(text="Выберите студента.")
        else:
            self.status_label.config(text="В выбранной группе нет студентов.")

    def _on_student_selected(self, _event: object | None = None) -> None:
        self.generate_button.config(state=tk.NORMAL if self._selected_student() else tk.DISABLED)

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
        window.geometry("560x320")
        window.minsize(480, 260)
        window.transient(self)
        window.grab_set()

        frame = ttk.Frame(window, padding=24)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text=f"Ваш билет № {ticket.number}", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        if is_repeat:
            ttk.Label(frame, text="Повторная выдача: номер взят из первой записи журнала.").pack(
                anchor="w",
                pady=(4, 12),
            )
        else:
            ttk.Label(frame, text="Новая выдача билета.").pack(anchor="w", pady=(4, 12))

        for index, question in enumerate(ticket.questions, start=1):
            ttk.Label(frame, text=f"{index}. {question}", wraplength=500, justify=tk.LEFT).pack(anchor="w", pady=4)

        ttk.Label(frame, text="Нажмите Esc, чтобы вернуться к выбору.").pack(anchor="w", pady=(18, 0))
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
