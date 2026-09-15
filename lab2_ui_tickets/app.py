from __future__ import annotations

import logging
import os
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

BG_COLOR = "#eef3fb"
CARD_COLOR = "#ffffff"
SIDEBAR_COLOR = "#101827"
SIDEBAR_MUTED = "#8ea0bd"
TEXT_COLOR = "#172033"
MUTED_COLOR = "#667085"
LINE_COLOR = "#d9e2f1"
PRIMARY_COLOR = "#2563eb"
PRIMARY_DARK = "#1d4ed8"
SUCCESS_COLOR = "#0f766e"
WARNING_COLOR = "#b45309"
ERROR_COLOR = "#b91c1c"


class TicketGeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Exam Ticket Service")
        self.geometry("1180x720")
        self.minsize(1050, 640)
        self.configure(bg=BG_COLOR)

        self.logger = _configure_logger()
        self.student_store: StudentStore | None = None
        self.results_store = ResultsStore(RESULTS_FILE)
        self.tickets: dict[int, Ticket] = {}
        self.current_students: list[Student] = []
        self.nav_buttons: dict[str, tk.Label] = {}
        self.active_view = "Dashboard"

        self.group_var = tk.StringVar()
        self.student_var = tk.StringVar()

        self._configure_styles()
        self._build_ui()
        self._load_sources()

    def _configure_styles(self) -> None:
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("App.TFrame", background=BG_COLOR)
        self.style.configure("Sidebar.TFrame", background=SIDEBAR_COLOR)
        self.style.configure("Card.TFrame", background=CARD_COLOR)
        self.style.configure("Title.TLabel", background=BG_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 24, "bold"))
        self.style.configure("CardTitle.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 13, "bold"))
        self.style.configure("Subtitle.TLabel", background=BG_COLOR, foreground=MUTED_COLOR, font=("Segoe UI", 10))
        self.style.configure("Muted.TLabel", background=CARD_COLOR, foreground=MUTED_COLOR, font=("Segoe UI", 9))
        self.style.configure("Field.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 10, "bold"))
        self.style.configure("Metric.TLabel", background=CARD_COLOR, foreground=TEXT_COLOR, font=("Segoe UI", 20, "bold"))
        self.style.configure("MetricName.TLabel", background=CARD_COLOR, foreground=MUTED_COLOR, font=("Segoe UI", 9))
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
        self.style.configure(
            "Ghost.TButton",
            background="#edf3ff",
            foreground=PRIMARY_COLOR,
            borderwidth=0,
            focusthickness=0,
            padding=(14, 10),
            font=("Segoe UI", 10, "bold"),
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", PRIMARY_DARK), ("disabled", "#cbd5e1")],
            foreground=[("disabled", "#64748b")],
        )
        self.style.map("Ghost.TButton", background=[("active", "#dbe8ff")])
        self.style.configure(
            "History.Treeview",
            background=CARD_COLOR,
            foreground=TEXT_COLOR,
            fieldbackground=CARD_COLOR,
            rowheight=30,
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "History.Treeview.Heading",
            background="#f4f7fc",
            foreground=MUTED_COLOR,
            relief="flat",
            font=("Segoe UI", 9, "bold"),
        )

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="App.TFrame")
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)

        self._build_sidebar(root)

        self.main = ttk.Frame(root, style="App.TFrame", padding=(28, 22, 28, 22))
        self.main.grid(row=0, column=1, sticky="nsew")
        self.main.columnconfigure(0, weight=1)
        self.main.rowconfigure(3, weight=1)

        self._render_active_view()

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        sidebar = ttk.Frame(parent, style="Sidebar.TFrame", padding=(22, 24))
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.configure(width=235)
        sidebar.grid_propagate(False)

        tk.Label(sidebar, text="ExamFlow", bg=SIDEBAR_COLOR, fg="#ffffff", font=("Segoe UI", 19, "bold")).pack(anchor="w")
        tk.Label(
            sidebar,
            text="Ticket generation panel",
            bg=SIDEBAR_COLOR,
            fg=SIDEBAR_MUTED,
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(4, 30))

        for item in ("Dashboard", "Students", "Tickets", "Results", "Logs"):
            label = tk.Label(sidebar, text=item, anchor="w", font=("Segoe UI", 10, "bold"), padx=14, pady=10, cursor="hand2")
            label.pack(fill=tk.X, pady=3)
            label.bind("<Button-1>", lambda _event, view=item: self._switch_view(view))
            self.nav_buttons[item] = label
        self._sync_sidebar()

        tk.Label(
            sidebar,
            text="Input files\nstudents.xlsx\ntickets.docx\n\nOutput\nresults.xlsx",
            bg=SIDEBAR_COLOR,
            fg=SIDEBAR_MUTED,
            justify=tk.LEFT,
            font=("Segoe UI", 9),
        ).pack(side=tk.BOTTOM, anchor="w")

    def _render_active_view(self) -> None:
        for child in self.main.winfo_children():
            child.destroy()

        if self.active_view == "Dashboard":
            self._build_header(self.main, "Генератор билетов", "Расширенная панель выдачи билетов с журналом, статистикой и быстрыми действиями.")
            self._build_metrics(self.main)
            self._build_workspace(self.main)
            self._build_history(self.main)
            self._sync_dashboard_inputs()
            self._refresh_dashboard()
        elif self.active_view == "Students":
            self._build_header(self.main, "Студенты", "Просмотр групп и студентов из students.xlsx.")
            self._build_students_view(self.main)
        elif self.active_view == "Tickets":
            self._build_header(self.main, "Билеты", "Просмотр корректно распознанных билетов из tickets.docx.")
            self._build_tickets_view(self.main)
        elif self.active_view == "Results":
            self._build_header(self.main, "Журнал результатов", "История всех генераций из results.xlsx.")
            self._build_results_view(self.main)
        elif self.active_view == "Logs":
            self._build_header(self.main, "Логи", "Ошибки парсинга и служебные сообщения приложения.")
            self._build_logs_view(self.main)

    def _switch_view(self, view: str) -> None:
        self.active_view = view
        self._sync_sidebar()
        self._render_active_view()

    def _sync_sidebar(self) -> None:
        for view, label in self.nav_buttons.items():
            active = view == self.active_view
            label.config(bg="#1d293b" if active else SIDEBAR_COLOR, fg="#ffffff" if active else SIDEBAR_MUTED)

    def _build_header(self, parent: ttk.Frame, title: str, subtitle: str) -> None:
        header = ttk.Frame(parent, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text=title, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=subtitle, style="Subtitle.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))

        actions = ttk.Frame(header, style="App.TFrame")
        actions.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Button(actions, text="Обновить данные", style="Ghost.TButton", command=self._reload_sources).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(actions, text="Открыть журнал", style="Ghost.TButton", command=self._open_results_file).pack(side=tk.LEFT)

    def _build_students_view(self, parent: ttk.Frame) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=22)
        card.grid(row=1, column=0, sticky="nsew", pady=(24, 0))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        ttk.Label(card, text="Все студенты по группам", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))
        columns = ("group", "last_name", "first_name")
        table = ttk.Treeview(card, columns=columns, show="headings", style="History.Treeview")
        table.heading("group", text="Группа")
        table.heading("last_name", text="Фамилия")
        table.heading("first_name", text="Имя")
        table.column("group", width=180)
        table.column("last_name", width=240)
        table.column("first_name", width=240)
        table.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(card, orient=tk.VERTICAL, command=table.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        table.configure(yscrollcommand=scrollbar.set)

        if self.student_store is None:
            return
        for group_name in self.student_store.group_names():
            students = self.student_store.students_for_group(group_name)
            if not students:
                table.insert("", tk.END, values=(group_name, "Нет студентов", ""))
            for student in students:
                table.insert("", tk.END, values=(group_name, student.last_name, student.first_name))

    def _build_tickets_view(self, parent: ttk.Frame) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=22)
        card.grid(row=1, column=0, sticky="nsew", pady=(24, 0))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)

        ttk.Label(card, text=f"Корректные билеты: {len(self.tickets)}", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))
        columns = ("number", "q1", "q2", "q3")
        table = ttk.Treeview(card, columns=columns, show="headings", style="History.Treeview")
        table.heading("number", text="Билет")
        table.heading("q1", text="Вопрос 1")
        table.heading("q2", text="Вопрос 2")
        table.heading("q3", text="Вопрос 3")
        table.column("number", width=70, anchor="center")
        table.column("q1", width=260)
        table.column("q2", width=260)
        table.column("q3", width=260)
        table.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(card, orient=tk.VERTICAL, command=table.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        table.configure(yscrollcommand=scrollbar.set)

        for number, ticket in sorted(self.tickets.items()):
            table.insert("", tk.END, values=(number, *ticket.questions))

    def _build_results_view(self, parent: ttk.Frame) -> None:
        stats = self.results_store.stats()
        metrics = ttk.Frame(parent, style="App.TFrame")
        metrics.grid(row=1, column=0, sticky="ew", pady=(24, 18))
        for column in range(4):
            metrics.columnconfigure(column, weight=1)
        values = [
            (stats.total_generations, "Всего выдач"),
            (stats.unique_students, "Уникальных студентов"),
            (stats.repeat_generations, "Повторов"),
            (stats.top_ticket if stats.top_ticket is not None else "-", "Самый частый билет"),
        ]
        for column, (value, title) in enumerate(values):
            card = ttk.Frame(metrics, style="Card.TFrame", padding=18)
            card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 10, 0))
            ttk.Label(card, text=str(value), style="Metric.TLabel").pack(anchor="w")
            ttk.Label(card, text=title, style="MetricName.TLabel").pack(anchor="w", pady=(3, 0))

        card = ttk.Frame(parent, style="Card.TFrame", padding=22)
        card.grid(row=2, column=0, sticky="nsew")
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        ttk.Label(card, text="Полный журнал", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))
        table = self._create_results_table(card)
        table.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(card, orient=tk.VERTICAL, command=table.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        table.configure(yscrollcommand=scrollbar.set)
        for entry in self.results_store.recent_entries(limit=10_000):
            table.insert("", tk.END, values=(entry.created_at, entry.group_name, entry.full_name, entry.ticket_number, entry.repeat))

    def _build_logs_view(self, parent: ttk.Frame) -> None:
        card = ttk.Frame(parent, style="Card.TFrame", padding=22)
        card.grid(row=1, column=0, sticky="nsew", pady=(24, 0))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(1, weight=1)
        ttk.Label(card, text="app.log", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))

        text = tk.Text(card, wrap=tk.WORD, borderwidth=0, font=("Consolas", 10), bg="#f8fafc", fg=TEXT_COLOR)
        text.grid(row=1, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(card, orient=tk.VERTICAL, command=text.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        text.configure(yscrollcommand=scrollbar.set)
        if LOG_FILE.exists():
            content = LOG_FILE.read_text(encoding="utf-8")
        else:
            content = "Лог пока пуст. Ошибки парсинга билетов появятся здесь."
        text.insert("1.0", content)
        text.configure(state=tk.DISABLED)

    def _create_results_table(self, parent: ttk.Frame) -> ttk.Treeview:
        columns = ("time", "group", "student", "ticket", "repeat")
        table = ttk.Treeview(parent, columns=columns, show="headings", style="History.Treeview")
        table.heading("time", text="Дата и время")
        table.heading("group", text="Группа")
        table.heading("student", text="Студент")
        table.heading("ticket", text="Билет")
        table.heading("repeat", text="Повтор")
        table.column("time", width=160, anchor="w")
        table.column("group", width=100, anchor="w")
        table.column("student", width=220, anchor="w")
        table.column("ticket", width=70, anchor="center")
        table.column("repeat", width=80, anchor="center")
        return table

    def _build_metrics(self, parent: ttk.Frame) -> None:
        self.metric_frame = ttk.Frame(parent, style="App.TFrame")
        self.metric_frame.grid(row=1, column=0, sticky="ew", pady=(24, 18))
        for column in range(4):
            self.metric_frame.columnconfigure(column, weight=1)

        self.metric_labels: dict[str, ttk.Label] = {}
        metrics = [
            ("groups", "0", "Групп"),
            ("tickets", "0", "Билетов"),
            ("generations", "0", "Записей в журнале"),
            ("repeats", "0", "Повторов"),
        ]
        for column, (key, value, title) in enumerate(metrics):
            card = ttk.Frame(self.metric_frame, style="Card.TFrame", padding=18)
            card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 10, 0))
            label = ttk.Label(card, text=value, style="Metric.TLabel")
            label.pack(anchor="w")
            ttk.Label(card, text=title, style="MetricName.TLabel").pack(anchor="w", pady=(3, 0))
            self.metric_labels[key] = label

    def _build_workspace(self, parent: ttk.Frame) -> None:
        workspace = ttk.Frame(parent, style="App.TFrame")
        workspace.grid(row=2, column=0, sticky="ew")
        workspace.columnconfigure(0, weight=2)
        workspace.columnconfigure(1, weight=1)

        generator = ttk.Frame(workspace, style="Card.TFrame", padding=22)
        generator.grid(row=0, column=0, sticky="nsew", padx=(0, 18))
        generator.columnconfigure(0, weight=1)
        generator.columnconfigure(1, weight=1)

        ttk.Label(generator, text="Выдача билета", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(generator, text="Выберите группу и студента. Кнопка активируется автоматически.", style="Muted.TLabel").grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(4, 18),
        )

        ttk.Label(generator, text="Группа", style="Field.TLabel").grid(row=2, column=0, sticky="w", padx=(0, 12))
        ttk.Label(generator, text="Студент", style="Field.TLabel").grid(row=2, column=1, sticky="w", padx=(12, 0))

        self.group_combo = ttk.Combobox(generator, textvariable=self.group_var, state="readonly", font=("Segoe UI", 10))
        self.group_combo.grid(row=3, column=0, sticky="ew", padx=(0, 12), pady=(8, 16), ipady=5)
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_selected)

        self.student_combo = ttk.Combobox(generator, textvariable=self.student_var, state="readonly", font=("Segoe UI", 10))
        self.student_combo.grid(row=3, column=1, sticky="ew", padx=(12, 0), pady=(8, 16), ipady=5)
        self.student_combo.bind("<<ComboboxSelected>>", self._on_student_selected)

        self.generate_button = ttk.Button(
            generator,
            text="Сгенерировать билет",
            command=self._generate_ticket,
            state=tk.DISABLED,
            style="Primary.TButton",
        )
        self.generate_button.grid(row=4, column=0, sticky="ew", padx=(0, 12))
        ttk.Button(generator, text="Сбросить выбор", style="Ghost.TButton", command=self._clear_selection).grid(
            row=4,
            column=1,
            sticky="ew",
            padx=(12, 0),
        )

        self.status_accent = tk.Frame(generator, bg=PRIMARY_COLOR, height=4)
        self.status_accent.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(22, 14))
        self.status_title = ttk.Label(generator, text="Ожидание данных", style="Field.TLabel")
        self.status_title.grid(row=6, column=0, columnspan=2, sticky="w")
        self.status_label = ttk.Label(generator, text="", style="Muted.TLabel", wraplength=610)
        self.status_label.grid(row=7, column=0, columnspan=2, sticky="w", pady=(6, 0))

        students = ttk.Frame(workspace, style="Card.TFrame", padding=22)
        students.grid(row=0, column=1, sticky="nsew")
        students.columnconfigure(0, weight=1)
        students.rowconfigure(2, weight=1)

        ttk.Label(students, text="Студенты группы", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w")
        self.group_count_label = ttk.Label(students, text="Выберите группу", style="Muted.TLabel")
        self.group_count_label.grid(row=1, column=0, sticky="w", pady=(4, 12))

        self.student_list = tk.Listbox(
            students,
            height=7,
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=LINE_COLOR,
            selectbackground=PRIMARY_COLOR,
            font=("Segoe UI", 10),
        )
        self.student_list.grid(row=2, column=0, sticky="nsew")
        self.student_list.bind("<<ListboxSelect>>", self._on_student_list_selected)

    def _build_history(self, parent: ttk.Frame) -> None:
        history = ttk.Frame(parent, style="Card.TFrame", padding=22)
        history.grid(row=3, column=0, sticky="nsew", pady=(18, 0))
        history.columnconfigure(0, weight=1)
        history.rowconfigure(1, weight=1)

        ttk.Label(history, text="Последние выдачи", style="CardTitle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 12))

        columns = ("time", "group", "student", "ticket", "repeat")
        self.history_table = ttk.Treeview(history, columns=columns, show="headings", style="History.Treeview", height=8)
        self.history_table.heading("time", text="Дата и время")
        self.history_table.heading("group", text="Группа")
        self.history_table.heading("student", text="Студент")
        self.history_table.heading("ticket", text="Билет")
        self.history_table.heading("repeat", text="Повтор")
        self.history_table.column("time", width=160, anchor="w")
        self.history_table.column("group", width=100, anchor="w")
        self.history_table.column("student", width=220, anchor="w")
        self.history_table.column("ticket", width=70, anchor="center")
        self.history_table.column("repeat", width=80, anchor="center")
        self.history_table.grid(row=1, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(history, orient=tk.VERTICAL, command=self.history_table.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.history_table.configure(yscrollcommand=scrollbar.set)

    def _load_sources(self) -> None:
        missing_files = [str(path.name) for path in (STUDENTS_FILE, TICKETS_FILE) if not path.exists()]
        if missing_files:
            self._set_unavailable(f"Не найдены файлы: {', '.join(missing_files)}.")
            self._refresh_dashboard()
            return

        try:
            self.student_store = StudentStore(STUDENTS_FILE)
            self.tickets = parse_tickets(TICKETS_FILE, self.logger)
        except Exception as error:
            self._set_unavailable(str(error))
            self._refresh_dashboard()
            return

        if not self.tickets:
            self._set_unavailable("В tickets.docx не найдено ни одного корректного билета.")
            self._refresh_dashboard()
            return

        groups = self.student_store.group_names()
        self._sync_dashboard_inputs(groups)
        self._set_status("Данные загружены", f"Корректных билетов: {len(self.tickets)}. Выберите группу и студента.", PRIMARY_COLOR)
        self._refresh_dashboard()

    def _reload_sources(self) -> None:
        self.group_var.set("")
        self.student_var.set("")
        self.current_students = []
        if hasattr(self, "group_combo") and self.group_combo.winfo_exists():
            self.group_combo["values"] = []
            self.student_combo["values"] = []
            self.student_list.delete(0, tk.END)
            self.generate_button.config(state=tk.DISABLED)
        self._load_sources()
        self._render_active_view()

    def _set_unavailable(self, message: str) -> None:
        self._set_status("Ошибка загрузки", message, ERROR_COLOR)
        if hasattr(self, "group_combo") and self.group_combo.winfo_exists():
            self.group_combo.config(state=tk.DISABLED)
            self.student_combo.config(state=tk.DISABLED)
            self.generate_button.config(state=tk.DISABLED)
        messagebox.showerror("Ошибка загрузки", message)

    def _set_status(self, title: str, message: str, color: str = PRIMARY_COLOR) -> None:
        if hasattr(self, "status_title") and self.status_title.winfo_exists():
            self.status_title.config(text=title)
            self.status_label.config(text=message)
            self.status_accent.config(bg=color)

    def _sync_dashboard_inputs(self, groups: list[str] | None = None) -> None:
        if self.active_view != "Dashboard" or not hasattr(self, "group_combo") or not self.group_combo.winfo_exists():
            return
        if groups is None and self.student_store is not None:
            groups = self.student_store.group_names()
        groups = groups or []
        self.group_combo.config(state="readonly")
        self.student_combo.config(state="readonly")
        self.group_combo["values"] = groups
        if self.group_var.get() in groups:
            self._on_group_selected()

    def _refresh_dashboard(self) -> None:
        if self.active_view != "Dashboard" or not hasattr(self, "metric_labels"):
            return
        groups = list(self.group_combo["values"]) if self.group_combo["values"] else []
        stats = self.results_store.stats()
        self.metric_labels["groups"].config(text=str(len(groups)))
        self.metric_labels["tickets"].config(text=str(len(self.tickets)))
        self.metric_labels["generations"].config(text=str(stats.total_generations))
        self.metric_labels["repeats"].config(text=str(stats.repeat_generations))

        self.history_table.delete(*self.history_table.get_children())
        for entry in self.results_store.recent_entries(limit=12):
            self.history_table.insert("", tk.END, values=(entry.created_at, entry.group_name, entry.full_name, entry.ticket_number, entry.repeat))

    def _on_group_selected(self, _event: object | None = None) -> None:
        if self.student_store is None:
            return
        group_name = self.group_var.get()
        self.current_students = self.student_store.students_for_group(group_name)
        self.student_var.set("")
        self.student_combo["values"] = [student.full_name for student in self.current_students]
        self.student_list.delete(0, tk.END)
        for student in self.current_students:
            self.student_list.insert(tk.END, student.full_name)
        self.generate_button.config(state=tk.DISABLED)

        if self.current_students:
            self.group_count_label.config(text=f"Найдено студентов: {len(self.current_students)}")
            self._set_status(
                "Группа выбрана",
                f"В группе найдено студентов: {len(self.current_students)}. Теперь выберите студента.",
                SUCCESS_COLOR,
            )
        else:
            self.group_count_label.config(text="В группе нет студентов")
            self._set_status("Пустая группа", "В выбранной группе нет студентов. Кнопка генерации остается выключенной.", WARNING_COLOR)

    def _on_student_selected(self, _event: object | None = None) -> None:
        student = self._selected_student()
        self.generate_button.config(state=tk.NORMAL if student else tk.DISABLED)
        if student:
            self._set_status("Готово к генерации", f"Выбран студент: {student.full_name}.", SUCCESS_COLOR)

    def _on_student_list_selected(self, _event: object | None = None) -> None:
        selection = self.student_list.curselection()
        if not selection:
            return
        selected_name = self.student_list.get(selection[0])
        self.student_var.set(selected_name)
        self._on_student_selected()

    def _clear_selection(self) -> None:
        self.group_var.set("")
        self.student_var.set("")
        self.current_students = []
        self.student_combo["values"] = []
        self.student_list.delete(0, tk.END)
        self.group_count_label.config(text="Выберите группу")
        self.generate_button.config(state=tk.DISABLED)
        self._set_status("Выбор сброшен", "Выберите группу и студента для новой выдачи.", PRIMARY_COLOR)

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

        self._refresh_dashboard()
        self._show_result(ticket, assignment.is_repeat)

    def _open_results_file(self) -> None:
        if not RESULTS_FILE.exists():
            messagebox.showinfo("Журнал", "results.xlsx пока не создан. Сгенерируйте первый билет.")
            return
        os.startfile(RESULTS_FILE)

    def _show_result(self, ticket: Ticket, is_repeat: bool) -> None:
        window = tk.Toplevel(self)
        window.title(f"Билет № {ticket.number}")
        window.geometry("760x500")
        window.minsize(650, 420)
        window.configure(bg=BG_COLOR)
        window.transient(self)
        window.grab_set()

        shell = ttk.Frame(window, style="App.TFrame", padding=28)
        shell.pack(fill=tk.BOTH, expand=True)

        frame = ttk.Frame(shell, style="Card.TFrame", padding=30)
        frame.pack(fill=tk.BOTH, expand=True)

        badge_text = "Повторная выдача" if is_repeat else "Новая выдача"
        badge_color = WARNING_COLOR if is_repeat else SUCCESS_COLOR
        badge = tk.Label(frame, text=badge_text, bg=badge_color, fg="#ffffff", font=("Segoe UI", 9, "bold"), padx=12, pady=5)
        badge.pack(anchor="w")

        ttk.Label(frame, text=f"Ваш билет № {ticket.number}", style="CardTitle.TLabel").pack(anchor="w", pady=(18, 4))
        help_text = "Номер взят из первой записи журнала для этого студента." if is_repeat else "Билет выбран случайно из корректных билетов."
        ttk.Label(frame, text=help_text, style="Muted.TLabel").pack(anchor="w", pady=(0, 22))

        for index, question in enumerate(ticket.questions, start=1):
            question_frame = ttk.Frame(frame, style="Card.TFrame")
            question_frame.pack(fill=tk.X, pady=8)
            number = tk.Label(question_frame, text=str(index), bg=PRIMARY_COLOR, fg="#ffffff", font=("Segoe UI", 11, "bold"), width=3)
            number.pack(side=tk.LEFT, anchor="n", padx=(0, 14))
            ttk.Label(question_frame, text=question, wraplength=620, justify=tk.LEFT, style="Question.TLabel").pack(
                side=tk.LEFT,
                fill=tk.X,
                expand=True,
                anchor="w",
            )

        ttk.Label(frame, text="Esc - закрыть окно результата и вернуться к панели.", style="Muted.TLabel").pack(anchor="w", pady=(28, 0))
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
