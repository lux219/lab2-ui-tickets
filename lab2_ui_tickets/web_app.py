from __future__ import annotations

import logging
from pathlib import Path

from flask import Flask, jsonify, render_template_string, request

from results_store import ResultsStore, ResultsWriteError
from student_store import Student, StudentStore
from ticket_parser import parse_tickets


APP_DIR = Path(__file__).resolve().parent
STUDENTS_FILE = APP_DIR / "students.xlsx"
TICKETS_FILE = APP_DIR / "tickets.docx"
RESULTS_FILE = APP_DIR / "results.xlsx"
LOG_FILE = APP_DIR / "app.log"

app = Flask(__name__)


def _logger() -> logging.Logger:
    logger = logging.getLogger("lab2_ui_tickets_web")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
    return logger


def _load_sources() -> tuple[StudentStore | None, dict, list[str]]:
    errors: list[str] = []
    missing = [path.name for path in (STUDENTS_FILE, TICKETS_FILE) if not path.exists()]
    if missing:
        return None, {}, [f"Не найдены файлы: {', '.join(missing)}."]

    student_store: StudentStore | None = None
    tickets = {}
    try:
        student_store = StudentStore(STUDENTS_FILE)
        tickets = parse_tickets(TICKETS_FILE, _logger())
        if not tickets:
            errors.append("В tickets.docx не найдено ни одного корректного билета.")
    except Exception as error:
        errors.append(str(error))
    return student_store, tickets, errors


def _state() -> dict:
    student_store, tickets, errors = _load_sources()
    results = ResultsStore(RESULTS_FILE)
    stats = results.stats()
    groups = student_store.group_names() if student_store else []
    recent = results.recent_entries(limit=12)
    logs = LOG_FILE.read_text(encoding="utf-8") if LOG_FILE.exists() else "Лог пока пуст."
    return {
        "ok": not errors,
        "errors": errors,
        "groups": groups,
        "tickets": [
            {"number": ticket.number, "questions": list(ticket.questions)}
            for _, ticket in sorted(tickets.items())
        ],
        "stats": {
            "groups": len(groups),
            "tickets": len(tickets),
            "totalGenerations": stats.total_generations,
            "uniqueStudents": stats.unique_students,
            "repeatGenerations": stats.repeat_generations,
            "topTicket": stats.top_ticket,
        },
        "recent": [
            {
                "time": entry.created_at,
                "group": entry.group_name,
                "student": entry.full_name,
                "ticket": entry.ticket_number,
                "repeat": entry.repeat,
            }
            for entry in recent
        ],
        "logs": logs,
    }


@app.get("/")
def index():
    return render_template_string(PAGE)


@app.get("/api/state")
def api_state():
    return jsonify(_state())


@app.get("/api/students")
def api_students():
    group_name = request.args.get("group", "")
    student_store, _tickets, errors = _load_sources()
    if errors or student_store is None:
        return jsonify({"ok": False, "errors": errors, "students": []}), 400
    students = student_store.students_for_group(group_name)
    return jsonify(
        {
            "ok": True,
            "students": [
                {"lastName": student.last_name, "firstName": student.first_name, "fullName": student.full_name}
                for student in students
            ],
        }
    )


@app.post("/api/generate")
def api_generate():
    data = request.get_json(silent=True) or {}
    group_name = str(data.get("group", "")).strip()
    full_name = str(data.get("student", "")).strip()
    if not group_name or not full_name:
        return jsonify({"ok": False, "message": "Выберите группу и студента."}), 400

    student_store, tickets, errors = _load_sources()
    if errors or student_store is None:
        return jsonify({"ok": False, "message": " ".join(errors)}), 400

    student = next((item for item in student_store.students_for_group(group_name) if item.full_name == full_name), None)
    if student is None:
        return jsonify({"ok": False, "message": "Студент не найден в выбранной группе."}), 404

    try:
        assignment = ResultsStore(RESULTS_FILE).record_assignment(group_name, student, sorted(tickets))
    except ResultsWriteError as error:
        return jsonify({"ok": False, "message": str(error)}), 423

    ticket = tickets.get(assignment.ticket_number)
    if ticket is None:
        return jsonify({"ok": False, "message": f"Билет № {assignment.ticket_number} не найден."}), 500

    return jsonify(
        {
            "ok": True,
            "ticket": {"number": ticket.number, "questions": list(ticket.questions)},
            "repeat": assignment.is_repeat,
            "state": _state(),
        }
    )


PAGE = r"""
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ExamFlow Ticket Service</title>
  <style>
    :root {
      --bg: #eef3fb;
      --panel: #ffffff;
      --ink: #172033;
      --muted: #6b7280;
      --line: #dbe4f2;
      --blue: #2563eb;
      --blue-dark: #1d4ed8;
      --green: #0f766e;
      --orange: #b45309;
      --red: #b91c1c;
      --nav: #101827;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      min-height: 100vh;
      font-family: "Segoe UI", system-ui, sans-serif;
      color: var(--ink);
      background: radial-gradient(circle at top left, #dbeafe 0, transparent 34rem), var(--bg);
    }
    .app { display: grid; grid-template-columns: 250px 1fr; min-height: 100vh; }
    .sidebar { background: var(--nav); padding: 28px 22px; color: #fff; }
    .brand { font-size: 26px; font-weight: 800; letter-spacing: .2px; }
    .tagline { color: #8ea0bd; margin: 6px 0 34px; font-size: 13px; }
    .nav button {
      width: 100%;
      border: 0;
      background: transparent;
      color: #8ea0bd;
      text-align: left;
      padding: 13px 14px;
      border-radius: 8px;
      font: 700 15px "Segoe UI";
      cursor: pointer;
      margin-bottom: 5px;
    }
    .nav button.active, .nav button:hover { background: #1d293b; color: #fff; }
    .file-note { position: fixed; bottom: 28px; width: 190px; color: #8ea0bd; font-size: 13px; line-height: 1.55; }
    main { padding: 30px; overflow: auto; }
    .topbar { display: flex; justify-content: space-between; gap: 20px; align-items: center; margin-bottom: 24px; }
    h1 { margin: 0; font-size: 34px; }
    .subtitle { margin-top: 6px; color: var(--muted); }
    .actions { display: flex; gap: 10px; }
    .btn {
      border: 0;
      border-radius: 10px;
      padding: 12px 16px;
      font-weight: 800;
      cursor: pointer;
      color: var(--blue);
      background: #eaf1ff;
    }
    .btn.primary { color: #fff; background: var(--blue); }
    .btn.primary:hover { background: var(--blue-dark); }
    .btn:disabled { cursor: not-allowed; background: #cbd5e1; color: #64748b; }
    .metrics { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; margin-bottom: 20px; }
    .card {
      background: var(--panel);
      border: 1px solid rgba(148, 163, 184, .22);
      border-radius: 18px;
      padding: 22px;
      box-shadow: 0 18px 45px rgba(15, 23, 42, .08);
    }
    .metric-value { font-size: 30px; font-weight: 850; }
    .metric-title { color: var(--muted); font-size: 13px; margin-top: 4px; }
    .grid { display: grid; grid-template-columns: 2fr 1fr; gap: 18px; margin-bottom: 20px; }
    h2 { margin: 0 0 6px; font-size: 20px; }
    .hint { color: var(--muted); font-size: 14px; margin-bottom: 18px; }
    label { display: block; font-weight: 800; margin-bottom: 8px; }
    select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 12px;
      padding: 13px 14px;
      font: 15px "Segoe UI";
      background: #f8fafc;
      color: var(--ink);
    }
    .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }
    .form-actions { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-top: 20px; }
    .status {
      border-top: 5px solid var(--blue);
      padding-top: 15px;
      margin-top: 22px;
      color: var(--muted);
    }
    .students-list { list-style: none; padding: 0; margin: 12px 0 0; max-height: 255px; overflow: auto; }
    .students-list li { padding: 10px 12px; border-bottom: 1px solid #edf2f7; border-radius: 9px; }
    .students-list li:hover { background: #eff6ff; cursor: pointer; }
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th, td { padding: 12px 10px; border-bottom: 1px solid #e5edf7; text-align: left; vertical-align: top; }
    th { color: var(--muted); font-size: 12px; text-transform: uppercase; letter-spacing: .04em; background: #f8fafc; }
    .section { display: none; }
    .section.active { display: block; }
    .ticket-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px; }
    .ticket-card { background: #f8fafc; border: 1px solid var(--line); border-radius: 16px; padding: 18px; }
    .ticket-number { color: var(--blue); font-weight: 900; margin-bottom: 10px; }
    pre { white-space: pre-wrap; background: #0f172a; color: #dbeafe; border-radius: 16px; padding: 18px; min-height: 420px; overflow: auto; }
    .modal {
      position: fixed; inset: 0; background: rgba(15, 23, 42, .55); display: none;
      align-items: center; justify-content: center; padding: 24px;
    }
    .modal.open { display: flex; }
    .modal-card { max-width: 760px; width: 100%; }
    .badge { display: inline-block; color: #fff; border-radius: 999px; padding: 7px 12px; font-weight: 800; font-size: 13px; }
    .badge.new { background: var(--green); }
    .badge.repeat { background: var(--orange); }
    .question { display: grid; grid-template-columns: 36px 1fr; gap: 14px; margin: 16px 0; }
    .qnum { background: var(--blue); color: #fff; border-radius: 10px; text-align: center; padding: 8px 0; font-weight: 900; height: 36px; }
    .error { color: var(--red); font-weight: 800; }
    @media (max-width: 920px) {
      .app { grid-template-columns: 1fr; }
      .sidebar { position: static; }
      .file-note { position: static; margin-top: 28px; }
      .metrics, .grid, .form-grid { grid-template-columns: 1fr; }
      .topbar { flex-direction: column; align-items: stretch; }
    }
  </style>
</head>
<body>
  <div class="app">
    <aside class="sidebar">
      <div class="brand">ExamFlow</div>
      <div class="tagline">Web ticket generation service</div>
      <nav class="nav">
        <button class="active" data-view="dashboard">Dashboard</button>
        <button data-view="students">Students</button>
        <button data-view="tickets">Tickets</button>
        <button data-view="results">Results</button>
        <button data-view="logs">Logs</button>
      </nav>
      <div class="file-note">Input files<br>students.xlsx<br>tickets.docx<br><br>Output<br>results.xlsx</div>
    </aside>
    <main>
      <div class="topbar">
        <div>
          <h1 id="pageTitle">Генератор билетов</h1>
          <div class="subtitle" id="pageSubtitle">Веб-панель выдачи билетов с журналом, статистикой и историей.</div>
        </div>
        <div class="actions">
          <button class="btn" id="refreshBtn">Обновить данные</button>
          <a class="btn" href="/api/state" target="_blank">API state</a>
        </div>
      </div>

      <section id="dashboard" class="section active">
        <div class="metrics">
          <div class="card"><div class="metric-value" id="mGroups">0</div><div class="metric-title">Групп</div></div>
          <div class="card"><div class="metric-value" id="mTickets">0</div><div class="metric-title">Билетов</div></div>
          <div class="card"><div class="metric-value" id="mGenerations">0</div><div class="metric-title">Записей</div></div>
          <div class="card"><div class="metric-value" id="mRepeats">0</div><div class="metric-title">Повторов</div></div>
        </div>
        <div class="grid">
          <div class="card">
            <h2>Выдача билета</h2>
            <div class="hint">Выберите группу и студента. Результат сразу сохранится в журнал.</div>
            <div class="form-grid">
              <div><label for="groupSelect">Группа</label><select id="groupSelect"></select></div>
              <div><label for="studentSelect">Студент</label><select id="studentSelect" disabled></select></div>
            </div>
            <div class="form-actions">
              <button class="btn primary" id="generateBtn" disabled>Сгенерировать билет</button>
              <button class="btn" id="resetBtn">Сбросить выбор</button>
            </div>
            <div class="status" id="statusBox">Загрузка данных...</div>
          </div>
          <div class="card">
            <h2>Студенты группы</h2>
            <div class="hint" id="groupHint">Выберите группу</div>
            <ul class="students-list" id="studentsList"></ul>
          </div>
        </div>
        <div class="card">
          <h2>Последние выдачи</h2>
          <div class="hint">Последние записи из results.xlsx</div>
          <table><thead><tr><th>Дата</th><th>Группа</th><th>Студент</th><th>Билет</th><th>Повтор</th></tr></thead><tbody id="recentRows"></tbody></table>
        </div>
      </section>

      <section id="students" class="section"><div class="card"><h2>Студенты</h2><div class="hint">Все группы и студенты из Excel.</div><table><thead><tr><th>Группа</th><th>Фамилия</th><th>Имя</th></tr></thead><tbody id="studentsRows"></tbody></table></div></section>
      <section id="tickets" class="section"><div class="ticket-grid" id="ticketCards"></div></section>
      <section id="results" class="section"><div class="card"><h2>Полный журнал</h2><div class="hint">История генераций в обратном хронологическом порядке.</div><table><thead><tr><th>Дата</th><th>Группа</th><th>Студент</th><th>Билет</th><th>Повтор</th></tr></thead><tbody id="resultRows"></tbody></table></div></section>
      <section id="logs" class="section"><pre id="logsText">Лог пока пуст.</pre></section>
    </main>
  </div>

  <div class="modal" id="resultModal">
    <div class="card modal-card">
      <span class="badge new" id="resultBadge">Новая выдача</span>
      <h1 id="resultTitle">Ваш билет</h1>
      <div class="subtitle" id="resultHelp"></div>
      <div id="resultQuestions"></div>
      <button class="btn primary" id="closeModal">Закрыть</button>
    </div>
  </div>

  <script>
    const titles = {
      dashboard: ["Генератор билетов", "Веб-панель выдачи билетов с журналом, статистикой и историей."],
      students: ["Студенты", "Просмотр групп и студентов из students.xlsx."],
      tickets: ["Билеты", "Просмотр корректно распознанных билетов из tickets.docx."],
      results: ["Журнал результатов", "Полная история выдач из results.xlsx."],
      logs: ["Логи", "Ошибки парсинга и служебные сообщения приложения."]
    };
    let state = null;

    async function loadState() {
      const response = await fetch("/api/state");
      state = await response.json();
      renderState();
    }

    function renderState() {
      document.getElementById("mGroups").textContent = state.stats.groups;
      document.getElementById("mTickets").textContent = state.stats.tickets;
      document.getElementById("mGenerations").textContent = state.stats.totalGenerations;
      document.getElementById("mRepeats").textContent = state.stats.repeatGenerations;
      renderGroups();
      renderRecent();
      renderTickets();
      renderResults();
      renderLogs();
      renderStudentsOverview();
      if (!state.ok) setStatus(state.errors.join(" "), true);
      else setStatus("Данные загружены. Выберите группу и студента.", false);
    }

    function renderGroups() {
      const select = document.getElementById("groupSelect");
      const selected = select.value;
      select.innerHTML = `<option value="">Выберите группу</option>` + state.groups.map(group => `<option>${escapeHtml(group)}</option>`).join("");
      if (state.groups.includes(selected)) select.value = selected;
    }

    async function loadStudents(group) {
      const studentSelect = document.getElementById("studentSelect");
      const list = document.getElementById("studentsList");
      studentSelect.innerHTML = `<option value="">Выберите студента</option>`;
      studentSelect.disabled = true;
      list.innerHTML = "";
      document.getElementById("generateBtn").disabled = true;
      if (!group) {
        document.getElementById("groupHint").textContent = "Выберите группу";
        return;
      }
      const response = await fetch(`/api/students?group=${encodeURIComponent(group)}`);
      const data = await response.json();
      document.getElementById("groupHint").textContent = data.students.length ? `Найдено студентов: ${data.students.length}` : "В группе нет студентов";
      studentSelect.disabled = data.students.length === 0;
      studentSelect.innerHTML += data.students.map(student => `<option>${escapeHtml(student.fullName)}</option>`).join("");
      list.innerHTML = data.students.map(student => `<li data-name="${escapeAttr(student.fullName)}">${escapeHtml(student.fullName)}</li>`).join("");
      list.querySelectorAll("li").forEach(item => item.addEventListener("click", () => {
        studentSelect.value = item.dataset.name;
        document.getElementById("generateBtn").disabled = false;
        setStatus(`Выбран студент: ${item.dataset.name}`, false);
      }));
    }

    function renderRecent() {
      const rows = state.recent.map(entry => row(entry)).join("");
      document.getElementById("recentRows").innerHTML = rows || `<tr><td colspan="5">Журнал пока пуст.</td></tr>`;
    }

    function renderResults() {
      document.getElementById("resultRows").innerHTML = state.recent.map(entry => row(entry)).join("") || `<tr><td colspan="5">Журнал пока пуст.</td></tr>`;
    }

    function renderTickets() {
      document.getElementById("ticketCards").innerHTML = state.tickets.map(ticket => `
        <article class="ticket-card">
          <div class="ticket-number">Билет ${ticket.number}</div>
          ${ticket.questions.map((question, index) => `<div class="question"><div class="qnum">${index + 1}</div><div>${escapeHtml(question)}</div></div>`).join("")}
        </article>`).join("") || `<div class="card error">Корректных билетов нет.</div>`;
    }

    async function renderStudentsOverview() {
      const rows = [];
      for (const group of state.groups) {
        const response = await fetch(`/api/students?group=${encodeURIComponent(group)}`);
        const data = await response.json();
        if (!data.students.length) rows.push(`<tr><td>${escapeHtml(group)}</td><td>Нет студентов</td><td></td></tr>`);
        data.students.forEach(student => rows.push(`<tr><td>${escapeHtml(group)}</td><td>${escapeHtml(student.lastName)}</td><td>${escapeHtml(student.firstName)}</td></tr>`));
      }
      document.getElementById("studentsRows").innerHTML = rows.join("") || `<tr><td colspan="3">Нет данных.</td></tr>`;
    }

    function renderLogs() {
      document.getElementById("logsText").textContent = state.logs || "Лог пока пуст.";
    }

    function row(entry) {
      return `<tr><td>${escapeHtml(entry.time)}</td><td>${escapeHtml(entry.group)}</td><td>${escapeHtml(entry.student)}</td><td>${entry.ticket}</td><td>${escapeHtml(entry.repeat)}</td></tr>`;
    }

    function setStatus(message, isError) {
      const box = document.getElementById("statusBox");
      box.textContent = message;
      box.style.borderTopColor = isError ? "var(--red)" : "var(--blue)";
      box.className = isError ? "status error" : "status";
    }

    async function generateTicket() {
      const response = await fetch("/api/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ group: groupSelect.value, student: studentSelect.value })
      });
      const data = await response.json();
      if (!data.ok) {
        setStatus(data.message, true);
        return;
      }
      state = data.state;
      renderState();
      openResult(data.ticket, data.repeat);
    }

    function openResult(ticket, repeat) {
      resultBadge.textContent = repeat ? "Повторная выдача" : "Новая выдача";
      resultBadge.className = repeat ? "badge repeat" : "badge new";
      resultTitle.textContent = `Ваш билет № ${ticket.number}`;
      resultHelp.textContent = repeat ? "Номер взят из первой записи журнала для этого студента." : "Билет выбран случайно из корректных билетов.";
      resultQuestions.innerHTML = ticket.questions.map((question, index) => `<div class="question"><div class="qnum">${index + 1}</div><div>${escapeHtml(question)}</div></div>`).join("");
      resultModal.classList.add("open");
    }

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
    }
    function escapeAttr(value) { return escapeHtml(value).replace(/"/g, "&quot;"); }

    document.querySelectorAll(".nav button").forEach(button => button.addEventListener("click", () => {
      document.querySelectorAll(".nav button").forEach(item => item.classList.remove("active"));
      button.classList.add("active");
      document.querySelectorAll(".section").forEach(item => item.classList.remove("active"));
      document.getElementById(button.dataset.view).classList.add("active");
      pageTitle.textContent = titles[button.dataset.view][0];
      pageSubtitle.textContent = titles[button.dataset.view][1];
    }));
    groupSelect.addEventListener("change", () => loadStudents(groupSelect.value));
    studentSelect.addEventListener("change", () => {
      generateBtn.disabled = !studentSelect.value;
      if (studentSelect.value) setStatus(`Выбран студент: ${studentSelect.value}`, false);
    });
    resetBtn.addEventListener("click", () => {
      groupSelect.value = "";
      studentSelect.innerHTML = `<option value="">Выберите студента</option>`;
      studentSelect.disabled = true;
      studentsList.innerHTML = "";
      groupHint.textContent = "Выберите группу";
      generateBtn.disabled = true;
      setStatus("Выбор сброшен.", false);
    });
    generateBtn.addEventListener("click", generateTicket);
    refreshBtn.addEventListener("click", loadState);
    closeModal.addEventListener("click", () => resultModal.classList.remove("open"));
    document.addEventListener("keydown", event => {
      if (event.key === "Escape") resultModal.classList.remove("open");
    });
    loadState();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    app.run(debug=True)
