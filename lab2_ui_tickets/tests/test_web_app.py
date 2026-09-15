from __future__ import annotations

from sample_data import create_students, create_tickets


def test_web_state_and_generate_api(tmp_path, monkeypatch):
    import web_app

    students_path = tmp_path / "students.xlsx"
    tickets_path = tmp_path / "tickets.docx"
    results_path = tmp_path / "results.xlsx"
    log_path = tmp_path / "app.log"
    create_students(students_path)
    create_tickets(tickets_path)

    monkeypatch.setattr(web_app, "STUDENTS_FILE", students_path)
    monkeypatch.setattr(web_app, "TICKETS_FILE", tickets_path)
    monkeypatch.setattr(web_app, "RESULTS_FILE", results_path)
    monkeypatch.setattr(web_app, "LOG_FILE", log_path)

    web_app.app.config.update(TESTING=True)
    client = web_app.app.test_client()

    state_response = client.get("/api/state")
    state = state_response.get_json()
    assert state_response.status_code == 200
    assert state["ok"] is True
    assert "ИС-21" in state["groups"]
    assert state["stats"]["tickets"] == 5

    students_response = client.get("/api/students?group=ИС-21")
    students = students_response.get_json()
    assert students_response.status_code == 200
    assert students["students"][0]["fullName"] == "Иванов Иван"

    generate_response = client.post("/api/generate", json={"group": "ИС-21", "student": "Иванов Иван"})
    generated = generate_response.get_json()
    assert generate_response.status_code == 200
    assert generated["ok"] is True
    assert len(generated["ticket"]["questions"]) == 3
    assert generated["state"]["stats"]["totalGenerations"] == 1
