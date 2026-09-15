from openpyxl import Workbook

from student_store import Student, StudentStore


def test_students_for_group_reads_selected_sheet(tmp_path):
    workbook_path = tmp_path / "students.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ИС-21"
    sheet.append(["Фамилия", "Имя"])
    sheet.append(["Иванов", "Иван"])
    sheet.append(["Петров", "Петр"])
    empty_sheet = workbook.create_sheet("ИС-22")
    empty_sheet.append(["Фамилия", "Имя"])
    workbook.save(workbook_path)

    store = StudentStore(workbook_path)

    assert store.group_names() == ["ИС-21", "ИС-22"]
    assert store.students_for_group("ИС-21") == [
        Student(last_name="Иванов", first_name="Иван"),
        Student(last_name="Петров", first_name="Петр"),
    ]
    assert store.students_for_group("ИС-22") == []
