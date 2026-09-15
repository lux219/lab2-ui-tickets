from __future__ import annotations

import random
from collections import Counter
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from collections.abc import Sequence
from typing import Callable

from openpyxl import Workbook, load_workbook

from student_store import Student


HEADERS = ("Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)")


class ResultsWriteError(RuntimeError):
    pass


@dataclass(frozen=True)
class Assignment:
    ticket_number: int
    is_repeat: bool


@dataclass(frozen=True)
class ResultEntry:
    group_name: str
    last_name: str
    first_name: str
    ticket_number: int
    created_at: str
    repeat: str

    @property
    def full_name(self) -> str:
        return f"{self.last_name} {self.first_name}"


@dataclass(frozen=True)
class ResultStats:
    total_generations: int
    unique_students: int
    repeat_generations: int
    top_ticket: int | None


class ResultsStore:
    def __init__(self, path: str | Path) -> None:
        self.path = Path(path)

    def record_assignment(
        self,
        group_name: str,
        student: Student,
        valid_ticket_numbers: Sequence[int],
        now: Callable[[], datetime] = datetime.now,
        rng: random.Random | None = None,
    ) -> Assignment:
        if not valid_ticket_numbers:
            raise ValueError("Нет доступных билетов для генерации.")

        random_source = rng or random
        workbook = self._open_or_create_workbook()
        sheet = workbook.active

        existing_ticket = self._find_first_ticket(sheet, group_name, student)
        is_repeat = existing_ticket is not None
        ticket_number = existing_ticket if existing_ticket is not None else random_source.choice(list(valid_ticket_numbers))

        sheet.append(
            [
                group_name,
                student.last_name,
                student.first_name,
                ticket_number,
                now().strftime("%Y-%m-%d %H:%M:%S"),
                "да" if is_repeat else "нет",
            ]
        )

        try:
            workbook.save(self.path)
        except PermissionError as error:
            raise ResultsWriteError(
                f"Не удалось записать журнал результатов. Возможно, файл открыт в Excel: {self.path}"
            ) from error
        except OSError as error:
            raise ResultsWriteError(f"Не удалось сохранить журнал результатов: {self.path}") from error
        finally:
            workbook.close()

        return Assignment(ticket_number=ticket_number, is_repeat=is_repeat)

    def recent_entries(self, limit: int = 10) -> list[ResultEntry]:
        if not self.path.exists():
            return []

        workbook = load_workbook(self.path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            entries = [_row_to_entry(row) for row in sheet.iter_rows(min_row=2, values_only=True)]
            entries = [entry for entry in entries if entry is not None]
            return list(reversed(entries[-limit:]))
        finally:
            workbook.close()

    def stats(self) -> ResultStats:
        if not self.path.exists():
            return ResultStats(total_generations=0, unique_students=0, repeat_generations=0, top_ticket=None)

        workbook = load_workbook(self.path, read_only=True, data_only=True)
        try:
            sheet = workbook.active
            entries = [_row_to_entry(row) for row in sheet.iter_rows(min_row=2, values_only=True)]
            entries = [entry for entry in entries if entry is not None]
        finally:
            workbook.close()

        unique_students = {(entry.group_name, entry.last_name, entry.first_name) for entry in entries}
        repeats = sum(1 for entry in entries if entry.repeat == "да")
        ticket_counts = Counter(entry.ticket_number for entry in entries)
        top_ticket = ticket_counts.most_common(1)[0][0] if ticket_counts else None
        return ResultStats(
            total_generations=len(entries),
            unique_students=len(unique_students),
            repeat_generations=repeats,
            top_ticket=top_ticket,
        )

    def _open_or_create_workbook(self) -> Workbook:
        if self.path.exists():
            workbook = load_workbook(self.path)
            sheet = workbook.active
            if sheet.max_row == 0 or [sheet.cell(1, index).value for index in range(1, 7)] != list(HEADERS):
                if sheet.max_row == 1 and all(sheet.cell(1, index).value is None for index in range(1, 7)):
                    for index, header in enumerate(HEADERS, start=1):
                        sheet.cell(1, index).value = header
            return workbook

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Results"
        sheet.append(list(HEADERS))
        return workbook

    @staticmethod
    def _find_first_ticket(sheet, group_name: str, student: Student) -> int | None:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            group, last_name, first_name, ticket_number = row[:4]
            if (
                str(group).strip() == group_name
                and str(last_name).strip() == student.last_name
                and str(first_name).strip() == student.first_name
            ):
                return int(ticket_number)
        return None


def _row_to_entry(row: tuple[object, ...]) -> ResultEntry | None:
    if len(row) < 6 or not row[0] or not row[1] or not row[2] or row[3] is None:
        return None
    return ResultEntry(
        group_name=str(row[0]).strip(),
        last_name=str(row[1]).strip(),
        first_name=str(row[2]).strip(),
        ticket_number=int(row[3]),
        created_at="" if row[4] is None else str(row[4]).strip(),
        repeat="" if row[5] is None else str(row[5]).strip(),
    )
