from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document


TICKET_HEADER_RE = re.compile(r"^\s*Билет\s+(\d+)\s*$", re.IGNORECASE)
QUESTION_RE = re.compile(r"^\s*\d+[\.)]\s*(.+?)\s*$")


@dataclass(frozen=True)
class Ticket:
    number: int
    questions: tuple[str, str, str]


def parse_tickets(path: str | Path, logger: logging.Logger | None = None) -> dict[int, Ticket]:
    """Read tickets from a DOCX file and return only valid tickets.

    A valid ticket has a header like "Билет 1" and at least three numbered
    question paragraphs. Extra questions are ignored because the UI shows three.
    """

    log = logger or logging.getLogger(__name__)
    path = Path(path)

    if not path.exists():
        raise FileNotFoundError(f"Файл с билетами не найден: {path}")

    document = Document(path)
    tickets: dict[int, Ticket] = {}
    current_number: int | None = None
    current_questions: list[str] = []

    def finish_current() -> None:
        nonlocal current_number, current_questions
        if current_number is None:
            return
        if len(current_questions) < 3:
            log.warning(
                "Билет %s пропущен: найдено вопросов %s, требуется минимум 3",
                current_number,
                len(current_questions),
            )
        else:
            tickets[current_number] = Ticket(
                number=current_number,
                questions=tuple(current_questions[:3]),  # type: ignore[arg-type]
            )
        current_number = None
        current_questions = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        header_match = TICKET_HEADER_RE.match(text)
        if header_match:
            finish_current()
            current_number = int(header_match.group(1))
            current_questions = []
            continue

        if text.lower().startswith("билет"):
            finish_current()
            log.warning("Заголовок билета не распознан и пропущен: %r", text)
            continue

        question_match = QUESTION_RE.match(text)
        if question_match and current_number is not None:
            current_questions.append(question_match.group(1))
        elif question_match:
            log.warning("Вопрос без заголовка билета пропущен: %r", text)

    finish_current()
    return tickets
