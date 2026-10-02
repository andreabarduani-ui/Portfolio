import importlib.util
from email.message import EmailMessage

import pytest

from pathlib import Path


REPO = Path(__file__).resolve().parents[1]
MODULE_PATH = REPO / "projects/fnc-email-reminder/email_reminder_automation.py"
SPEC = importlib.util.spec_from_file_location("fnc_email_reminder", MODULE_PATH)
assert SPEC is not None and SPEC.loader is not None
reminder = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(reminder)


def test_load_csv_and_build_personalized_message(tmp_path):
    source = tmp_path / "partecipanti.csv"
    source.write_text(
        "Nome;Email;Corsi da finire\n"
        "Mario Rossi;mario@example.com;Corso A | Corso <B>\n",
        encoding="utf-8",
    )

    participants = reminder.load_participants(source)
    assert participants == [
        reminder.Participant("Mario Rossi", "mario@example.com", ("Corso A", "Corso <B>"))
    ]

    message = reminder.build_message(participants[0], "sender@example.com")
    assert isinstance(message, EmailMessage)
    assert message["To"] == "mario@example.com"
    assert message["Subject"] == reminder.SUBJECT
    assert "Gentile Mario Rossi" in message.get_body(("plain",)).get_content()
    assert "Corso &lt;B&gt;" in message.get_body(("html",)).get_content()
    assert reminder.VADEMECUM_URL in message.get_body(("plain",)).get_content()


@pytest.mark.parametrize(
    ("row", "expected"),
    [
        ("Mario;not-an-email;Corso A\n", "email non valida"),
        ("Mario;mario@example.com;\n", "nessun corso indicato"),
        (
            "Mario;mario@example.com;Corso A\n"
            "Luigi;MARIO@example.com;Corso B\n",
            "email duplicata",
        ),
    ],
)
def test_invalid_rows_block_processing(tmp_path, row, expected):
    source = tmp_path / "invalid.csv"
    source.write_text("Nome;Email;Corsi\n" + row, encoding="utf-8")

    with pytest.raises(ValueError, match=expected):
        reminder.load_participants(source)
