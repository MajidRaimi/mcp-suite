from __future__ import annotations

from datetime import UTC, date

import pytest
from conftest import TODAY

from office_mcp.shared.clock import list_label_date, message_time
from office_mcp.shared.text import normalize


@pytest.mark.parametrize(
    ("label", "expected"),
    [
        ("4:09 AM", date(2026, 10, 20)),
        ("16:05", date(2026, 10, 20)),
        ("Yesterday", date(2026, 10, 19)),
        ("Sat", date(2026, 10, 17)),
        ("Tuesday", date(2026, 10, 13)),
        ("10/19", date(2026, 10, 19)),
        ("12/28", date(2025, 12, 28)),
        ("12/28/2025", date(2025, 12, 28)),
        ("2/30", None),
        ("", None),
        ("soon", None),
    ],
)
def test_list_labels(label: str, expected: date | None) -> None:
    assert list_label_date(label, TODAY) == expected


def test_message_time_needs_epoch_millis() -> None:
    assert message_time("1790000000000", UTC) is not None
    assert message_time("card1", UTC) is None
    assert message_time("123", UTC) is None


def test_normalize_folds_case_marks_and_arabic_letters() -> None:
    assert normalize("  Jane   DOE ") == "jane doe"
    assert normalize("سارة أحمد") == normalize("ساره احمد")
    assert normalize("José") == "jose"
