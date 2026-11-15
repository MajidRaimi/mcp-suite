from __future__ import annotations

from office_mcp.shared.text import normalize


def test_tatweel_is_removed() -> None:
    assert normalize("كتــاب") == normalize("كتاب")


def test_hamza_seats_fold() -> None:
    assert normalize("مؤتمر") == normalize("موتمر")
    assert normalize("مسائل") == normalize("مسايل")


def test_tashkeel_and_alef_forms_still_fold() -> None:
    assert normalize("أَهْلًا") == normalize("اهلا")


def test_latin_text_is_casefolded_and_spaced() -> None:
    assert normalize("  Project   ATLAS  ") == "project atlas"
