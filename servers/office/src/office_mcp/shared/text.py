from __future__ import annotations

import unicodedata

ARABIC_FOLDS = str.maketrans(
    {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ٱ": "ا",
        "ة": "ه",
        "ى": "ي",
        "ؤ": "و",
        "ئ": "ي",
        "\u0640": None,
    }
)


def normalize(text: str) -> str:
    decomposed = unicodedata.normalize("NFKD", text)
    kept = (
        char
        for char in decomposed
        if not unicodedata.combining(char) and unicodedata.category(char) != "Cf"
    )
    folded = "".join(kept).translate(ARABIC_FOLDS)
    return " ".join(folded.casefold().split())
