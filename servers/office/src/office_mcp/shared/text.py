from __future__ import annotations

import unicodedata

ARABIC_FOLDS = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ة": "ه", "ى": "ي"})


def normalize(text: str) -> str:
    decomposed = unicodedata.normalize("NFKD", text)
    kept = (
        char
        for char in decomposed
        if not unicodedata.combining(char) and unicodedata.category(char) != "Cf"
    )
    return " ".join("".join(kept).translate(ARABIC_FOLDS).casefold().split())
