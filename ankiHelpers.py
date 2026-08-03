"""Reusable helpers for adding Excel vocabulary to a local Anki collection."""

from __future__ import annotations

import os
from pathlib import Path

import anki
from anki.collection import Collection

COLLECTION_PATH = Path(os.environ["APPDATA"]) / "Anki2" / "User 1" / "collection.anki2"
NOTE_TYPE_NAME = "Three Sided"


def list_decks() -> list[str]:
    """Return the names of all decks in the local Anki collection.

    Anki must be fully closed before calling this function, including while media
    syncing.
    """
    if not COLLECTION_PATH.is_file():
        raise FileNotFoundError(f"Anki collection not found: {COLLECTION_PATH}")

    collection = Collection(str(COLLECTION_PATH))
    try:
        return sorted(deck.name for deck in collection.decks.all_names_and_ids())
    finally:
        collection.close()


def add_three_sided_card(deck_name: str, english: str, chinese: str, pinyin: str) -> int:
    """Add a Three Sided note to an existing deck and return its note ID.

    Anki must be fully closed before calling this function, including while media
    syncing. The target note type must contain fields named English, Chinese, and
    Pinyin (capitalization does not matter).
    """
    if not COLLECTION_PATH.is_file():
        raise FileNotFoundError(f"Anki collection not found: {COLLECTION_PATH}")

    collection = Collection(str(COLLECTION_PATH))
    try:
        deck_id = collection.decks.id(deck_name, create=False)
        if deck_id is None:
            raise ValueError(f"Anki deck does not exist: {deck_name!r}")

        note_type = collection.models.by_name(NOTE_TYPE_NAME)
        if note_type is None:
            raise ValueError(f"Anki note type does not exist: {NOTE_TYPE_NAME!r}")

        note = collection.new_note(note_type)
        fields_by_name = {field.casefold(): field for field in note.keys()}
        values = {"english": english, "chinese": chinese, "pinyin": pinyin}
        missing_fields = [name.title() for name in values if name not in fields_by_name]
        if missing_fields:
            raise ValueError(
                f"The {NOTE_TYPE_NAME!r} note type is missing field(s): "
                + ", ".join(missing_fields)
            )

        for field_name, value in values.items():
            note[fields_by_name[field_name]] = value

        collection.add_note(note, deck_id)
        collection.save()
        return int(note.id)
    finally:
        collection.close()
