"""Helpers for loading and looking up entries in the bundled CC-CEDICT dictionary."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
import re
from types import MappingProxyType
from typing import Mapping


_CEDICT_ENTRY = re.compile(
    r"^(?P<traditional>\S+)\s+(?P<simplified>\S+)\s+\[(?P<pinyin>[^\]]+)\]\s+/(?P<definitions>.*)/$"
)


@dataclass(frozen=True)
class DictionaryEntry:
    """One CC-CEDICT entry, retaining both character forms."""

    traditional: str
    simplified: str
    pinyin: str
    definitions: tuple[str, ...]


DictionaryIndex = Mapping[str, tuple[DictionaryEntry, ...]]


def dictionary_path() -> Path:
    """Return the path to the CC-CEDICT file bundled with the application."""
    return Path(__file__).with_name("data") / "cedict_ts.u8"


def parse_cedict_line(line: str) -> DictionaryEntry | None:
    """Parse one non-comment CC-CEDICT line, or return ``None`` for blanks/comments."""
    line = line.strip()
    if not line or line.startswith("#"):
        return None

    match = _CEDICT_ENTRY.match(line)
    if match is None:
        raise ValueError(f"Invalid CC-CEDICT entry: {line!r}")

    definitions = tuple(part for part in match.group("definitions").split("/") if part)
    return DictionaryEntry(
        traditional=match.group("traditional"),
        simplified=match.group("simplified"),
        pinyin=match.group("pinyin"),
        definitions=definitions,
    )


@lru_cache(maxsize=None)
def _load_dictionary_cached(path: Path) -> DictionaryIndex:
    entries_by_word: defaultdict[str, list[DictionaryEntry]] = defaultdict(list)
    with path.open(encoding="utf-8") as dictionary_file:
        for line_number, line in enumerate(dictionary_file, start=1):
            try:
                entry = parse_cedict_line(line)
            except ValueError as error:
                raise ValueError(f"Could not parse {path.name} line {line_number}: {error}") from error
            if entry is None:
                continue
            entries_by_word[entry.simplified].append(entry)
            if entry.traditional != entry.simplified:
                entries_by_word[entry.traditional].append(entry)

    return MappingProxyType({word: tuple(entries) for word, entries in entries_by_word.items()})


def load_dictionary(path: Path | None = None) -> DictionaryIndex:
    """Load and cache CC-CEDICT, indexed by simplified and traditional headword."""
    resolved_path = (path or dictionary_path()).resolve()
    if not resolved_path.is_file():
        raise FileNotFoundError(f"CC-CEDICT file not found: {resolved_path}")
    return _load_dictionary_cached(resolved_path)


def lookup(word: str, index: DictionaryIndex | None = None) -> tuple[DictionaryEntry, ...]:
    """Return every CC-CEDICT entry matching a simplified or traditional headword."""
    query = word.strip()
    if not query:
        return ()
    dictionary = index if index is not None else load_dictionary()
    return dictionary.get(query, ())


def format_definition(entry: DictionaryEntry) -> str:
    """Format definitions for display while leaving the individual meanings selectable."""
    return " / ".join(entry.definitions)
