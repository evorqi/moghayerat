"""In-memory files for the local app. Nothing is sent off this computer."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Uploaded:
    id: str
    filename: str
    sheets: dict[str, list[list]]


@dataclass
class Export:
    path: Path
    filename: str


files: dict[str, Uploaded] = {}
exports: dict[str, Export] = {}


def reset() -> None:
    files.clear()
    for item in exports.values():
        item.path.unlink(missing_ok=True)
    exports.clear()
