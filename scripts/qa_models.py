from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class QaSample:
    name: str
    source_path: Path
    output_path: Path
    expected_tables: int
    expected_merged_cells: int


@dataclass(frozen=True, slots=True)
class QaRunReport:
    work_dir: Path
    samples: tuple[QaSample, ...]
    issues: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not self.issues
