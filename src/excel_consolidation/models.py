"""Plain data models shared by the generator and the consolidator."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class FileManifest:
    """The generator's own ground truth for one workbook it wrote. Tests
    compare the consolidator's actual output against this, not against
    itself, so a bug that's consistent between generator and consolidator
    can't hide."""

    filename: str
    kind: str  # "normal" | "corrupt" | "wrong_structure"
    rows_in: int
    valid_unique: int
    duplicates: int
    rejected: dict[str, int] = field(default_factory=dict)

    @property
    def rejected_total(self) -> int:
        return sum(self.rejected.values())

    def reconciles(self) -> bool:
        return self.rows_in == self.valid_unique + self.duplicates + self.rejected_total


@dataclass
class RejectedRow:
    file: str
    row_ref: int
    reason: str
    detail: str = ""


@dataclass
class DuplicateRow:
    file: str
    row_ref: int
    order_id: str
    first_seen_file: str
    first_seen_row: int


@dataclass
class FileResult:
    filename: str
    sheet_name: str | None
    status: str  # "ok" | "exception" | "unreadable" | "unmappable_structure"
    rows_in: int
    rows_out: int
    rejected: int
    duplicates: int
    detail: str = ""


@dataclass
class ConsolidationResult:
    data_rows: list[dict]
    file_results: list[FileResult]
    rejected_rows: list[RejectedRow]
    duplicate_rows: list[DuplicateRow]
    change_log: dict[str, int]

    @property
    def totals(self) -> dict[str, int]:
        return {
            "rows_in": sum(fr.rows_in for fr in self.file_results),
            "rows_out": sum(fr.rows_out for fr in self.file_results),
            "rejected": sum(fr.rejected for fr in self.file_results),
            "duplicates": sum(fr.duplicates for fr in self.file_results),
        }

    def reconciles(self) -> bool:
        t = self.totals
        return t["rows_in"] == t["rows_out"] + t["rejected"] + t["duplicates"]
