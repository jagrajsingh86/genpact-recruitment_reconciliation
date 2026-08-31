"""Synthetic library discovery and chronological reconciliation orchestration."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re

from .io_excel import load_inputs
from .reconcile import ReconciliationResult, reconcile


DATE_FOLDER_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")


class LibraryContractError(ValueError):
    """Raised when the bundled library does not match its folder/file contract."""


@dataclass(frozen=True)
class LibraryExtract:
    run_date: str
    requisition_path: Path
    position_path: Path


def discover_library_extracts(library_root: str | Path) -> tuple[LibraryExtract, ...]:
    root = Path(library_root)
    extracts_root = root / "extracts"
    if not extracts_root.is_dir():
        raise LibraryContractError(f"Synthetic library extracts folder not found: {extracts_root}")

    dated_folders: list[Path] = []
    for folder in extracts_root.iterdir():
        if not folder.is_dir() or not DATE_FOLDER_PATTERN.fullmatch(folder.name):
            continue
        try:
            date.fromisoformat(folder.name)
        except ValueError as error:
            raise LibraryContractError(f"Invalid dated extract folder: {folder.name}") from error
        dated_folders.append(folder)

    if not dated_folders:
        raise LibraryContractError(f"No YYYY-MM-DD extract folders found in: {extracts_root}")

    discovered: list[LibraryExtract] = []
    for folder in sorted(dated_folders, key=lambda item: item.name):
        requisitions = sorted(folder.glob("RCM_Requisition_Report_*.xlsx"))
        positions = sorted(folder.glob("EC_Position_Report_*.xlsx"))
        if len(requisitions) != 1:
            raise LibraryContractError(
                f"{folder.name} must contain exactly one RCM_Requisition_Report_*.xlsx; "
                f"found {len(requisitions)}"
            )
        if len(positions) != 1:
            raise LibraryContractError(
                f"{folder.name} must contain exactly one EC_Position_Report_*.xlsx; "
                f"found {len(positions)}"
            )
        discovered.append(
            LibraryExtract(
                run_date=folder.name,
                requisition_path=requisitions[0],
                position_path=positions[0],
            )
        )
    return tuple(discovered)


def run_library(
    library_root: str | Path,
    *,
    apply_normalisers: bool = True,
) -> dict[str, ReconciliationResult]:
    """Run every bundled date in chronological order with chained findings."""

    root = Path(library_root)
    config_path = root / "config" / "config_field_map.xlsx"
    if not config_path.is_file():
        raise LibraryContractError(f"Synthetic library config not found: {config_path}")

    results: dict[str, ReconciliationResult] = {}
    previous_findings = ()
    for extract in discover_library_extracts(root):
        loaded = load_inputs(extract.requisition_path, extract.position_path, config_path)
        result = reconcile(
            loaded.requisitions,
            loaded.positions,
            loaded.config,
            extract.run_date,
            previous_findings,
            apply_normalisers=apply_normalisers,
        )
        results[extract.run_date] = result
        previous_findings = result.findings
    return results
