from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class NormalizedFact:
    family: str
    source_file: str
    source_row: int | None
    entity_num: str
    partner_num: str
    account_code: str
    amount: float
    direction: str
    raw_entity: str
    raw_partner: str
    raw_account: str
    meta: dict[str, Any] = field(default_factory=dict)


@dataclass
class PairOwner:
    pair_group: tuple[str, str]
    canonical_entity: str
    canonical_partner: str
    owner_type: str
    owner_reason: str
    icm_row_num: int | None
    display_entity: str
    display_partner: str


@dataclass
class RoutedCell:
    row_key: tuple[str, str]
    block: str
    side: str
    account_code: str
    amount: float
    source_fact_ids: list[str] = field(default_factory=list)


@dataclass
class LayoutColumn:
    code: str
    description: str
    series: str
    tag: str


@dataclass
class LayoutSpec:
    ent_cols: list[LayoutColumn]
    par_cols: list[LayoutColumn]
    plug_code: str | None = None
    plug_label_base: str | None = None


@dataclass
class OutputRow:
    row_key: tuple[str, str]
    pair_group: tuple[str, str]
    display_entity: str
    display_partner: str
    source_row_num: int | None
    is_synthetic: bool
    owner_reason: str
    owner_type: str


@dataclass
class BaseRow:
    pair_key: tuple[str, str]
    pair_group: tuple[str, str]
    row_num: int | None
    row_direction: str
    entity_label: str
    partner_label: str
    source_file: str
    meta: dict[str, Any] = field(default_factory=dict)


@dataclass
class WorkbookSource:
    filepath: str
    kind: str
    sheet_name: str
    header_row: int
    data_start: int
    has_total_column: bool = False
    total_column: int | None = None


@dataclass
class FactBuildResult:
    fact_registry: dict[str, NormalizedFact]
    facts_icm: list[str]
    facts_parent: list[str]
    facts_contrib: list[str]
    facts_plug_source: list[str]
    facts_ic_elim: list[str]
    base_rows: list[BaseRow]
    layout: LayoutSpec
    label_maps: dict[str, Any]
    diagnostics: dict[str, list[dict[str, Any]]]
    plug_code: str | None
    elim_codes: set[str]
    sources: dict[str, WorkbookSource | None]


@dataclass
class PipelineResult:
    fact_build: FactBuildResult
    pair_registry: dict[tuple[str, str], PairOwner]
    row_registry: list[OutputRow]
    base_value_map: dict[tuple[tuple[str, str], str, str], float]
    cell_ledger: dict[tuple[tuple[str, str], str, str, str], RoutedCell]
    fact_assignment_log: dict[str, dict[str, Any]]
    diagnostics: dict[str, list[dict[str, Any]]]
    plug_reconciliation_log: list[dict[str, Any]]
