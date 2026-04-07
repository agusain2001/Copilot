from __future__ import annotations

from .diagnostics import diagnostics_totals
from .facts import build_all_facts
from .ledger import build_base_value_map, build_cell_ledger
from .models import PipelineResult
from .pairs import build_pair_registry
from .plug import build_plug_reconciliation_log, build_plug_section_facts, derive_plug_facts
from .writer import write_output_v2


def _register_derived_facts(fact_registry, derived_facts):
    ids = []
    for fact in derived_facts:
        fact_id = f"fact_{len(fact_registry) + 1:06d}"
        fact_registry[fact_id] = fact
        ids.append(fact_id)
    return ids


def _row_relevant_fact_ids(fact_build):
    layout_codes = {column.code for column in fact_build.layout.ent_cols} | {column.code for column in fact_build.layout.par_cols}
    relevant = []
    for fact_id in fact_build.facts_parent + fact_build.facts_contrib:
        fact = fact_build.fact_registry[fact_id]
        if fact.account_code in layout_codes or fact.account_code in fact_build.elim_codes:
            relevant.append(fact_id)
    for fact_id in fact_build.facts_ic_elim:
        fact = fact_build.fact_registry[fact_id]
        if fact.account_code in layout_codes or fact.account_code == fact_build.plug_code:
            relevant.append(fact_id)
    return relevant


def run_pipeline_v2(icm_path, journal_paths, report_inputs_path=None) -> PipelineResult:
    fact_build = build_all_facts(icm_path, journal_paths, report_inputs_path=report_inputs_path)
    diagnostics = fact_build.diagnostics
    fact_registry = fact_build.fact_registry

    pair_fact_ids = _row_relevant_fact_ids(fact_build)
    pair_registry, row_registry = build_pair_registry(
        fact_build.base_rows,
        pair_fact_ids,
        fact_registry,
        fact_build.label_maps,
        diagnostics,
    )

    parent_plug = derive_plug_facts(fact_build.facts_parent, fact_registry, fact_build.elim_codes, fact_build.plug_code, "parent")
    contrib_plug = derive_plug_facts(fact_build.facts_contrib, fact_registry, fact_build.elim_codes, fact_build.plug_code, "contrib")
    standalone_plug = build_plug_section_facts(fact_build.facts_ic_elim, fact_registry, fact_build.plug_code)

    parent_plug_ids = _register_derived_facts(fact_registry, parent_plug)
    contrib_plug_ids = _register_derived_facts(fact_registry, contrib_plug)
    standalone_plug_ids = _register_derived_facts(fact_registry, standalone_plug)

    base_value_map = build_base_value_map(
        fact_build.facts_icm,
        fact_registry,
    )
    ledger_fact_ids = fact_build.facts_parent + fact_build.facts_contrib + parent_plug_ids + contrib_plug_ids + standalone_plug_ids
    diagnostics["row_registry"] = row_registry
    cell_ledger, fact_assignment_log = build_cell_ledger(
        ledger_fact_ids,
        fact_registry,
        pair_registry,
        fact_build.layout,
        diagnostics,
    )

    for fact_id in fact_build.facts_plug_source:
        fact = fact_registry[fact_id]
        fact_assignment_log.setdefault(
            fact_id,
            {
                "source_file": fact.source_file,
                "source_row": fact.source_row,
                "raw_entity": fact.raw_entity,
                "raw_partner": fact.raw_partner,
                "entity_num": fact.entity_num,
                "partner_num": fact.partner_num,
                "account_code": fact.account_code,
                "amount": fact.amount,
                "reason": "plug_source_reconciliation_only",
                "destination": None,
            },
        )

    plug_reconciliation_log = build_plug_reconciliation_log(
        fact_registry,
        parent_plug_ids,
        contrib_plug_ids,
        standalone_plug_ids,
        fact_build.facts_plug_source,
    )
    diagnostics["summary"] = [{"reason": key, "amount": value} for key, value in diagnostics_totals(diagnostics).items()]

    return PipelineResult(
        fact_build=fact_build,
        pair_registry=pair_registry,
        row_registry=row_registry,
        base_value_map=base_value_map,
        cell_ledger=cell_ledger,
        fact_assignment_log=fact_assignment_log,
        diagnostics=diagnostics,
        plug_reconciliation_log=plug_reconciliation_log,
    )


def process_icm_report_v2(icm_path, journal_paths, output_path, report_inputs_path=None):
    result = run_pipeline_v2(icm_path, journal_paths, report_inputs_path=report_inputs_path)
    write_output_v2(
        result.row_registry,
        result.cell_ledger,
        result.base_value_map,
        output_path,
        result.fact_build.layout,
        result.diagnostics,
        result.plug_reconciliation_log,
        result.fact_assignment_log,
    )
    return output_path
