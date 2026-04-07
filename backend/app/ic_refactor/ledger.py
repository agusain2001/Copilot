from __future__ import annotations

from .diagnostics import record_diagnostic
from .models import RoutedCell


def _group_key(entity_num: str, partner_num: str) -> tuple[str, str]:
    return tuple(sorted((entity_num, partner_num)))


def resolve_fact_side(fact, pair_owner):
    canonical = (pair_owner.canonical_entity, pair_owner.canonical_partner)
    reversed_pair = (pair_owner.canonical_partner, pair_owner.canonical_entity)
    fact_pair = (fact.entity_num, fact.partner_num)

    if fact.meta.get("derived_from") == "parent":
        if fact_pair in (canonical, reversed_pair):
            return "plug_parent"
        return None
    if fact.meta.get("derived_from") == "contrib":
        if fact_pair in (canonical, reversed_pair):
            return "plug_contrib"
        return None
    if fact.meta.get("derived_from") == "ic_elim_total":
        if fact_pair in (canonical, reversed_pair):
            return "plug_section"
        return None

    if fact_pair == canonical:
        return "entity_side"
    if fact_pair == reversed_pair:
        return "partner_side"
    return None


def route_fact_to_cell(fact_id, fact, pair_registry, layout, diagnostics, row_lookup=None):
    pair_group = _group_key(fact.entity_num, fact.partner_num)
    exact_row = row_lookup.get((fact.entity_num, fact.partner_num)) if row_lookup else None
    if exact_row is not None and exact_row.owner_reason == "non_e_bidirectional_split":
        side = "plug_parent" if fact.meta.get("derived_from") == "parent" else (
            "plug_contrib" if fact.meta.get("derived_from") == "contrib" else (
                "plug_section" if fact.meta.get("derived_from") == "ic_elim_total" else "entity_side"
            )
        )
        return exact_row.row_key, ("plug" if side.startswith("plug_") else fact.family), side, fact.account_code

    pair_owner = pair_registry.get(pair_group)
    if pair_owner is None:
        record_diagnostic(
            diagnostics,
            "unmatched_facts",
            "missing_pair_owner",
            source_file=fact.source_file,
            source_row=fact.source_row,
            raw_entity=fact.raw_entity,
            raw_partner=fact.raw_partner,
            normalized_entity=fact.entity_num,
            normalized_partner=fact.partner_num,
            account=fact.account_code,
            amount=fact.amount,
            extra={"fact_id": fact_id, "family": fact.family},
        )
        return None

    side = resolve_fact_side(fact, pair_owner)
    if side is None:
        record_diagnostic(
            diagnostics,
            "unmatched_facts",
            "pair_orientation_mismatch",
            source_file=fact.source_file,
            source_row=fact.source_row,
            raw_entity=fact.raw_entity,
            raw_partner=fact.raw_partner,
            normalized_entity=fact.entity_num,
            normalized_partner=fact.partner_num,
            account=fact.account_code,
            amount=fact.amount,
            extra={"fact_id": fact_id, "family": fact.family},
        )
        return None

    ent_codes = {column.code for column in layout.ent_cols}
    par_codes = {column.code for column in layout.par_cols}

    if side == "entity_side" and fact.account_code not in ent_codes:
        record_diagnostic(
            diagnostics,
            "unmatched_facts",
            "account_not_in_entity_layout",
            source_file=fact.source_file,
            source_row=fact.source_row,
            raw_entity=fact.raw_entity,
            raw_partner=fact.raw_partner,
            normalized_entity=fact.entity_num,
            normalized_partner=fact.partner_num,
            account=fact.account_code,
            amount=fact.amount,
            extra={"fact_id": fact_id, "family": fact.family},
        )
        return None
    if side == "partner_side" and fact.account_code not in par_codes:
        record_diagnostic(
            diagnostics,
            "unmatched_facts",
            "account_not_in_partner_layout",
            source_file=fact.source_file,
            source_row=fact.source_row,
            raw_entity=fact.raw_entity,
            raw_partner=fact.raw_partner,
            normalized_entity=fact.entity_num,
            normalized_partner=fact.partner_num,
            account=fact.account_code,
            amount=fact.amount,
            extra={"fact_id": fact_id, "family": fact.family},
        )
        return None

    block = "plug" if side.startswith("plug_") else fact.family
    if block == "derived_plug":
        block = "plug"
    if block == "ic_elim":
        block = "plug"
    row_key = (pair_owner.canonical_entity, pair_owner.canonical_partner)
    return row_key, block, side, fact.account_code


def build_base_value_map(base_fact_ids, fact_registry):
    base_value_map = {}
    for fact_id in base_fact_ids:
        fact = fact_registry[fact_id]
        source_row = fact.source_row
        if source_row is None:
            continue
        column_tag = fact.meta.get("column_tag")
        side = "entity_side" if column_tag == "Entity" else "partner_side"
        key = (source_row, side, fact.account_code)
        base_value_map[key] = base_value_map.get(key, 0.0) + fact.amount
    return base_value_map


def build_cell_ledger(fact_ids, fact_registry, pair_registry, layout, diagnostics):
    ledger: dict[tuple[tuple[str, str], str, str, str], RoutedCell] = {}
    fact_assignment_log: dict[str, dict] = {}
    row_lookup = {}
    if "row_registry" in diagnostics:
        for row in diagnostics["row_registry"]:
            row_lookup[row.row_key] = row

    for fact_id in fact_ids:
        fact = fact_registry[fact_id]
        route = route_fact_to_cell(fact_id, fact, pair_registry, layout, diagnostics, row_lookup=row_lookup)
        if route is None:
            fact_assignment_log[fact_id] = {
                "source_file": fact.source_file,
                "source_row": fact.source_row,
                "raw_entity": fact.raw_entity,
                "raw_partner": fact.raw_partner,
                "entity_num": fact.entity_num,
                "partner_num": fact.partner_num,
                "account_code": fact.account_code,
                "amount": fact.amount,
                "reason": "diagnosed_unmatched",
                "destination": None,
            }
            continue

        row_key, block, side, account_code = route
        key = (row_key, block, side, account_code)
        routed = ledger.get(key)
        if routed is None:
            routed = RoutedCell(
                row_key=row_key,
                block=block,
                side=side,
                account_code=account_code,
                amount=0.0,
                source_fact_ids=[],
            )
            ledger[key] = routed
        routed.amount += fact.amount
        routed.source_fact_ids.append(fact_id)
        fact_assignment_log[fact_id] = {
            "source_file": fact.source_file,
            "source_row": fact.source_row,
            "raw_entity": fact.raw_entity,
            "raw_partner": fact.raw_partner,
            "entity_num": fact.entity_num,
            "partner_num": fact.partner_num,
            "account_code": fact.account_code,
            "amount": fact.amount,
            "reason": "routed",
            "destination": {"row_key": row_key, "block": block, "side": side, "account_code": account_code},
        }

    return ledger, fact_assignment_log
