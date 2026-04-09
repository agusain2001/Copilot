from __future__ import annotations

from collections import defaultdict

from .models import NormalizedFact


def derive_plug_facts(fact_ids, fact_registry, elim_codes: set[str], plug_code: str | None, stream: str) -> list[NormalizedFact]:
    if not plug_code:
        return []

    derived = []
    for fact_id in fact_ids:
        fact = fact_registry[fact_id]
        # Plug columns must reflect only the configured plug-account postings.
        if fact.account_code != plug_code:
            continue
        derived.append(
            NormalizedFact(
                family="derived_plug",
                source_file=fact.source_file,
                source_row=fact.source_row,
                entity_num=fact.entity_num,
                partner_num=fact.partner_num,
                account_code=plug_code,
                amount=fact.amount,
                direction=fact.direction,
                raw_entity=fact.raw_entity,
                raw_partner=fact.raw_partner,
                raw_account=fact.raw_account,
                meta={
                    "derived_from": stream,
                    "source_fact_ids": [fact_id],
                    "origin_account_code": fact.account_code,
                },
            )
        )
    return derived


def build_plug_section_facts(fact_ids, fact_registry, plug_code: str | None) -> list[NormalizedFact]:
    if not plug_code:
        return []

    derived = []
    for fact_id in fact_ids:
        fact = fact_registry[fact_id]
        if fact.family != "ic_elim":
            continue
        if fact.account_code != plug_code:
            continue
        if fact.meta.get("column_role") != "total":
            continue
        derived.append(
            NormalizedFact(
                family="derived_plug",
                source_file=fact.source_file,
                source_row=fact.source_row,
                entity_num=fact.entity_num,
                partner_num=fact.partner_num,
                account_code=plug_code,
                amount=fact.amount,
                direction=fact.direction,
                raw_entity=fact.raw_entity,
                raw_partner=fact.raw_partner,
                raw_account=fact.raw_account,
                meta={
                    "derived_from": "ic_elim_total",
                    "source_fact_ids": [fact_id],
                },
            )
        )
    return derived


def build_plug_reconciliation_log(
    fact_registry,
    parent_plug_ids: list[str],
    contrib_plug_ids: list[str],
    standalone_plug_ids: list[str],
    plug_source_ids: list[str],
) -> list[dict]:
    totals = defaultdict(lambda: {"parent": 0.0, "contrib": 0.0, "standalone": 0.0, "plug_source": 0.0})
    for label, fact_ids in (
        ("parent", parent_plug_ids),
        ("contrib", contrib_plug_ids),
        ("standalone", standalone_plug_ids),
        ("plug_source", plug_source_ids),
    ):
        for fact_id in fact_ids:
            fact = fact_registry[fact_id]
            pair_group = tuple(sorted((fact.entity_num, fact.partner_num)))
            totals[pair_group][label] += fact.amount

    rows = []
    for pair_group, values in sorted(totals.items()):
        rows.append(
            {
                "source_file": "",
                "source_row": None,
                "raw_entity": pair_group[0],
                "raw_partner": pair_group[1],
                "normalized_entity": pair_group[0],
                "normalized_partner": pair_group[1],
                "account": "",
                "amount": values["standalone"],
                "reason": "plug_reconciliation",
                "parent_derived": values["parent"],
                "contrib_derived": values["contrib"],
                "standalone_total": values["standalone"],
                "plug_source_total": values["plug_source"],
                "variance": values["standalone"] - values["plug_source"],
            }
        )
    return rows
