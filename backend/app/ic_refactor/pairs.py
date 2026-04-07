from __future__ import annotations

from collections import defaultdict

from .config import PAIR_OWNER_OVERRIDES
from .diagnostics import record_diagnostic
from .models import OutputRow, PairOwner


def _group_key(entity_num: str, partner_num: str) -> tuple[str, str]:
    return tuple(sorted((entity_num, partner_num)))


def _classify_label_direction(raw_entity: str) -> str:
    text = str(raw_entity or "").strip()
    return "entity_to_partner" if text.upper().startswith("E") else "partner_to_entity"


def build_pair_candidates(base_rows, fact_ids, fact_registry, label_maps):
    base_rows_by_group = defaultdict(list)
    base_rows_by_pair = defaultdict(list)
    all_pairs = set()
    pair_fact_families = defaultdict(set)
    pair_first_seen = {}

    for row in base_rows:
        base_rows_by_group[row.pair_group].append(row)
        base_rows_by_pair[row.pair_key].append(row)
        all_pairs.add(row.pair_key)

    for order, fact_id in enumerate(fact_ids):
        fact = fact_registry[fact_id]
        pair_key = (fact.entity_num, fact.partner_num)
        all_pairs.add(pair_key)
        pair_fact_families[_group_key(*pair_key)].add(fact.family)
        pair_first_seen.setdefault(pair_key, order)

    return {
        "base_rows_by_group": base_rows_by_group,
        "base_rows_by_pair": base_rows_by_pair,
        "all_pairs": all_pairs,
        "pair_labels": label_maps["pair_labels"],
        "entity_labels": label_maps["entity_labels"],
        "partner_labels": label_maps["partner_labels"],
        "journal_pair_votes": label_maps["journal_pair_votes"],
        "pair_fact_families": pair_fact_families,
        "pair_first_seen": pair_first_seen,
    }


def choose_pair_owner(pair_group, candidates, alias_map=None, overrides=None):
    overrides = overrides or PAIR_OWNER_OVERRIDES
    base_rows_by_group = candidates["base_rows_by_group"]
    base_rows_by_pair = candidates["base_rows_by_pair"]
    pair_labels = candidates["pair_labels"]
    entity_labels = candidates["entity_labels"]
    partner_labels = candidates["partner_labels"]
    journal_pair_votes = candidates["journal_pair_votes"]

    a, b = pair_group
    orientations = [(a, b)] if a == b else [(a, b), (b, a)]

    trusted_rows = []
    for orientation in orientations:
        rows = base_rows_by_pair.get(orientation, [])
        forward_rows = [row for row in rows if row.row_direction == "entity_to_partner"]
        if forward_rows:
            chosen = min(forward_rows, key=lambda row: (row.row_num is None, row.row_num or 0))
            trusted_rows.append((orientation, chosen))
    if trusted_rows:
        orientation, row = min(trusted_rows, key=lambda item: (item[1].row_num is None, item[1].row_num or 0))
        return PairOwner(
            pair_group=pair_group,
            canonical_entity=orientation[0],
            canonical_partner=orientation[1],
            owner_type="icm",
            owner_reason="icm_trusted_direction",
            icm_row_num=row.row_num,
            display_entity=row.entity_label,
            display_partner=row.partner_label,
        )

    if pair_group in overrides:
        orientation = overrides[pair_group]
        pair_label = pair_labels.get(orientation, {})
        return PairOwner(
            pair_group=pair_group,
            canonical_entity=orientation[0],
            canonical_partner=orientation[1],
            owner_type="synthetic",
            owner_reason="explicit_override",
            icm_row_num=None,
            display_entity=pair_label.get("entity") or entity_labels.get(orientation[0], orientation[0]),
            display_partner=pair_label.get("partner") or partner_labels.get(orientation[1], f"ICP_{orientation[1]}"),
        )

    forward_pair_labels = []
    for orientation in orientations:
        pair_label = pair_labels.get(orientation)
        if pair_label and _classify_label_direction(pair_label.get("entity")) == "entity_to_partner":
            forward_pair_labels.append(orientation)
    if forward_pair_labels:
        orientation = min(forward_pair_labels)
        pair_label = pair_labels.get(orientation, {})
        return PairOwner(
            pair_group=pair_group,
            canonical_entity=orientation[0],
            canonical_partner=orientation[1],
            owner_type=pair_label.get("family", "journal"),
            owner_reason="journal_label_evidence",
            icm_row_num=None,
            display_entity=pair_label.get("entity") or entity_labels.get(orientation[0], orientation[0]),
            display_partner=pair_label.get("partner") or partner_labels.get(orientation[1], f"ICP_{orientation[1]}"),
        )

    forward_entity_labels = []
    for orientation in orientations:
        entity_label = entity_labels.get(orientation[0], "")
        if entity_label and _classify_label_direction(entity_label) == "entity_to_partner":
            forward_entity_labels.append(orientation)
    if forward_entity_labels:
        orientation = min(forward_entity_labels)
        pair_label = pair_labels.get(orientation, {})
        return PairOwner(
            pair_group=pair_group,
            canonical_entity=orientation[0],
            canonical_partner=orientation[1],
            owner_type="journal",
            owner_reason="entity_label_evidence",
            icm_row_num=None,
            display_entity=pair_label.get("entity") or entity_labels.get(orientation[0], orientation[0]),
            display_partner=pair_label.get("partner") or partner_labels.get(orientation[1], f"ICP_{orientation[1]}"),
        )

    vote_scores = [(journal_pair_votes.get(orientation, 0), orientation) for orientation in orientations]
    top_votes = max(vote_scores, key=lambda item: item[0])[0] if vote_scores else 0
    if top_votes > 0:
        orientation = min(
            [orientation for score, orientation in vote_scores if score == top_votes],
            key=lambda item: item,
        )
        pair_label = pair_labels.get(orientation, {})
        return PairOwner(
            pair_group=pair_group,
            canonical_entity=orientation[0],
            canonical_partner=orientation[1],
            owner_type="journal",
            owner_reason="journal_vote_evidence",
            icm_row_num=None,
            display_entity=pair_label.get("entity") or entity_labels.get(orientation[0], orientation[0]),
            display_partner=pair_label.get("partner") or partner_labels.get(orientation[1], f"ICP_{orientation[1]}"),
        )

    icm_rows = []
    for orientation in orientations:
        rows = base_rows_by_pair.get(orientation, [])
        if rows:
            icm_rows.append((orientation, min(rows, key=lambda row: (row.row_num is None, row.row_num or 0))))
    if icm_rows:
        orientation, row = min(icm_rows, key=lambda item: (item[1].row_num is None, item[1].row_num or 0))
        return PairOwner(
            pair_group=pair_group,
            canonical_entity=orientation[0],
            canonical_partner=orientation[1],
            owner_type="icm",
            owner_reason="icm_existing_pair",
            icm_row_num=row.row_num,
            display_entity=row.entity_label,
            display_partner=row.partner_label,
        )

    orientation = min(orientations)
    return PairOwner(
        pair_group=pair_group,
        canonical_entity=orientation[0],
        canonical_partner=orientation[1],
        owner_type="synthetic",
        owner_reason="stable_fallback",
        icm_row_num=None,
        display_entity=entity_labels.get(orientation[0], orientation[0]),
        display_partner=partner_labels.get(orientation[1], f"ICP_{orientation[1]}"),
    )


def build_pair_registry(base_rows, fact_ids, fact_registry, label_maps, diagnostics, overrides=None):
    candidates = build_pair_candidates(base_rows, fact_ids, fact_registry, label_maps)
    pair_groups = {_group_key(*pair) for pair in candidates["all_pairs"] if pair[0] and pair[1]}
    pair_labels = candidates["pair_labels"]
    entity_labels = candidates["entity_labels"]
    partner_labels = candidates["partner_labels"]
    pair_first_seen = candidates["pair_first_seen"]

    pair_registry: dict[tuple[str, str], PairOwner] = {}
    for pair_group in sorted(pair_groups):
        owner = choose_pair_owner(pair_group, candidates, overrides=overrides)
        pair_registry[pair_group] = owner
        record_diagnostic(
            diagnostics,
            "pairs",
            owner.owner_reason,
            source_row=owner.icm_row_num,
            raw_entity=owner.display_entity,
            raw_partner=owner.display_partner,
            normalized_entity=owner.canonical_entity,
            normalized_partner=owner.canonical_partner,
            extra={"owner_type": owner.owner_type, "pair_group": pair_group},
        )

    real_group_order = []
    seen_groups = set()
    for row in sorted(base_rows, key=lambda item: (item.row_num is None, item.row_num or 0)):
        if row.pair_group in seen_groups:
            continue
        seen_groups.add(row.pair_group)
        real_group_order.append(row.pair_group)

    synthetic_groups = sorted(pair_group for pair_group in pair_registry if pair_group not in seen_groups)
    row_registry = []
    for pair_group in real_group_order + synthetic_groups:
        base_rows_for_group = candidates["base_rows_by_group"].get(pair_group, [])
        orientations = [pair for pair in candidates["all_pairs"] if _group_key(*pair) == pair_group]
        non_e_bidirectional = (
            not base_rows_for_group
            and len(set(orientations)) > 1
            and any(family in {"parent", "contrib"} for family in candidates["pair_fact_families"].get(pair_group, set()))
        )
        if non_e_bidirectional:
            orientation_candidates = []
            for orientation in sorted(set(orientations), key=lambda pair: pair_first_seen.get(pair, 10**9)):
                entity_label = pair_labels.get(orientation, {}).get("entity") or entity_labels.get(orientation[0], orientation[0])
                partner_label = pair_labels.get(orientation, {}).get("partner") or partner_labels.get(orientation[1], f"ICP_{orientation[1]}")
                if _classify_label_direction(entity_label) != "partner_to_entity":
                    non_e_bidirectional = False
                    break
                orientation_candidates.append((orientation, entity_label, partner_label))
            if non_e_bidirectional and len(orientation_candidates) > 1:
                for orientation, entity_label, partner_label in orientation_candidates:
                    row_registry.append(
                        OutputRow(
                            row_key=orientation,
                            pair_group=pair_group,
                            display_entity=entity_label,
                            display_partner=partner_label,
                            source_row_num=None,
                            is_synthetic=True,
                            owner_reason="non_e_bidirectional_split",
                            owner_type="journal",
                        )
                    )
                continue

        owner = pair_registry[pair_group]
        row_registry.append(
            OutputRow(
                row_key=(owner.canonical_entity, owner.canonical_partner),
                pair_group=pair_group,
                display_entity=owner.display_entity,
                display_partner=owner.display_partner,
                source_row_num=owner.icm_row_num,
                is_synthetic=owner.icm_row_num is None,
                owner_reason=owner.owner_reason,
                owner_type=owner.owner_type,
            )
        )

    return pair_registry, row_registry
