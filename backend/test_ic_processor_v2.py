import os
import sys
import unittest
import uuid

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "venv", "Lib", "site-packages"))
sys.path.insert(0, os.path.dirname(__file__))

import openpyxl

from app.ic_processor import (
    compare_v1_v2_report31,
    process_icm_report,
)
from app.ic_refactor.facts import normalize_account_code, normalize_party_code
from app.ic_refactor.ledger import build_cell_ledger
from app.ic_refactor.models import LayoutColumn, LayoutSpec, NormalizedFact, PairOwner
from app.ic_refactor.pairs import build_pair_registry
from app.ic_refactor.pipeline import run_pipeline_v2
from app.ic_refactor.plug import build_plug_section_facts, derive_plug_facts


class IcProcessorV2Tests(unittest.TestCase):
    def test_normalize_party_code_rejects_suffix_code(self):
        diagnostics = {"bad_codes": []}
        self.assertEqual(normalize_party_code("E117100"), "117100")
        self.assertEqual(normalize_party_code("ICP_007009"), "007009")
        self.assertIsNone(normalize_party_code("117000A", diagnostics=diagnostics, field_name="entity"))
        self.assertEqual(len(diagnostics["bad_codes"]), 1)

    def test_normalize_account_code_extracts_numeric(self):
        self.assertEqual(normalize_account_code("534018 - 534018:Interest expense - related parties"), "534018")
        self.assertEqual(normalize_account_code("[111000].[111001]:111001:Land for undertermined use"), "111001")
        self.assertIsNone(normalize_account_code("[FCCS_Long Term Assets].[Plug_InvSh]:Investment-share capital plug account"))

    def test_pair_registry_prefers_trusted_icm_direction(self):
        base_rows = [
            type("Row", (), {
                "pair_key": ("007009", "117100"),
                "pair_group": ("007009", "117100"),
                "row_num": 10,
                "row_direction": "partner_to_entity",
                "entity_label": "007009 - Example",
                "partner_label": "ICP_E117100 - Example ICP",
            })(),
            type("Row", (), {
                "pair_key": ("117100", "007009"),
                "pair_group": ("007009", "117100"),
                "row_num": 11,
                "row_direction": "entity_to_partner",
                "entity_label": "E117100 - Example",
                "partner_label": "ICP_007009 - Example ICP",
            })(),
        ]
        pair_registry, row_registry = build_pair_registry(
            base_rows,
            [],
            {},
            {"entity_labels": {}, "partner_labels": {}, "pair_labels": {}, "journal_pair_votes": {}},
            {"pairs": []},
        )
        owner = pair_registry[("007009", "117100")]
        self.assertEqual((owner.canonical_entity, owner.canonical_partner), ("117100", "007009"))
        self.assertEqual(owner.owner_reason, "icm_trusted_direction")
        self.assertEqual(row_registry[0].row_key, ("117100", "007009"))

    def test_derive_plug_facts_and_plug_section(self):
        fact_registry = {
            "fact_000001": NormalizedFact(
                family="parent",
                source_file="parent.xlsx",
                source_row=31,
                entity_num="117100",
                partner_num="007009",
                account_code="534018",
                amount=100.0,
                direction="entity_to_partner",
                raw_entity="E117100:Entity",
                raw_partner="ICP_007009:Partner",
                raw_account="534018",
                meta={},
            ),
            "fact_000002": NormalizedFact(
                family="ic_elim",
                source_file="elim.xlsx",
                source_row=33,
                entity_num="117100",
                partner_num="007009",
                account_code="188600",
                amount=75.0,
                direction="entity_to_partner",
                raw_entity="E117100:Entity",
                raw_partner="ICP_007009:Partner",
                raw_account="Total",
                meta={"column_role": "total"},
            ),
        }
        parent_plug = derive_plug_facts(["fact_000001"], fact_registry, {"534018"}, "188600", "parent")
        standalone = build_plug_section_facts(["fact_000002"], fact_registry, "188600")
        self.assertEqual(len(parent_plug), 1)
        self.assertEqual(parent_plug[0].account_code, "188600")
        self.assertEqual(parent_plug[0].meta["derived_from"], "parent")
        self.assertEqual(len(standalone), 1)
        self.assertEqual(standalone[0].meta["derived_from"], "ic_elim_total")

    def test_cell_ledger_assigns_exactly_once(self):
        fact_registry = {
            "fact_000001": NormalizedFact(
                family="parent",
                source_file="parent.xlsx",
                source_row=31,
                entity_num="117100",
                partner_num="007009",
                account_code="534018",
                amount=25.0,
                direction="entity_to_partner",
                raw_entity="E117100:Entity",
                raw_partner="ICP_007009:Partner",
                raw_account="534018",
                meta={},
            ),
            "fact_000002": NormalizedFact(
                family="parent",
                source_file="parent.xlsx",
                source_row=32,
                entity_num="007009",
                partner_num="117100",
                account_code="534018",
                amount=40.0,
                direction="partner_to_entity",
                raw_entity="007009:Partner",
                raw_partner="ICP_E117100:Entity ICP",
                raw_account="534018",
                meta={},
            ),
        }
        pair_registry = {
            ("007009", "117100"): PairOwner(
                pair_group=("007009", "117100"),
                canonical_entity="117100",
                canonical_partner="007009",
                owner_type="icm",
                owner_reason="icm_trusted_direction",
                icm_row_num=5,
                display_entity="E117100 - Entity",
                display_partner="ICP_007009 - Partner",
            )
        }
        layout = LayoutSpec(
            ent_cols=[LayoutColumn(code="534018", description="534018", series="S1", tag="Entity")],
            par_cols=[LayoutColumn(code="534018", description="534018", series="S1", tag="Partner")],
            plug_code="188600",
        )
        ledger, assignment = build_cell_ledger(
            ["fact_000001", "fact_000002"],
            fact_registry,
            pair_registry,
            layout,
            {"unmatched_facts": []},
        )
        self.assertEqual(ledger[(("117100", "007009"), "parent", "entity_side", "534018")].amount, 25.0)
        self.assertEqual(ledger[(("117100", "007009"), "parent", "partner_side", "534018")].amount, 40.0)
        self.assertEqual(len(assignment), 2)
        self.assertTrue(all(payload["reason"] == "routed" for payload in assignment.values()))

    def test_process_icm_report_v2_runs_small_fixture(self):
        from test_ic_processor_bidirectional import _ic_elim_workbook, _journal_workbook, _report_inputs_workbook

        base_dir = os.path.dirname(__file__)
        suffix = uuid.uuid4().hex
        paths = {
            "icm": os.path.join(base_dir, f"tmp_v2_ic_elim_{suffix}.xlsx"),
            "parent": os.path.join(base_dir, f"tmp_v2_parent_{suffix}.xlsx"),
            "contrib": os.path.join(base_dir, f"tmp_v2_contrib_{suffix}.xlsx"),
            "plug": os.path.join(base_dir, f"tmp_v2_plug_{suffix}.xlsx"),
            "inputs": os.path.join(base_dir, f"tmp_v2_inputs_{suffix}.xlsx"),
            "output": os.path.join(base_dir, f"tmp_v2_output_{suffix}.xlsx"),
        }
        for path in paths.values():
            self.addCleanup(lambda p=path: os.path.exists(p) and os.remove(p))

        _ic_elim_workbook(
            paths["icm"],
            [
                {
                    "entity": "E117100 - Entity",
                    "partner": "ICP_007009 - Partner ICP",
                    "entity_value": 100,
                    "partner_value": -25,
                    "total": 75,
                }
            ],
            plug_code="188600",
        )
        _journal_workbook(
            paths["parent"],
            [
                {
                    "entity": "E117100:Entity",
                    "account": "534018 - 534018:Interest expense - related parties",
                    "icp": "ICP_007009:Partner ICP",
                    "debit": 100,
                    "credit": 0,
                }
            ],
        )
        _journal_workbook(paths["contrib"], [])
        _journal_workbook(paths["plug"], [])
        _report_inputs_workbook(paths["inputs"], plug_code="188600", elim_code="534018")

        process_icm_report(
            paths["icm"],
            {
                "parent_journal": paths["parent"],
                "contribution_journal": paths["contrib"],
                "plugaccount_journal": paths["plug"],
            },
            paths["output"],
            report_inputs_path=paths["inputs"],
        )

        wb = openpyxl.load_workbook(paths["output"], data_only=True)
        self.assertIn("ICM Matched", wb.sheetnames)
        self.assertIn("Diagnostics_Assignment", wb.sheetnames)
        self.assertIn("Diagnostics_Pairs", wb.sheetnames)

    def test_report31_comparison_harness(self):
        outputs = compare_v1_v2_report31()
        self.assertTrue(os.path.exists(outputs["current_output"]))
        self.assertTrue(os.path.exists(outputs["refactor_output"]))
        self.assertTrue(os.path.exists(outputs["comparison_output"]))

    def test_run_pipeline_v2_assigns_or_diagnoses_every_routed_fact(self):
        base_dir = r"g:\FCCS\backend\uploads\reports\31\inputs"
        result = run_pipeline_v2(
            os.path.join(base_dir, "IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx"),
            {
                "parent_journal": os.path.join(base_dir, "Parent report.xlsx"),
                "contribution_journal": os.path.join(base_dir, "Contribution report.xlsx"),
                "plugaccount_journal": os.path.join(base_dir, "Journal Report (4).xlsx"),
            },
            report_inputs_path=os.path.join(base_dir, "report Inputs.xlsx"),
        )
        routed_or_diagnosed = len(result.fact_assignment_log)
        self.assertGreater(routed_or_diagnosed, 0)
        self.assertGreaterEqual(
            routed_or_diagnosed,
            len(result.fact_build.facts_parent) + len(result.fact_build.facts_contrib) + len(result.fact_build.facts_plug_source),
        )
        self.assertTrue(
            all(
                payload["reason"] in {"routed", "diagnosed_unmatched", "plug_source_reconciliation_only"}
                for payload in result.fact_assignment_log.values()
            )
        )


if __name__ == "__main__":
    unittest.main()
