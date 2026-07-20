import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from bom_intelligence import BOMIntelligencePlatform
from rule_engine import RuleLibrary


class BOMIntelligencePlatformTests(unittest.TestCase):
    def setUp(self):
        self.bom = pd.DataFrame(
            [
                {
                    "Part Number": "PN1",
                    "Quantity": 2,
                    "Part Reference": "C1",
                    "Package_Type": "SMD",
                    "Component_Name": "MLCC 100NF 16V (0402) X7R",
                    "Description": "MURATA/GRM",
                    "Subsystem_ID": "PCH",
                },
                {
                    "Part Number": "PN2",
                    "Quantity": 1,
                    "Part Reference": "C2",
                    "Package_Type": "SMD",
                    "Component_Name": "MLCC 104 16V (0402) X7R",
                    "Description": "SAMSUNG/CL",
                    "Subsystem_ID": "PCH",
                },
                {
                    "Part Number": "PN3",
                    "Quantity": 1,
                    "Part Reference": "C3",
                    "Package_Type": "SMD",
                    "Component_Name": "MLCC 220NF 16V (0402) X7R",
                    "Description": "YAGEO/CC",
                    "Subsystem_ID": "PCH",
                },
                {
                    "Part Number": "PN4",
                    "Quantity": 1,
                    "Part Reference": "C4",
                    "Package_Type": "SMD",
                    "Component_Name": "MLCC 10UF 6.3V (0402) X5R",
                    "Description": "MURATA/GRM",
                    "Subsystem_ID": "P-CPU",
                },
                {
                    "Part Number": "PR1",
                    "Quantity": 1,
                    "Part Reference": "R1",
                    "Package_Type": "SMD",
                    "Component_Name": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Description": "VISHAY/CRCW",
                    "Subsystem_ID": "PCH",
                },
                {
                    "Part Number": "PR2",
                    "Quantity": 1,
                    "Part Reference": "R2",
                    "Package_Type": "SMD",
                    "Component_Name": "RES 12K 1/16W 0402 1% THICK FILM",
                    "Description": "VISHAY/CRCW",
                    "Subsystem_ID": "PCH",
                },
            ]
        )
        self.platform = BOMIntelligencePlatform()

    def test_analysis_normalizes_and_protects_expected_rows(self):
        reports = self.platform.analyze_dataframe(self.bom)
        normalized_bom = reports["normalized_bom"]

        c2_value = normalized_bom.loc[
            normalized_bom["Reference"] == "C2", "Normalized_Value"
        ].iloc[0]
        self.assertEqual(c2_value, "100nF")
        self.assertEqual(len(reports["duplicate_pn"]), 1)
        self.assertGreaterEqual(len(reports["near_value"]), 2)
        self.assertEqual(len(reports["critical_parts"]), 1)
        self.assertGreaterEqual(len(reports["cost_down"]), 3)
        self.assertIn("Murata", reports["vendor_distribution"]["Vendor"].tolist())

    def test_report_contains_all_review_sheets(self):
        reports = self.platform.analyze_dataframe(self.bom)
        expected_sheets = [
            "Dashboard",
            "Merge Candidate",
            "Specification Summary",
            "Specification Detail",
            "AVL Candidate",
            "Risk Review",
            "Settings",
        ]

        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "report.xlsx"
            self.platform.write_excel_report(reports, output_file)
            workbook = load_workbook(output_file, read_only=False)
            self.assertEqual(workbook.sheetnames, expected_sheets)
            self.assertEqual(workbook["Dashboard"]["A1"].value, "PN Optimization Dashboard")
            self.assertEqual(workbook["Settings"].sheet_state, "hidden")
            self.assertEqual(workbook["Merge Candidate"].freeze_panes, "A2")
            self.assertEqual(workbook["Specification Summary"].freeze_panes, "A2")
            self.assertEqual(workbook["Specification Detail"].freeze_panes, "A2")
            merge_headers = [cell.value for cell in workbook["Merge Candidate"][1]]
            summary_headers = [cell.value for cell in workbook["Specification Summary"][1]]
            detail_headers = [cell.value for cell in workbook["Specification Detail"][1]]
            self.assertEqual(
                merge_headers[:8],
                [
                    "Priority",
                    "Merge Difficulty",
                    "Difference",
                    "Spec",
                    "Current PN",
                    "Current Qty",
                    "Target PN",
                    "Target Qty",
                ],
            )
            self.assertEqual(summary_headers[:9], ["Value", "Spec Detail", "PN Count", "Total Qty", "Target PN", "Priority", "Reason", "Detail", "Group"])
            self.assertEqual(detail_headers[:9], ["Group", "Row Type", "Merge Tree", "PN", "Qty", "Qty Share", "Difference", "Can Merge", "Reason"])
            self.assertTrue(workbook["Specification Summary"].column_dimensions["I"].hidden)
            self.assertIsNotNone(workbook["Specification Summary"]["H2"].hyperlink)
            workbook.close()

    def test_merge_candidates_rank_low_quantity_pns_to_target_pn(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "R-TARGET",
                    "Part Reference": "R102",
                    "Component_Name": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Vendor": "Yageo",
                    "Quantity": 128,
                },
                {
                    "Part Number": "R-LOW-1",
                    "Part Reference": "R153",
                    "Component_Name": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Vendor": "Yageo",
                    "Quantity": 2,
                },
                {
                    "Part Number": "R-LOW-2",
                    "Part Reference": "R608",
                    "Component_Name": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Vendor": "Vishay",
                    "Quantity": 5,
                },
                {
                    "Part Number": "R-LATER",
                    "Part Reference": "R777",
                    "Component_Name": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Vendor": "Yageo",
                    "Quantity": 17,
                },
            ]
        )

        merge_candidates = self.platform.analyze_dataframe(bom)["merge_candidates"]

        self.assertEqual(merge_candidates["Current_PN"].tolist(), ["R-LOW-1", "R-LOW-2"])
        self.assertTrue((merge_candidates["Target_PN"] == "R-TARGET").all())
        self.assertEqual(merge_candidates.iloc[0]["Priority"], "Priority 1")
        self.assertEqual(merge_candidates.iloc[0]["Priority_Stars"], "★★★★★")
        self.assertEqual(merge_candidates.iloc[0]["Quantity_Score"], 100)
        self.assertEqual(merge_candidates.iloc[1]["Spec_Similarity"], 95)

    def test_merge_candidates_ignore_second_source_rows(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "R-TARGET",
                    "Part Reference": "R1",
                    "Item Description": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Comp Type": "",
                    "Quantity": 100,
                },
                {
                    "Part Number": "R-SECOND-SOURCE",
                    "Part Reference": "R1-S",
                    "Item Description": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Comp Type": "S",
                    "Quantity": 1,
                },
                {
                    "Part Number": "R-MAIN-LOW",
                    "Part Reference": "R2",
                    "Item Description": "RES 10K 1/16W 0402 1% THICK FILM",
                    "Comp Type": "",
                    "Quantity": 2,
                },
            ]
        )

        reports = self.platform.analyze_dataframe(bom)
        merge_candidates = reports["merge_candidates"]

        self.assertEqual(merge_candidates["Current_PN"].tolist(), ["R-MAIN-LOW"])
        self.assertNotIn("R-SECOND-SOURCE", merge_candidates["Current_PN"].tolist())
        normalized = reports["normalized_bom"].set_index("Part_Number")
        self.assertTrue(normalized.loc["R-SECOND-SOURCE", "Is_Second_Source"])

    def test_project_name_uses_first_item_description_prefix(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "ASM",
                    "Part Reference": "",
                    "Item Description": "P500MV MAIN BD MODULE//MECHANICAL",
                    "Quantity": 1,
                }
            ]
        )

        reports = self.platform.analyze_dataframe(bom)
        metadata = reports["report_metadata"].set_index("Property")["Value"]

        self.assertEqual(metadata["Project Name"], "P500MV")

    def test_finds_variants_avl_cost_and_risk_opportunities(self):
        rows = [
            ("C1", "PN1", "MLCC 100NF 16V 0402 X7R 10%", "Murata", 10, 0.05, "Active"),
            ("C2", "PN2", "MLCC 104 16V 0402 X7R 10%", "Samsung", 20, 0.02, "Active"),
            ("C3", "PN3", "MLCC 100NF 25V 0402 X7R 10%", "Murata", 1, None, "Active"),
            ("C4", "PN4", "MLCC 100NF 16V 0402 X5R 10%", "TDK", 1, None, "Active"),
            ("C5", "PN5", "MLCC 100NF 16V 0603 X7R 10%", "Yageo", 1, None, "Active"),
            ("R1", "PR1", "RES 10K OHM 1/16W 0402 1% THICK FILM", "Yageo", 1, None, "EOL"),
            ("R2", "PR2", "RES 12K OHM 1/16W 0402 1% THICK FILM", "Yageo", 1, None, "Active"),
        ]
        dataframe = pd.DataFrame(
            [
                {
                    "Part Reference": reference,
                    "Part Number": part_number,
                    "Component_Name": description,
                    "Vendor": vendor,
                    "Quantity": quantity,
                    "Unit Price": unit_price,
                    "Lifecycle": lifecycle,
                }
                for reference, part_number, description, vendor, quantity, unit_price, lifecycle in rows
            ]
        )

        reports = self.platform.analyze_dataframe(dataframe)

        self.assertEqual(len(reports["duplicate_pn"]), 1)
        self.assertEqual(len(reports["different_package"]), 1)
        self.assertEqual(len(reports["different_voltage"]), 1)
        self.assertEqual(len(reports["different_material"]), 1)
        self.assertEqual(len(reports["near_resistance"]), 1)
        self.assertIn("Ready for unified AVL", set(reports["avl_candidates"]["AVL_Readiness"]))
        self.assertIn("LIFE-001", set(reports["risk_components"]["Rule_ID"]))
        self.assertAlmostEqual(
            reports["duplicate_pn"].iloc[0]["Estimated_BOM_Savings"],
            0.3,
        )

    def test_report_escapes_formula_like_source_text(self):
        bom = self.bom.copy()
        bom.loc[0, "Part Number"] = "\t=1+1"
        bom["=HYPERLINK(\"bad\")"] = "source value"
        reports = self.platform.analyze_dataframe(bom)

        with tempfile.TemporaryDirectory() as temporary_directory:
            output_file = Path(temporary_directory) / "safe_report.xlsx"
            self.platform.write_excel_report(reports, output_file)
            workbook = load_workbook(output_file, data_only=False)
            formula_like_cells = [
                cell
                for worksheet in workbook.worksheets
                for row in worksheet.iter_rows()
                for cell in row
                if isinstance(cell.value, str) and "=1+1" in cell.value
            ]
            self.assertTrue(formula_like_cells)
            self.assertTrue(
                all(
                    str(cell.value).startswith("'")
                    for cell in formula_like_cells
                    if str(cell.value).lstrip("\t\r\n\ufeff").startswith("=1+1")
                )
            )
            workbook.close()

    def test_description_only_signal_context_is_protected(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "PR1",
                    "Part Reference": "R1",
                    "Component_Name": "RES 22R 1/16W 0402 1% THICK FILM",
                    "Description": "USB TX termination",
                }
            ]
        )

        normalized = self.platform.analyze_dataframe(bom)["normalized_bom"].iloc[0]

        self.assertTrue(normalized["Is_Critical"])
        self.assertEqual(normalized["Critical_Reason"], "High-speed or RF signal context")

    def test_avl_preferred_part_is_approved_and_vendor_identity_is_casefolded(self):
        rules = RuleLibrary().to_dict()
        rules["avl"]["approved_vendors"] = ["Murata", "TDK"]
        platform = BOMIntelligencePlatform(rule_library=RuleLibrary(rules))
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "PN-UNAPPROVED",
                    "Part Reference": "C1",
                    "Component_Name": "MLCC 100NF 16V 0402 X7R 10%",
                    "Vendor": "CheapCo",
                    "Unit Price": 0.01,
                },
                {
                    "Part Number": "PN-MURATA",
                    "Part Reference": "C2",
                    "Component_Name": "MLCC 100NF 16V 0402 X7R 10%",
                    "Vendor": "MURATA",
                    "Unit Price": 0.05,
                },
                {
                    "Part Number": "PN-TDK",
                    "Part Reference": "C3",
                    "Component_Name": "MLCC 100NF 16V 0402 X7R 10%",
                    "Vendor": "tdk",
                    "Unit Price": 0.06,
                },
            ]
        )

        avl = platform.analyze_dataframe(bom)["avl_candidates"].iloc[0]

        self.assertEqual(avl["Preferred_PN"], "PN-MURATA")
        self.assertEqual(avl["Vendor_Count"], 3)
        self.assertEqual(avl["AVL_Readiness"], "Ready for unified AVL")

    def test_generic_material_does_not_hide_capacitor_dielectric_difference(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "PN-X5R",
                    "Part Reference": "C1",
                    "Component_Name": "MLCC 100NF 16V 0402 X5R 10%",
                    "Material": "Ceramic",
                    "Vendor": "Murata",
                },
                {
                    "Part Number": "PN-X7R",
                    "Part Reference": "C2",
                    "Component_Name": "MLCC 100NF 16V 0402 X7R 10%",
                    "Material": "Ceramic",
                    "Vendor": "TDK",
                },
            ]
        )

        reports = self.platform.analyze_dataframe(bom)

        self.assertTrue(reports["duplicate_pn"].empty)
        self.assertEqual(len(reports["different_material"]), 1)
        self.assertEqual(
            reports["different_material"].iloc[0]["Attribute_Values"],
            "X5R, X7R",
        )

    def test_incomplete_specifications_are_not_comparison_candidates(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "PR1",
                    "Part Reference": "R1",
                    "Component_Name": "RES 10K 0402",
                    "Vendor": "Yageo",
                },
                {
                    "Part Number": "PR2",
                    "Part Reference": "R2",
                    "Component_Name": "RES 12K 0402",
                    "Vendor": "Yageo",
                },
            ]
        )

        reports = self.platform.analyze_dataframe(bom)

        self.assertFalse(reports["normalized_bom"]["Comparison_Eligible"].any())
        self.assertTrue(reports["duplicate_pn"].empty)
        self.assertTrue(reports["near_resistance"].empty)
        self.assertTrue(reports["avl_candidates"].empty)
        self.assertIn("DATA-003", set(reports["ai_rule_findings"]["Rule_ID"]))

    def test_quantity_and_embedded_voltage_are_normalized_safely(self):
        bom = pd.DataFrame(
            [
                {
                    "Part Number": "PC1",
                    "Part Reference": "C1",
                    "Component_Name": "MLCC 10UF 6V3 0402 X5R 10%",
                    "Quantity": "1,000",
                    "Vendor": "Murata",
                },
                {
                    "Part Number": "PC2",
                    "Part Reference": "C2",
                    "Component_Name": "MLCC 10UF 6V3 0402 X5R 10%",
                    "Quantity": "inf",
                    "Vendor": "TDK",
                },
            ]
        )

        reports = self.platform.analyze_dataframe(bom)
        normalized = reports["normalized_bom"].set_index("Part_Number")

        self.assertEqual(normalized.loc["PC1", "Voltage"], "6.3V")
        self.assertEqual(normalized.loc["PC1", "Quantity_Normalized"], 1000)
        self.assertTrue(normalized.loc["PC1", "Quantity_Valid"])
        self.assertEqual(normalized.loc["PC2", "Quantity_Normalized"], 1)
        self.assertFalse(normalized.loc["PC2", "Quantity_Valid"])
        self.assertIn("DATA-004", set(reports["ai_rule_findings"]["Rule_ID"]))


if __name__ == "__main__":
    unittest.main()