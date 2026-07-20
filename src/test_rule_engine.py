import unittest

import pandas as pd

from rule_engine import ExplainableRuleEngine, RuleConfigurationError, RuleLibrary


class RuleEngineTests(unittest.TestCase):
    def test_custom_rule_produces_explainable_finding(self):
        rules = RuleLibrary().to_dict()
        rules["custom_rules"] = [
            {
                "id": "ORG-001",
                "field": "Lifecycle_Status",
                "operator": "equals_any",
                "values": ["Restricted"],
                "severity": "High",
                "confidence": 0.95,
                "finding": "Restricted lifecycle state",
                "recommendation": "Obtain sourcing approval.",
            }
        ]
        dataframe = pd.DataFrame(
            [
                {
                    "Component_Type": "R",
                    "Part_Number": "PR1",
                    "Reference": "R1",
                    "Normalized_Value": "10kOhm",
                    "Normalize_Key": "R_10kOhm_0402",
                    "Vendor": "Yageo",
                    "Lifecycle_Status": "Restricted",
                    "Data_Quality_Score": 100,
                }
            ]
        )

        findings = ExplainableRuleEngine(RuleLibrary(rules)).evaluate(dataframe)
        custom_finding = findings[findings["Rule_ID"] == "ORG-001"].iloc[0]

        self.assertEqual(custom_finding["Severity"], "High")
        self.assertEqual(custom_finding["Confidence"], 0.95)
        self.assertIn("Lifecycle_Status=Restricted", custom_finding["Evidence"])

    def test_rejects_conflicting_avl_policy(self):
        rules = RuleLibrary().to_dict()
        rules["avl"]["approved_vendors"] = ["Vendor A"]
        rules["avl"]["blocked_vendors"] = ["vendor a"]

        with self.assertRaises(RuleConfigurationError):
            RuleLibrary(rules)

    def test_rejects_out_of_range_confidence(self):
        rules = RuleLibrary().to_dict()
        rules["confidence"]["near_value"] = 1.1

        with self.assertRaises(RuleConfigurationError):
            RuleLibrary(rules)

    def test_rejects_nonfinite_and_malformed_numeric_policies(self):
        cases = (
            (("analysis", "near_value_ratio", "R"), float("nan")),
            (("analysis", "minimum_complete_spec_fields"), True),
            (("risk", "high_fragmentation_pn_count"), "three"),
            (("avl", "minimum_vendors"), 0),
            (("cost_down", "minimum_impact_quantity"), float("inf")),
            (("cost_down", "same_spec_priority"), 101),
        )

        for path, invalid_value in cases:
            with self.subTest(path=path):
                rules = RuleLibrary().to_dict()
                target = rules
                for key in path[:-1]:
                    target = target[key]
                target[path[-1]] = invalid_value
                with self.assertRaises(RuleConfigurationError):
                    RuleLibrary(rules)

    def test_vendor_case_differences_remain_single_source(self):
        dataframe = pd.DataFrame(
            [
                {
                    "Component_Type": "C",
                    "Part_Number": "P1",
                    "Reference": "C1",
                    "Normalized_Value": "100nF",
                    "Normalize_Key": "C_100nF_16V_X7R_0402",
                    "Vendor": "MURATA",
                    "Data_Quality_Score": 100,
                },
                {
                    "Component_Type": "C",
                    "Part_Number": "P2",
                    "Reference": "C2",
                    "Normalized_Value": "100nF",
                    "Normalize_Key": "C_100nF_16V_X7R_0402",
                    "Vendor": "Murata",
                    "Data_Quality_Score": 100,
                },
            ]
        )

        findings = ExplainableRuleEngine().evaluate(dataframe)

        self.assertIn("SUPPLY-002", set(findings["Rule_ID"]))


if __name__ == "__main__":
    unittest.main()