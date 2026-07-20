import copy
import json
import math
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


DEFAULT_RULES = {
    "schema_version": 1,
    "analysis": {
        "near_value_ratio": {"R": 2.2, "C": 2.2},
        "voltage_ratio_limit": 4.0,
        "minimum_complete_spec_fields": 4,
    },
    "risk": {
        "lifecycle_keywords": [
            "EOL",
            "OBSOLETE",
            "DISCONTINUED",
            "NRND",
            "NOT RECOMMENDED",
            "LAST TIME BUY",
        ],
        "single_source_vendor_limit": 1,
        "high_fragmentation_pn_count": 3,
        "unknown_vendor_is_risk": True,
        "missing_part_number_is_risk": True,
    },
    "avl": {
        "minimum_vendors": 2,
        "approved_vendors": [],
        "blocked_vendors": [],
    },
    "cost_down": {
        "minimum_impact_quantity": 1,
        "same_spec_priority": 90,
        "near_value_priority": 60,
        "package_priority": 45,
    },
    "confidence": {
        "exact_specification": 0.98,
        "structured_attribute": 0.92,
        "near_value": 0.82,
        "keyword_risk": 0.88,
        "missing_data": 0.99,
    },
    "custom_rules": [],
}


class RuleConfigurationError(ValueError):
    """Raised when a rule library contains invalid configuration."""


def _deep_merge(base, override):
    merged = copy.deepcopy(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _deep_merge(merged[key], value)
        else:
            merged[key] = copy.deepcopy(value)
    return merged


class RuleLibrary:
    """Versioned rule configuration with JSON override support."""

    def __init__(self, rules=None, source="built-in defaults"):
        self.rules = _deep_merge(DEFAULT_RULES, rules or {})
        self.source = source
        self._validate()

    @classmethod
    def load(cls, path=None):
        if path is None:
            return cls()

        rule_path = Path(path).expanduser().resolve()
        if not rule_path.is_file():
            raise FileNotFoundError(f"Rule library not found: {rule_path}")

        try:
            with rule_path.open("r", encoding="utf-8") as rule_file:
                overrides = json.load(rule_file)
        except json.JSONDecodeError as error:
            raise RuleConfigurationError(
                f"Invalid JSON rule library at line {error.lineno}: {error.msg}"
            ) from error

        if not isinstance(overrides, dict):
            raise RuleConfigurationError("The rule library root must be a JSON object.")

        return cls(overrides, str(rule_path))

    def get(self, *keys, default=None):
        value = self.rules
        for key in keys:
            if not isinstance(value, dict) or key not in value:
                return default
            value = value[key]
        return value

    def near_value_ratio(self, component_type):
        return float(self.get("analysis", "near_value_ratio", component_type, default=1.25))

    def to_dict(self):
        return copy.deepcopy(self.rules)

    def to_dataframe(self):
        rows = []

        def flatten(value, prefix=""):
            if isinstance(value, dict):
                for key, nested_value in value.items():
                    flatten(nested_value, f"{prefix}.{key}" if prefix else key)
                return

            rows.append(
                {
                    "Rule_Path": prefix,
                    "Value": json.dumps(value, ensure_ascii=False)
                    if isinstance(value, list)
                    else value,
                    "Source": self.source,
                }
            )

        flatten(self.rules)
        return pd.DataFrame(rows, columns=["Rule_Path", "Value", "Source"])

    def _validate(self):
        if self.get("schema_version") != 1:
            raise RuleConfigurationError("Only rule schema_version 1 is supported.")

        for component_type in ("R", "C"):
            ratio = self.get("analysis", "near_value_ratio", component_type)
            if not self._is_finite_number(ratio) or ratio <= 1:
                raise RuleConfigurationError(
                    f"analysis.near_value_ratio.{component_type} must be a finite number greater than 1."
                )

        voltage_ratio_limit = self.get("analysis", "voltage_ratio_limit")
        if not self._is_finite_number(voltage_ratio_limit) or voltage_ratio_limit <= 1:
            raise RuleConfigurationError(
                "analysis.voltage_ratio_limit must be a finite number greater than 1."
            )

        minimum_fields = self.get("analysis", "minimum_complete_spec_fields")
        if (
            isinstance(minimum_fields, bool)
            or not isinstance(minimum_fields, int)
            or not 1 <= minimum_fields <= 5
        ):
            raise RuleConfigurationError(
                "analysis.minimum_complete_spec_fields must be an integer from 1 to 5."
            )

        vendor_limit = self.get("risk", "single_source_vendor_limit")
        if isinstance(vendor_limit, bool) or not isinstance(vendor_limit, int) or vendor_limit < 1:
            raise RuleConfigurationError(
                "risk.single_source_vendor_limit must be an integer of at least 1."
            )

        fragmentation_limit = self.get("risk", "high_fragmentation_pn_count")
        if (
            isinstance(fragmentation_limit, bool)
            or not isinstance(fragmentation_limit, int)
            or fragmentation_limit < 2
        ):
            raise RuleConfigurationError(
                "risk.high_fragmentation_pn_count must be an integer of at least 2."
            )

        minimum_vendors = self.get("avl", "minimum_vendors")
        if (
            isinstance(minimum_vendors, bool)
            or not isinstance(minimum_vendors, int)
            or minimum_vendors < 1
        ):
            raise RuleConfigurationError(
                "avl.minimum_vendors must be an integer of at least 1."
            )

        minimum_quantity = self.get("cost_down", "minimum_impact_quantity")
        if not self._is_finite_number(minimum_quantity) or minimum_quantity < 0:
            raise RuleConfigurationError(
                "cost_down.minimum_impact_quantity must be a finite nonnegative number."
            )

        for priority_name in (
            "same_spec_priority",
            "near_value_priority",
            "package_priority",
        ):
            priority = self.get("cost_down", priority_name)
            if not self._is_finite_number(priority) or not 0 <= priority <= 100:
                raise RuleConfigurationError(
                    f"cost_down.{priority_name} must be a finite number from 0 to 100."
                )

        for rule_path in (
            ("risk", "lifecycle_keywords"),
            ("avl", "approved_vendors"),
            ("avl", "blocked_vendors"),
        ):
            value = self.get(*rule_path)
            if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
                raise RuleConfigurationError(f"{'.'.join(rule_path)} must be a list of strings.")

        confidence_rules = self.get("confidence", default={})
        if not isinstance(confidence_rules, dict):
            raise RuleConfigurationError("confidence must be a JSON object.")
        for name, confidence in confidence_rules.items():
            if not self._is_finite_number(confidence) or not 0 <= confidence <= 1:
                raise RuleConfigurationError(
                    f"confidence.{name} must be a number between 0 and 1."
                )

        approved_vendors = {
            vendor.casefold()
            for vendor in self.get("avl", "approved_vendors", default=[])
        }
        blocked_vendors = {
            vendor.casefold()
            for vendor in self.get("avl", "blocked_vendors", default=[])
        }
        policy_conflicts = approved_vendors & blocked_vendors
        if policy_conflicts:
            conflicts = ", ".join(sorted(policy_conflicts))
            raise RuleConfigurationError(
                f"Vendors cannot be both approved and blocked: {conflicts}"
            )

        custom_rules = self.get("custom_rules", default=[])
        if not isinstance(custom_rules, list):
            raise RuleConfigurationError("custom_rules must be a list.")

        identifiers = set()
        supported_operators = {"contains_any", "equals_any", "not_in", "missing"}
        supported_severities = {"Critical", "High", "Medium", "Low"}
        for rule in custom_rules:
            if not isinstance(rule, dict):
                raise RuleConfigurationError("Each custom rule must be a JSON object.")
            identifier = str(rule.get("id", "")).strip()
            if not identifier or identifier in identifiers:
                raise RuleConfigurationError("Custom rule IDs must be present and unique.")
            identifiers.add(identifier)
            if rule.get("operator") not in supported_operators:
                raise RuleConfigurationError(
                    f"Custom rule {identifier} uses an unsupported operator."
                )
            if rule.get("severity", "Medium") not in supported_severities:
                raise RuleConfigurationError(
                    f"Custom rule {identifier} uses an unsupported severity."
                )
            if not str(rule.get("field", "")).strip():
                raise RuleConfigurationError(f"Custom rule {identifier} requires a field.")
            values = rule.get("values", [])
            if not isinstance(values, list):
                raise RuleConfigurationError(
                    f"Custom rule {identifier} values must be a list."
                )
            confidence = rule.get("confidence", 0.9)
            if not self._is_finite_number(confidence) or not 0 <= confidence <= 1:
                raise RuleConfigurationError(
                    f"Custom rule {identifier} confidence must be between 0 and 1."
                )

    @staticmethod
    def _is_finite_number(value):
        return (
            not isinstance(value, bool)
            and isinstance(value, (int, float))
            and math.isfinite(float(value))
        )


@dataclass(frozen=True)
class RuleFinding:
    rule_id: str
    severity: str
    confidence: float
    part_number: str
    reference: str
    component_type: str
    finding: str
    evidence: str
    recommendation: str


class ExplainableRuleEngine:
    """Apply deterministic, auditable AI-assisted review rules to normalized rows."""

    COLUMNS = [
        "Finding_ID",
        "Rule_ID",
        "Severity",
        "Confidence",
        "Part_Number",
        "Reference",
        "Component_Type",
        "Finding",
        "Evidence",
        "Recommendation",
    ]

    def __init__(self, library=None):
        self.library = library or RuleLibrary()

    def evaluate(self, normalized_bom):
        findings = []
        for _, row in normalized_bom.iterrows():
            findings.extend(self._evaluate_row(row))
        findings.extend(self._evaluate_specification_groups(normalized_bom))

        records = []
        for index, finding in enumerate(findings, start=1):
            record = finding.__dict__.copy()
            record["Finding_ID"] = f"RULE-{index:04d}"
            records.append(
                {
                    "Finding_ID": record["Finding_ID"],
                    "Rule_ID": record["rule_id"],
                    "Severity": record["severity"],
                    "Confidence": record["confidence"],
                    "Part_Number": record["part_number"],
                    "Reference": record["reference"],
                    "Component_Type": record["component_type"],
                    "Finding": record["finding"],
                    "Evidence": record["evidence"],
                    "Recommendation": record["recommendation"],
                }
            )

        severity_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
        records.sort(key=lambda item: (severity_order.get(item["Severity"], 9), item["Rule_ID"]))
        return pd.DataFrame(records, columns=self.COLUMNS)

    def _evaluate_row(self, row):
        findings = []
        component_type = str(row.get("Component_Type", ""))
        part_number = str(row.get("Part_Number", "") or "")
        reference = str(row.get("Reference", "") or "")

        if (
            self.library.get("risk", "missing_part_number_is_risk")
            and not part_number.strip()
        ):
            findings.append(
                self._finding(
                    "DATA-001",
                    "High",
                    "Missing manufacturer part number",
                    f"Reference {reference or '(blank)'} has no part number",
                    "Complete the MPN before AVL, sourcing, or consolidation review.",
                    row,
                    "missing_data",
                )
            )

        if component_type in {"R", "C"} and not str(row.get("Normalized_Value", "")).strip():
            findings.append(
                self._finding(
                    "DATA-002",
                    "High",
                    "Electrical value could not be normalized",
                    self._source_evidence(row),
                    "Correct the description or add a parsing rule before comparison.",
                    row,
                    "missing_data",
                )
            )

        data_quality_score = float(row.get("Data_Quality_Score", 100) or 0)
        if component_type in {"R", "C"} and data_quality_score < 75:
            findings.append(
                self._finding(
                    "DATA-003",
                    "Medium",
                    "Engineering specification is incomplete",
                    f"Data quality score: {data_quality_score:g}/100",
                    "Complete package, tolerance, voltage, and material attributes as applicable.",
                    row,
                    "missing_data",
                )
            )

        if not bool(row.get("Quantity_Valid", True)):
            findings.append(
                self._finding(
                    "DATA-004",
                    "Medium",
                    "Source quantity is invalid",
                    str(row.get("Quantity_Issue", "Invalid quantity")),
                    "Correct the source quantity before demand, AVL, or savings decisions.",
                    row,
                    "missing_data",
                )
            )

        vendor = str(row.get("Vendor", "") or "")
        if self.library.get("risk", "unknown_vendor_is_risk") and vendor.casefold() in {
            "",
            "unknown",
        }:
            findings.append(
                self._finding(
                    "SUPPLY-001",
                    "Medium",
                    "Manufacturer is unknown",
                    f"Part number: {part_number or '(blank)'}",
                    "Identify the manufacturer before AVL or lifecycle assessment.",
                    row,
                    "missing_data",
                )
            )

        source_text = self._source_evidence(row).upper()
        matched_keywords = [
            keyword
            for keyword in self.library.get("risk", "lifecycle_keywords", default=[])
            if keyword.upper() in source_text
        ]
        if matched_keywords:
            findings.append(
                self._finding(
                    "LIFE-001",
                    "Critical",
                    "Lifecycle risk keyword detected",
                    ", ".join(matched_keywords),
                    "Validate lifecycle status and create a qualified replacement plan.",
                    row,
                    "keyword_risk",
                )
            )

        blocked_vendors = {
            vendor_name.casefold()
            for vendor_name in self.library.get("avl", "blocked_vendors", default=[])
        }
        if vendor.casefold() in blocked_vendors:
            findings.append(
                self._finding(
                    "AVL-002",
                    "Critical",
                    "Manufacturer is blocked by the configured AVL policy",
                    vendor,
                    "Use an approved source or obtain a documented exception.",
                    row,
                    "structured_attribute",
                )
            )

        findings.extend(self._evaluate_custom_rules(row))
        return findings

    def _evaluate_custom_rules(self, row):
        findings = []
        for rule in self.library.get("custom_rules", default=[]):
            if not rule.get("enabled", True):
                continue

            component_types = rule.get("component_types", [])
            component_type = str(row.get("Component_Type", "") or "")
            if component_types and component_type not in component_types:
                continue

            field = rule["field"]
            source_value = row.get(field, "")
            value = "" if pd.isna(source_value) else str(source_value).strip()
            configured_values = [str(item).strip() for item in rule.get("values", [])]
            operator = rule["operator"]
            matched_values = []

            if operator == "missing":
                matched = not value
            elif operator == "contains_any":
                matched_values = [
                    item for item in configured_values if item.casefold() in value.casefold()
                ]
                matched = bool(matched_values)
            elif operator == "equals_any":
                matched_values = [
                    item for item in configured_values if item.casefold() == value.casefold()
                ]
                matched = bool(matched_values)
            else:
                matched = bool(value) and value.casefold() not in {
                    item.casefold() for item in configured_values
                }

            if not matched:
                continue

            evidence = f"{field}={value or '(blank)'}"
            if matched_values:
                evidence += f"; matched={', '.join(matched_values)}"
            confidence = float(rule.get("confidence", 0.9))
            findings.append(
                RuleFinding(
                    rule_id=rule["id"],
                    severity=rule.get("severity", "Medium"),
                    confidence=max(0.0, min(1.0, confidence)),
                    part_number=str(row.get("Part_Number", "") or ""),
                    reference=str(row.get("Reference", "") or ""),
                    component_type=component_type,
                    finding=rule.get("finding", f"Custom rule {rule['id']} matched"),
                    evidence=evidence,
                    recommendation=rule.get(
                        "recommendation",
                        "Review this component against the configured policy.",
                    ),
                )
            )

        return findings

    def _evaluate_specification_groups(self, normalized_bom):
        findings = []
        candidates = normalized_bom[
            normalized_bom["Normalize_Key"].fillna("").ne("")
            & normalized_bom["Component_Type"].isin(["R", "C"])
        ]
        vendor_limit = self.library.get("risk", "single_source_vendor_limit")
        fragmentation_limit = self.library.get("risk", "high_fragmentation_pn_count")

        for normalize_key, group in candidates.groupby("Normalize_Key", sort=True):
            vendors = self._unique_values(group["Vendor"], excluded={"unknown"})
            part_numbers = self._unique_values(group["Part_Number"])
            representative = group.iloc[0]

            if len(vendors) <= vendor_limit:
                findings.append(
                    self._finding(
                        "SUPPLY-002",
                        "High",
                        "Single-source specification risk",
                        f"{normalize_key}; vendors: {', '.join(vendors) or 'unknown'}",
                        "Qualify an alternate source and maintain it in the AVL.",
                        representative,
                        "exact_specification",
                    )
                )

            if len(part_numbers) >= fragmentation_limit:
                findings.append(
                    self._finding(
                        "COST-001",
                        "Medium",
                        "High part-number fragmentation for one specification",
                        f"{len(part_numbers)} MPNs: {', '.join(part_numbers)}",
                        "Review demand aggregation, preferred MPN selection, and cost-down potential.",
                        representative,
                        "exact_specification",
                    )
                )

        return findings

    def _finding(
        self,
        rule_id,
        severity,
        finding,
        evidence,
        recommendation,
        row,
        confidence_key,
    ):
        return RuleFinding(
            rule_id=rule_id,
            severity=severity,
            confidence=float(self.library.get("confidence", confidence_key, default=0.8)),
            part_number=str(row.get("Part_Number", "") or ""),
            reference=str(row.get("Reference", "") or ""),
            component_type=str(row.get("Component_Type", "") or ""),
            finding=finding,
            evidence=evidence,
            recommendation=recommendation,
        )

    @staticmethod
    def _source_evidence(row):
        preferred_columns = (
            "Component_Name",
            "Component Name",
            "Description",
            "Part Description",
            "Lifecycle",
            "Lifecycle_Status",
            "Status",
            "料件狀態",
        )
        values = []
        for column in preferred_columns:
            value = row.get(column, "")
            if pd.notna(value) and str(value).strip():
                values.append(str(value).strip())
        return " | ".join(values) or "No descriptive source data"

    @staticmethod
    def _unique_values(series, excluded=None):
        excluded = {value.casefold() for value in (excluded or set())}
        values = {}
        for value in series:
            if pd.isna(value):
                continue
            text = str(value).strip()
            if text and text.casefold() not in excluded:
                values.setdefault(text.casefold(), text)
        return list(values.values())