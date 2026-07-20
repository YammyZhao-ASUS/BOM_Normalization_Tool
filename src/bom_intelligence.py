import itertools
import math
import re
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

try:
    from .bom_reader import BOMReader
    from .capacitor_normalizer import CapacitorNormalizer
    from .excel_reporter import ExcelReportWriter
    from .normalizer import ResistorNormalizer
    from .rule_engine import ExplainableRuleEngine, RuleLibrary
    from .specification_analyzer import SpecificationAnalyzer
    from .value_extractor import ValueExtractor
except ImportError:
    from bom_reader import BOMReader
    from capacitor_normalizer import CapacitorNormalizer
    from excel_reporter import ExcelReportWriter
    from normalizer import ResistorNormalizer
    from rule_engine import ExplainableRuleEngine, RuleLibrary
    from specification_analyzer import SpecificationAnalyzer
    from value_extractor import ValueExtractor


class BOMIntelligencePlatform:
    """Analyze a hardware BOM and return review-ready data tables."""

    VERSION = "2.0.0"

    SENSITIVE_SUBSYSTEMS = (
        "CPU",
        "VR",
        "VCORE",
        "USB",
        "TYPE-C",
        "TYPEC",
        "PCIE",
        "PCI-E",
        "DDR",
        "RF",
        "WLAN",
        "WI-FI",
        "HDMI",
        "DISPLAYPORT",
        "DP",
    )

    KNOWN_VENDORS = (
        "MURATA",
        "SAMSUNG",
        "WALSIN",
        "YAGEO",
        "TAIYO YUDEN",
        "KEMET",
        "TDK",
        "AVX",
        "VISHAY",
        "ROHM",
        "PANASONIC",
        "NICHICON",
        "KYOCERA",
        "NIC COMPONENTS",
        "MULTICOMP",
    )

    COLUMN_CANDIDATES = {
        "reference": ("Part Reference", "Reference", "RefDes", "Ref", "插件位置"),
        "part_number": (
            "Part Number",
            "Manufacturer Part Number",
            "Mfr Part Number",
            "MPN",
            "PN",
            "Item Number",
        ),
        "component_name": (
            "Component_Name",
            "Component Name",
            "Component",
            "Item Description",
        ),
        "description": ("Description", "Part Description", "Item Description"),
        "package": ("Package_Type", "Package", "PCB Footprint", "Footprint"),
        "quantity": ("Quantity", "Qty", "QTY", "Comp Quantity"),
        "vendor": ("Vendor", "Manufacturer", "Mfr", "Supplier", "OEM_Component"),
        "critical": ("Critical_Part", "Critical Part", "Critical"),
        "subsystem": ("Subsystem_ID", "Subsystem", "Circuit", "Block"),
        "voltage": ("Voltage", "Rated Voltage", "Voltage Rating"),
        "material": ("Material", "Dielectric", "Technology"),
        "tolerance": ("Tolerance", "Tol"),
        "power": ("Power", "Power Rating", "Rated Power"),
        "unit_price": ("Unit Price", "Unit Cost", "Price", "Cost"),
        "lifecycle": ("Lifecycle", "Lifecycle Status", "Part Status", "料件狀態"),
        "bom_comp_type": ("Comp Type", "BOM Comp Type"),
    }

    def __init__(self, near_value_ratio=None, rule_file=None, rule_library=None):
        if rule_file and rule_library:
            raise ValueError("Use either rule_file or rule_library, not both.")

        library = rule_library or RuleLibrary.load(rule_file)
        if near_value_ratio is not None:
            if isinstance(near_value_ratio, dict):
                ratios = {
                    component_type: float(near_value_ratio.get(component_type, library.near_value_ratio(component_type)))
                    for component_type in ("R", "C")
                }
            else:
                ratios = {"R": float(near_value_ratio), "C": float(near_value_ratio)}

            if any(ratio <= 1 for ratio in ratios.values()):
                raise ValueError("near_value_ratio must be greater than 1")

            rules = library.to_dict()
            rules["analysis"]["near_value_ratio"] = ratios
            library = RuleLibrary(rules, f"{library.source}; runtime near-value override")

        self.rule_library = library
        self.near_value_ratios = {
            component_type: library.near_value_ratio(component_type)
            for component_type in ("R", "C")
        }
        self.near_value_ratio = max(self.near_value_ratios.values())
        self.value_extractor = ValueExtractor()
        self.capacitor_normalizer = CapacitorNormalizer()
        self.resistor_normalizer = ResistorNormalizer()
        self.rule_engine = ExplainableRuleEngine(library)
        self.specification_analyzer = SpecificationAnalyzer(library)

    def analyze_file(self, input_file, sheet_name=None):
        reader = BOMReader(input_file, self.COLUMN_CANDIDATES)
        reports = self.analyze_dataframe(reader.load(sheet_name))
        input_metadata = pd.DataFrame(
            list(reader.metadata().items()),
            columns=["Property", "Value"],
        )
        reports["report_metadata"] = pd.concat(
            [reports["report_metadata"], input_metadata],
            ignore_index=True,
        )
        return reports

    def analyze_dataframe(self, dataframe):
        if dataframe.empty:
            raise ValueError("The BOM contains no rows.")

        columns = self._resolve_columns(dataframe.columns)
        project_name = self._project_name_from_dataframe(dataframe, columns)
        normalized_bom = self._normalize_bom(dataframe, columns)
        duplicate_pn = self._find_duplicate_specs(normalized_bom)
        merge_candidates = self._find_merge_candidates(normalized_bom)
        near_value = self._find_near_values(normalized_bom)
        near_resistance = near_value[near_value["Component_Type"] == "R"].reset_index(drop=True)
        near_capacitance = near_value[near_value["Component_Type"] == "C"].reset_index(drop=True)
        different_package = self.specification_analyzer.find_package_variants(normalized_bom)
        different_voltage = self.specification_analyzer.find_voltage_variants(normalized_bom)
        different_material = self.specification_analyzer.find_material_variants(normalized_bom)
        ai_rule_findings = self.rule_engine.evaluate(normalized_bom)
        risk_components = self.specification_analyzer.build_risk_components(ai_rule_findings)
        avl_candidates = self.specification_analyzer.build_avl_candidates(normalized_bom)
        vendor_distribution = self._vendor_distribution(normalized_bom)
        statistics = self._statistics(normalized_bom)
        critical_parts = normalized_bom[normalized_bom["Is_Critical"]].copy()
        cost_down = self._cost_down_candidates(
            duplicate_pn,
            near_value,
            different_package,
        )
        review_needed = self._review_needed(
            duplicate_pn,
            near_value,
            different_package,
            different_voltage,
            different_material,
            risk_components,
        )
        score = self._score(
            normalized_bom,
            duplicate_pn,
            near_value,
            critical_parts,
            risk_components,
        )
        summary = self._summary(
            normalized_bom,
            duplicate_pn,
            near_value,
            cost_down,
            critical_parts,
            score,
            different_package,
            different_voltage,
            different_material,
            avl_candidates,
            risk_components,
            merge_candidates,
        )

        return {
            "summary": summary,
            "report_metadata": self._report_metadata(len(normalized_bom), project_name),
            "normalized_bom": normalized_bom,
            "merge_candidates": merge_candidates,
            "duplicate_pn": duplicate_pn,
            "near_value": near_value,
            "near_resistance": near_resistance,
            "near_capacitance": near_capacitance,
            "different_package": different_package,
            "different_voltage": different_voltage,
            "different_material": different_material,
            "avl_candidates": avl_candidates,
            "risk_components": risk_components,
            "ai_rule_findings": ai_rule_findings,
            "rule_library": self.rule_library.to_dataframe(),
            "cost_down": cost_down,
            "vendor_distribution": vendor_distribution,
            "statistics": statistics,
            "review_needed": review_needed,
            "critical_parts": critical_parts,
            "bom_score": score,
        }

    def write_excel_report(self, reports, output_file):
        return ExcelReportWriter().write(reports, output_file)

    def _resolve_columns(self, columns):
        available_columns = list(columns)
        lookup = {str(column).casefold(): column for column in available_columns}
        resolved = {}

        for field, candidates in self.COLUMN_CANDIDATES.items():
            resolved[field] = None
            for candidate in candidates:
                match = lookup.get(candidate.casefold())
                if match is not None:
                    resolved[field] = match
                    break

            if resolved[field] is None:
                for column in available_columns:
                    column_text = str(column).casefold()
                    if any(candidate.casefold() in column_text for candidate in candidates):
                        resolved[field] = column
                        break

        return resolved

    def _normalize_bom(self, dataframe, columns):
        records = []

        for _, row in dataframe.iterrows():
            record = row.to_dict()
            reference = self._row_text(row, columns["reference"])
            component_name = self._row_text(row, columns["component_name"])
            description = self._row_text(row, columns["description"])
            package = self._row_text(row, columns["package"])
            part_number = self._row_text(row, columns["part_number"])
            subsystem = self._row_text(row, columns["subsystem"])
            critical_flag = self._row_text(row, columns["critical"])
            supplied_vendor = self._row_text(row, columns["vendor"])
            supplied_voltage = self._row_text(row, columns["voltage"])
            supplied_material = self._row_text(row, columns["material"])
            supplied_tolerance = self._row_text(row, columns["tolerance"])
            supplied_power = self._row_text(row, columns["power"])
            lifecycle_status = self._row_text(row, columns["lifecycle"])
            bom_comp_type = self._row_text(row, columns["bom_comp_type"])
            component_text = " ".join(
                value for value in (component_name, description) if value
            )

            component_type = self._detect_component(component_text, reference)
            original_value, normalized_value, numeric_value = self._normalize_value(
                component_text,
                component_type,
            )
            voltage = self._normalize_voltage(
                supplied_voltage
            ) or self.value_extractor.extract_voltage(component_text)
            material_column = str(columns["material"] or "").casefold()
            supplied_dielectric = (
                supplied_material.upper() if "dielectric" in material_column else ""
            )
            dielectric = supplied_dielectric or self.value_extractor.extract_dielectric(
                component_text
            )
            material = "" if supplied_dielectric else supplied_material.upper()
            if component_type == "R" and not material:
                material = self.value_extractor.extract_material(component_text, component_type)
            tolerance = supplied_tolerance or self.value_extractor.extract_tolerance(component_text)
            power_rating = supplied_power or self.value_extractor.extract_power_rating(component_text)
            size = self.value_extractor.extract_size(f"{component_text} {package}")
            package_identity = size or package.upper()
            comparison_eligible, specification_field_count, comparison_gap = (
                self._comparison_status(
                    component_type,
                    normalized_value,
                    voltage,
                    dielectric,
                    material,
                    tolerance,
                    power_rating,
                    package_identity,
                )
            )
            normalize_key = self._normalize_key(
                component_type,
                normalized_value,
                voltage,
                dielectric,
                material,
                size,
                tolerance,
                power_rating,
            ) if comparison_eligible else ""
            is_critical, critical_reason = self._critical_status(
                critical_flag,
                subsystem,
                component_text,
            )
            quantity, quantity_valid, quantity_issue = self._parse_quantity(
                row,
                columns["quantity"],
            )

            record.update(
                {
                    "Part_Number": part_number,
                    "Reference": reference,
                    "Component_Type": component_type,
                    "Original_Value": original_value,
                    "Normalized_Value": normalized_value,
                    "Numeric_Value": numeric_value,
                    "Voltage": voltage,
                    "Dielectric": dielectric,
                    "Material": material,
                    "Tolerance": tolerance,
                    "Power_Rating": power_rating,
                    "Size": size,
                    "Package": package,
                    "Package_Identity": package_identity,
                    "Normalize_Key": normalize_key,
                    "Specification_Field_Count": specification_field_count,
                    "Comparison_Eligible": comparison_eligible,
                    "Comparison_Gap": comparison_gap,
                    "Vendor": self._extract_vendor(supplied_vendor, description, component_name),
                    "Subsystem": subsystem,
                    "Quantity_Normalized": quantity,
                    "Quantity_Valid": quantity_valid,
                    "Quantity_Issue": quantity_issue,
                    "Unit_Price": self._numeric_value(row, columns["unit_price"]),
                    "Lifecycle_Status": lifecycle_status,
                    "BOM_Comp_Type": bom_comp_type,
                    "Is_Second_Source": bom_comp_type.strip().casefold() == "s",
                    "Is_Critical": is_critical,
                    "Critical_Reason": critical_reason,
                    "Data_Quality_Score": self._data_quality_score(
                        component_type,
                        normalized_value,
                        package_identity,
                        voltage,
                        dielectric,
                        material,
                        tolerance,
                        power_rating,
                    ),
                }
            )
            records.append(record)

        normalized = pd.DataFrame(records)
        normalized["Review_Status"] = normalized.apply(self._row_review_status, axis=1)
        return normalized

    @staticmethod
    def _row_text(row, column):
        if not column:
            return ""

        value = row[column]
        if pd.isna(value):
            return ""

        return re.sub(r"\s+", " ", str(value)).strip()

    @staticmethod
    def _project_name_from_dataframe(dataframe, columns):
        description_column = columns.get("description") or columns.get("component_name")
        if not description_column or dataframe.empty:
            return ""

        value = dataframe.iloc[0].get(description_column, "")
        if pd.isna(value):
            return ""

        text = re.sub(r"\s+", " ", str(value)).strip()
        return text[:6].upper() if text else ""

    @staticmethod
    def _quantity(row, column):
        return BOMIntelligencePlatform._parse_quantity(row, column)[0]

    @staticmethod
    def _parse_quantity(row, column):
        if not column:
            return 1.0, True, ""

        if pd.isna(row[column]):
            return 1.0, False, "Missing quantity; defaulted to 1"

        source = str(row[column]).strip().replace(",", "")
        try:
            quantity = float(source)
        except (TypeError, ValueError):
            return 1.0, False, f"Invalid quantity {row[column]!r}; defaulted to 1"

        if not math.isfinite(quantity) or quantity < 0:
            return 1.0, False, f"Invalid quantity {row[column]!r}; defaulted to 1"

        return quantity, True, ""

    @staticmethod
    def _numeric_value(row, column):
        if not column or pd.isna(row[column]):
            return None

        value = re.sub(r"[^\d.\-]", "", str(row[column]))
        try:
            numeric_value = float(value)
        except (TypeError, ValueError):
            return None
        return numeric_value if math.isfinite(numeric_value) and numeric_value >= 0 else None

    @staticmethod
    def _first_reference(reference):
        return re.split(r"[\s,;/]+", reference.strip().upper())[0] if reference else ""

    def _detect_component(self, component_text, reference):
        first_reference = self._first_reference(reference)
        text = component_text.upper()

        if re.search(r"\b(?:MLCC|CAPACITOR|CAP)\b", text):
            return "C"

        if re.search(r"\b(?:RESISTOR|RES)\b", text):
            return "R"

        if re.search(r"\b(?:INDUCTOR|IND)\b", text):
            return "L"

        if re.match(r"^C\d", first_reference):
            return "C"

        if re.match(r"^R\d", first_reference):
            return "R"

        if re.match(r"^L\d", first_reference):
            return "L"

        return "UNKNOWN"

    def _normalize_value(self, component_text, component_type):
        raw_value = self.value_extractor.extract(component_text, component_type)

        if component_type == "C":
            normalized_value = self.capacitor_normalizer.normalize(raw_value)
            return raw_value, normalized_value, self._capacitance_to_pf(normalized_value)

        if component_type == "R":
            resistance = self.resistor_normalizer.normalize(raw_value)
            if isinstance(resistance, (int, float)):
                return raw_value, self._format_resistance(resistance), float(resistance)
            return raw_value, raw_value, None

        return raw_value, raw_value, None

    @staticmethod
    def _capacitance_to_pf(value):
        match = re.fullmatch(r"(\d+(?:\.\d+)?)(pF|nF|uF|mF|F)", str(value))
        if not match:
            return None

        magnitude, unit = match.groups()
        multiplier = {
            "pF": 1,
            "nF": 1_000,
            "uF": 1_000_000,
            "mF": 1_000_000_000,
            "F": 1_000_000_000_000,
        }[unit]
        return float(magnitude) * multiplier

    @staticmethod
    def _format_resistance(resistance):
        if 0 < abs(resistance) < 1:
            return f"{resistance * 1_000:g}mOhm"

        if resistance >= 1_000_000:
            return f"{resistance / 1_000_000:g}MOhm"

        if resistance >= 1_000:
            return f"{resistance / 1_000:g}kOhm"

        return f"{resistance:g}Ohm"

    @staticmethod
    def _normalize_key(
        component_type,
        value,
        voltage,
        dielectric,
        material,
        size,
        tolerance="",
        power_rating="",
    ):
        if not value:
            return ""

        if component_type == "C":
            return "_".join(
                (component_type, value, voltage, dielectric, material, tolerance, size)
            )

        if component_type == "R":
            return "_".join((component_type, value, tolerance, power_rating, material, size))

        return "_".join((component_type, value, size))

    @staticmethod
    def _normalize_voltage(value):
        if not value:
            return ""

        embedded_decimal = re.search(r"(?<!\d)(\d+)V(\d+)(?!\d)", str(value).upper())
        if embedded_decimal:
            return f"{float(f'{embedded_decimal.group(1)}.{embedded_decimal.group(2)}'):g}V"

        match = re.search(r"\d+(?:\.\d*)?|\.\d+", str(value))
        return f"{float(match.group(0)):g}V" if match else str(value).strip().upper()

    def _comparison_status(
        self,
        component_type,
        normalized_value,
        voltage,
        dielectric,
        material,
        tolerance,
        power_rating,
        package_identity,
    ):
        if component_type == "C":
            fields = {
                "value": normalized_value,
                "voltage": voltage,
                "dielectric": dielectric,
                "package": package_identity,
                "tolerance": tolerance,
            }
            essential_fields = ("value", "voltage", "dielectric", "package")
        elif component_type == "R":
            fields = {
                "value": normalized_value,
                "package": package_identity,
                "tolerance": tolerance,
                "power": power_rating,
                "material": material,
            }
            essential_fields = ("value", "package", "tolerance", "power")
        else:
            return False, 0, "Component type is not comparison-enabled"

        completed = sum(bool(str(value).strip()) for value in fields.values())
        minimum_fields = int(
            self.rule_library.get(
                "analysis",
                "minimum_complete_spec_fields",
                default=4,
            )
        )
        missing_essential = [name for name in essential_fields if not fields[name]]
        eligible = not missing_essential and completed >= minimum_fields
        gaps = []
        if missing_essential:
            gaps.append(f"Missing required: {', '.join(missing_essential)}")
        if completed < minimum_fields:
            gaps.append(f"Only {completed} specification fields; require {minimum_fields}")
        return eligible, completed, "; ".join(gaps)

    @staticmethod
    def _data_quality_score(
        component_type,
        normalized_value,
        package_identity,
        voltage,
        dielectric,
        material,
        tolerance,
        power_rating,
    ):
        if component_type not in {"R", "C"}:
            return 100

        fields = [normalized_value, package_identity, tolerance, power_rating]
        if component_type == "C":
            fields = [normalized_value, package_identity, voltage, dielectric, tolerance]
        elif material:
            fields.append(material)
        completed = sum(bool(str(value).strip()) for value in fields)
        return round(completed / len(fields) * 100)

    def _critical_status(self, explicit_flag, subsystem, component_name):
        flag = explicit_flag.casefold()
        if flag and flag not in {
            "0",
            "false",
            "no",
            "n",
            "none",
            "nan",
            "na",
            "n/a",
            "-",
        }:
            return True, "Marked as critical in source BOM"

        subsystem_upper = subsystem.upper()
        for sensitive_subsystem in self.SENSITIVE_SUBSYSTEMS:
            if re.search(
                rf"(?<![A-Z0-9]){re.escape(sensitive_subsystem)}(?![A-Z0-9])",
                subsystem_upper,
            ):
                return True, f"Sensitive subsystem: {subsystem}"

        if re.search(r"\b(TX|RX|RF|TYPE[- ]?C|PCIE|DDR|USB)\b", component_name.upper()):
            return True, "High-speed or RF signal context"

        return False, ""

    def _extract_vendor(self, supplied_vendor, description, component_name):
        if supplied_vendor:
            return supplied_vendor

        source_text = f"{description} {component_name}".upper()
        for vendor in self.KNOWN_VENDORS:
            if vendor in source_text:
                return vendor.title() if vendor != "TDK" else vendor

        slash_prefix = re.match(r"\s*([A-Z][A-Z0-9 .-]{1,30})/", description.upper())
        if slash_prefix:
            return slash_prefix.group(1).strip().title()

        return "Unknown"

    @staticmethod
    def _row_review_status(row):
        if row["Is_Critical"]:
            return "Protected - retain design control"

        if row["Component_Type"] in {"C", "R"}:
            return "Analyzed" if row["Comparison_Eligible"] else "Review - incomplete specification"

        return "Informational"

    def _find_duplicate_specs(self, normalized_bom):
        columns = [
            "Group_ID",
            "Component_Type",
            "Normalize_Key",
            "Normalized_Value",
            "Voltage",
            "Dielectric",
            "Material",
            "Tolerance",
            "Power_Rating",
            "Size",
            "Package",
            "Part_Number_Count",
            "Part_Numbers",
            "Vendor_Count",
            "Vendors",
            "BOM_Lines",
            "Total_Quantity",
            "Min_Unit_Price",
            "Max_Unit_Price",
            "Estimated_BOM_Savings",
            "References",
            "Has_Critical_Part",
            "Recommendation",
        ]
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["C", "R"])
            & normalized_bom["Normalize_Key"].ne("")
        ]
        groups = []

        for _, group in candidates.groupby(["Component_Type", "Normalize_Key"], sort=True):
            part_numbers = self._unique_text(group["Part_Number"])
            if len(part_numbers) < 2:
                continue

            is_critical = bool(group["Is_Critical"].any())
            groups.append(
                {
                    "Component_Type": group.iloc[0]["Component_Type"],
                    "Normalize_Key": group.iloc[0]["Normalize_Key"],
                    "Normalized_Value": group.iloc[0]["Normalized_Value"],
                    "Voltage": group.iloc[0]["Voltage"],
                    "Dielectric": group.iloc[0]["Dielectric"],
                    "Material": group.iloc[0]["Material"],
                    "Tolerance": group.iloc[0]["Tolerance"],
                    "Power_Rating": group.iloc[0]["Power_Rating"],
                    "Size": group.iloc[0]["Size"],
                    "Package": group.iloc[0]["Package"],
                    "Part_Number_Count": len(part_numbers),
                    "Part_Numbers": ", ".join(part_numbers),
                    "Vendor_Count": len(
                        self._unique_text(group["Vendor"], case_insensitive=True)
                    ),
                    "Vendors": ", ".join(
                        self._unique_text(group["Vendor"], case_insensitive=True)
                    ),
                    "BOM_Lines": len(group),
                    "Total_Quantity": group["Quantity_Normalized"].sum(),
                    "Min_Unit_Price": self._minimum_positive(group["Unit_Price"]),
                    "Max_Unit_Price": self._maximum_positive(group["Unit_Price"]),
                    "Estimated_BOM_Savings": self.specification_analyzer._estimated_bom_savings(group),
                    "References": ", ".join(self._unique_text(group["Reference"])),
                    "Has_Critical_Part": "Yes" if is_critical else "No",
                    "Recommendation": (
                        "Protected specification: review only with design owner approval"
                        if is_critical
                        else "Same electrical specification uses multiple PNs; review consolidation"
                    ),
                }
            )

        for index, group in enumerate(groups, start=1):
            group["Group_ID"] = f"DUP-{index:03d}"

        return pd.DataFrame(groups, columns=columns)

    def _find_merge_candidates(self, normalized_bom):
        columns = [
            "Candidate_ID",
            "Priority",
            "Priority_Stars",
            "Merge_Score",
            "Quantity_Score",
            "Spec_Similarity",
            "Engineering_Risk_Score",
            "Group_ID",
            "Component_Type",
            "Item",
            "Current_PN",
            "Current_Qty",
            "Current_Vendor",
            "Target_PN",
            "Target_Qty",
            "Target_Vendor",
            "Normalize_Key",
            "Normalized_Value",
            "Voltage",
            "Dielectric",
            "Tolerance",
            "Size",
            "Reason",
            "Recommendation",
        ]
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["C", "R"])
            & normalized_bom["Normalize_Key"].ne("")
            & normalized_bom["Part_Number"].ne("")
            & ~normalized_bom.get("Is_Second_Source", pd.Series(False, index=normalized_bom.index))
        ]
        rows = []
        candidate_quantity_limit = float(
            self.rule_library.get(
                "merge_opportunity",
                "candidate_quantity_limit",
                default=10,
            )
        )

        for group_index, (_, group) in enumerate(
            candidates.groupby(["Component_Type", "Normalize_Key"], sort=True),
            start=1,
        ):
            part_groups = self._part_number_merge_groups(group)
            if len(part_groups) < 2:
                continue

            target = max(
                part_groups,
                key=lambda item: (item["quantity"], item["line_count"], item["part_number"]),
            )

            for current in part_groups:
                if current["part_number"] == target["part_number"]:
                    continue
                if current["quantity"] > candidate_quantity_limit:
                    continue

                quantity_score = self._merge_quantity_score(current["quantity"])
                spec_similarity = self._spec_similarity_score(current["representative"], target["representative"])
                engineering_risk_score = self._engineering_risk_score(current, target)
                merge_score = self._merge_score(
                    quantity_score,
                    spec_similarity,
                    engineering_risk_score,
                )
                priority, stars = self._merge_priority(merge_score)
                representative = current["representative"]
                rows.append(
                    {
                        "Priority": priority,
                        "Priority_Stars": stars,
                        "Merge_Score": merge_score,
                        "Quantity_Score": quantity_score,
                        "Spec_Similarity": spec_similarity,
                        "Engineering_Risk_Score": engineering_risk_score,
                        "Group_ID": f"MERGE-G{group_index:03d}",
                        "Component_Type": representative["Component_Type"],
                        "Item": current["references"],
                        "Current_PN": current["part_number"],
                        "Current_Qty": current["quantity"],
                        "Current_Vendor": current["vendor"],
                        "Target_PN": target["part_number"],
                        "Target_Qty": target["quantity"],
                        "Target_Vendor": target["vendor"],
                        "Normalize_Key": representative["Normalize_Key"],
                        "Normalized_Value": representative["Normalized_Value"],
                        "Voltage": representative["Voltage"],
                        "Dielectric": representative["Dielectric"],
                        "Tolerance": representative["Tolerance"],
                        "Size": representative["Size"],
                        "Reason": self._merge_reason(spec_similarity, target["quantity"]),
                        "Recommendation": f"Merge {current['part_number']} -> {target['part_number']}",
                    }
                )

        rows.sort(
            key=lambda item: (
                item["Priority"],
                -item["Merge_Score"],
                item["Current_Qty"],
                item["Current_PN"],
            )
        )
        for index, row in enumerate(rows, start=1):
            row["Candidate_ID"] = f"MERGE-{index:04d}"

        return pd.DataFrame(rows, columns=columns)

    def _part_number_merge_groups(self, group):
        part_groups = []
        for part_number, part_group in group.groupby("Part_Number", sort=False):
            representative = part_group.iloc[0]
            part_groups.append(
                {
                    "part_number": str(part_number),
                    "quantity": float(part_group["Quantity_Normalized"].sum()),
                    "line_count": len(part_group),
                    "vendor": ", ".join(self._unique_text(part_group["Vendor"], case_insensitive=True)),
                    "references": ", ".join(self._unique_text(part_group["Reference"])),
                    "critical": bool(part_group["Is_Critical"].any()),
                    "lifecycle_statuses": self._unique_text(part_group["Lifecycle_Status"]),
                    "representative": representative,
                }
            )
        return part_groups

    def _merge_quantity_score(self, quantity):
        bands = self.rule_library.get(
            "merge_opportunity",
            "quantity_score_bands",
            default=[],
        )
        for band in bands:
            if not isinstance(band, dict):
                continue
            maximum = band.get("max_quantity")
            score = band.get("score")
            if maximum is None or float(quantity) <= float(maximum):
                return int(score)
        return 20

    @staticmethod
    def _spec_similarity_score(current, target):
        if current["Normalized_Value"] != target["Normalized_Value"]:
            return 0
        if current["Package_Identity"] != target["Package_Identity"]:
            return 40
        if current["Voltage"] != target["Voltage"]:
            return 70
        if current["Tolerance"] != target["Tolerance"]:
            return 60
        if str(current["Vendor"]).casefold() != str(target["Vendor"]).casefold():
            return 95
        return 100

    @staticmethod
    def _engineering_risk_score(current, target):
        if current["critical"] or target["critical"]:
            return 40
        lifecycle_text = " ".join(current["lifecycle_statuses"] + target["lifecycle_statuses"]).upper()
        if any(keyword in lifecycle_text for keyword in ("EOL", "OBSOLETE", "DISCONTINUED", "NRND")):
            return 70
        return 100

    def _merge_score(self, quantity_score, spec_similarity, engineering_risk_score):
        weights = self.rule_library.get(
            "merge_opportunity",
            "score_weights",
            default={},
        )
        quantity_weight = float(weights.get("quantity", 0.45))
        spec_weight = float(weights.get("spec_similarity", 0.40))
        risk_weight = float(weights.get("engineering_risk", 0.15))
        total_weight = quantity_weight + spec_weight + risk_weight
        if total_weight <= 0:
            return 0
        score = (
            quantity_score * quantity_weight
            + spec_similarity * spec_weight
            + engineering_risk_score * risk_weight
        ) / total_weight
        return round(score, 1)

    @staticmethod
    def _merge_priority(merge_score):
        if merge_score >= 95:
            return "Priority 1", "★★★★★"
        if merge_score >= 85:
            return "Priority 2", "★★★★☆"
        return "Priority 3", "★★★☆☆"

    @staticmethod
    def _merge_reason(spec_similarity, target_quantity):
        if spec_similarity == 100:
            similarity_text = "Exact same specification"
        elif spec_similarity == 95:
            similarity_text = "Only vendor is different"
        elif spec_similarity == 70:
            similarity_text = "Voltage is different"
        elif spec_similarity == 60:
            similarity_text = "Tolerance is different"
        elif spec_similarity == 40:
            similarity_text = "Size is different"
        else:
            similarity_text = "Value is different"
        return f"{similarity_text}; target PN already has {target_quantity:g} pcs"

    def _find_near_values(self, normalized_bom):
        columns = [
            "Pair_ID",
            "Component_Type",
            "Value_A",
            "Value_B",
            "Ratio",
            "Difference_Percent",
            "Voltage",
            "Dielectric",
            "Material",
            "Tolerance",
            "Power_Rating",
            "Size",
            "Part_Numbers_A",
            "Part_Numbers_B",
            "Quantity_A",
            "Quantity_B",
            "References_A",
            "References_B",
            "Has_Critical_Part",
            "Recommendation",
        ]
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["C", "R"])
            & normalized_bom["Normalize_Key"].ne("")
            & normalized_bom["Numeric_Value"].notna()
        ].copy()

        if candidates.empty:
            return pd.DataFrame(columns=columns)

        candidates["Family_Key"] = candidates.apply(self._family_key, axis=1)
        results = []

        for _, family in candidates.groupby("Family_Key", sort=True):
            values = []
            for value, value_group in family.groupby("Normalized_Value", sort=False):
                values.append(
                    {
                        "value": value,
                        "numeric": value_group.iloc[0]["Numeric_Value"],
                        "quantity": value_group["Quantity_Normalized"].sum(),
                        "references": ", ".join(self._unique_text(value_group["Reference"])),
                        "critical": bool(value_group["Is_Critical"].any()),
                        "component_type": value_group.iloc[0]["Component_Type"],
                        "voltage": value_group.iloc[0]["Voltage"],
                        "dielectric": value_group.iloc[0]["Dielectric"],
                        "material": value_group.iloc[0]["Material"],
                        "tolerance": value_group.iloc[0]["Tolerance"],
                        "power_rating": value_group.iloc[0]["Power_Rating"],
                        "size": value_group.iloc[0]["Size"],
                        "part_numbers": ", ".join(self._unique_text(value_group["Part_Number"])),
                    }
                )

            for value_a, value_b in itertools.combinations(sorted(values, key=lambda item: item["numeric"]), 2):
                if value_a["numeric"] == 0:
                    continue

                ratio = value_b["numeric"] / value_a["numeric"]
                ratio_limit = self.near_value_ratios[value_a["component_type"]]
                if ratio > ratio_limit:
                    continue

                is_critical = value_a["critical"] or value_b["critical"]
                difference_percent = (
                    (value_b["numeric"] - value_a["numeric"]) / value_a["numeric"] * 100
                )
                results.append(
                    {
                        "Component_Type": value_a["component_type"],
                        "Value_A": value_a["value"],
                        "Value_B": value_b["value"],
                        "Ratio": round(ratio, 3),
                        "Difference_Percent": round(difference_percent, 2),
                        "Voltage": value_a["voltage"],
                        "Dielectric": value_a["dielectric"],
                        "Material": value_a["material"],
                        "Tolerance": value_a["tolerance"],
                        "Power_Rating": value_a["power_rating"],
                        "Size": value_a["size"],
                        "Part_Numbers_A": value_a["part_numbers"],
                        "Part_Numbers_B": value_b["part_numbers"],
                        "Quantity_A": value_a["quantity"],
                        "Quantity_B": value_b["quantity"],
                        "References_A": value_a["references"],
                        "References_B": value_b["references"],
                        "Has_Critical_Part": "Yes" if is_critical else "No",
                        "Recommendation": (
                            "Protected circuit: confirm circuit requirement; do not auto-consolidate"
                            if is_critical
                            else "Similar values detected; confirm circuit requirement before consolidation"
                        ),
                    }
                )

        for index, result in enumerate(results, start=1):
            result["Pair_ID"] = f"NEAR-{index:03d}"

        return pd.DataFrame(results, columns=columns)

    @staticmethod
    def _family_key(row):
        if row["Component_Type"] == "C":
            return "|".join(
                (
                    row["Component_Type"],
                    row["Voltage"],
                    row["Dielectric"],
                    row["Material"],
                    row["Tolerance"],
                    row["Package_Identity"],
                )
            )

        return "|".join(
            (
                row["Component_Type"],
                row["Material"],
                row["Tolerance"],
                row["Power_Rating"],
                row["Package_Identity"],
            )
        )

    def _vendor_distribution(self, normalized_bom):
        columns = [
            "Component_Type",
            "Normalize_Key",
            "Normalized_Value",
            "Vendor",
            "Part_Number_Count",
            "BOM_Lines",
            "Total_Quantity",
            "Part_Numbers",
            "References",
        ]
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["C", "R"])
            & normalized_bom["Normalize_Key"].ne("")
        ]
        results = []

        for keys, group in candidates.groupby(
            ["Component_Type", "Normalize_Key", "Normalized_Value", "Vendor"],
            sort=True,
        ):
            component_type, normalize_key, normalized_value, vendor = keys
            part_numbers = self._unique_text(group["Part_Number"])
            results.append(
                {
                    "Component_Type": component_type,
                    "Normalize_Key": normalize_key,
                    "Normalized_Value": normalized_value,
                    "Vendor": vendor,
                    "Part_Number_Count": len(part_numbers),
                    "BOM_Lines": len(group),
                    "Total_Quantity": group["Quantity_Normalized"].sum(),
                    "Part_Numbers": ", ".join(part_numbers),
                    "References": ", ".join(self._unique_text(group["Reference"])),
                }
            )

        return pd.DataFrame(results, columns=columns)

    def _statistics(self, normalized_bom):
        columns = [
            "Component_Type",
            "Normalized_Value",
            "Specification_Count",
            "BOM_Lines",
            "Total_Quantity",
            "Part_Number_Count",
            "Vendor_Count",
            "References",
        ]
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["C", "R"])
            & normalized_bom["Normalized_Value"].ne("")
        ]
        results = []

        for keys, group in candidates.groupby(["Component_Type", "Normalized_Value"], sort=True):
            component_type, normalized_value = keys
            results.append(
                {
                    "Component_Type": component_type,
                    "Normalized_Value": normalized_value,
                    "Specification_Count": group["Normalize_Key"].nunique(),
                    "BOM_Lines": len(group),
                    "Total_Quantity": group["Quantity_Normalized"].sum(),
                    "Part_Number_Count": len(self._unique_text(group["Part_Number"])),
                    "Vendor_Count": len(
                        self._unique_text(group["Vendor"], case_insensitive=True)
                    ),
                    "References": ", ".join(self._unique_text(group["Reference"])),
                }
            )

        return pd.DataFrame(results, columns=columns)

    def _cost_down_candidates(self, duplicate_pn, near_value, different_package):
        columns = [
            "Candidate_ID",
            "Candidate_Type",
            "Source_ID",
            "Priority",
            "Opportunity_Score",
            "Component_Type",
            "Specification",
            "Impact_Quantity",
            "Estimated_BOM_Savings",
            "Confidence",
            "Business_Rationale",
            "Recommendation",
        ]
        results = []

        for _, group in duplicate_pn[duplicate_pn["Has_Critical_Part"] == "No"].iterrows():
            priority = int(self.rule_library.get("cost_down", "same_spec_priority", default=90))
            results.append(
                {
                    "Candidate_Type": "Same specification",
                    "Source_ID": group["Group_ID"],
                    "Priority": "High",
                    "Opportunity_Score": self._opportunity_score(
                        priority,
                        group["Total_Quantity"],
                        group["Estimated_BOM_Savings"],
                    ),
                    "Component_Type": group["Component_Type"],
                    "Specification": group["Normalize_Key"],
                    "Impact_Quantity": group["Total_Quantity"],
                    "Estimated_BOM_Savings": group["Estimated_BOM_Savings"],
                    "Confidence": self.rule_library.get("confidence", "exact_specification"),
                    "Business_Rationale": "Aggregate demand and remove equivalent MPN fragmentation.",
                    "Recommendation": "Consolidate equivalent PNs after AVL and lifecycle review",
                }
            )

        for _, pair in near_value[near_value["Has_Critical_Part"] == "No"].iterrows():
            priority = int(self.rule_library.get("cost_down", "near_value_priority", default=60))
            results.append(
                {
                    "Candidate_Type": "Near value",
                    "Source_ID": pair["Pair_ID"],
                    "Priority": "Medium",
                    "Opportunity_Score": self._opportunity_score(
                        priority,
                        pair["Quantity_A"] + pair["Quantity_B"],
                    ),
                    "Component_Type": pair["Component_Type"],
                    "Specification": f"{pair['Value_A']} / {pair['Value_B']}",
                    "Impact_Quantity": pair["Quantity_A"] + pair["Quantity_B"],
                    "Estimated_BOM_Savings": 0.0,
                    "Confidence": self.rule_library.get("confidence", "near_value"),
                    "Business_Rationale": "Reduce low-value specification variety and improve purchasing leverage.",
                    "Recommendation": "Confirm circuit margin before standardizing a value",
                }
            )

        for _, variant in different_package[
            different_package["Has_Critical_Part"] == "No"
        ].iterrows():
            priority = int(self.rule_library.get("cost_down", "package_priority", default=45))
            results.append(
                {
                    "Candidate_Type": "Package standardization",
                    "Source_ID": variant["Variant_ID"],
                    "Priority": "Low",
                    "Opportunity_Score": self._opportunity_score(
                        priority,
                        variant["Total_Quantity"],
                    ),
                    "Component_Type": variant["Component_Type"],
                    "Specification": (
                        f"{variant['Normalized_Value']} / {variant['Attribute_Values']}"
                    ),
                    "Impact_Quantity": variant["Total_Quantity"],
                    "Estimated_BOM_Savings": 0.0,
                    "Confidence": variant["Confidence"],
                    "Business_Rationale": "Reduce package variety, feeders, and inventory complexity.",
                    "Recommendation": variant["Recommendation"],
                }
            )

        minimum_quantity = float(
            self.rule_library.get("cost_down", "minimum_impact_quantity", default=1)
        )
        results = [
            result for result in results if result["Impact_Quantity"] >= minimum_quantity
        ]
        results.sort(key=lambda item: item["Opportunity_Score"], reverse=True)

        for index, result in enumerate(results, start=1):
            result["Candidate_ID"] = f"CD-{index:03d}"

        return pd.DataFrame(results, columns=columns)

    def _review_needed(
        self,
        duplicate_pn,
        near_value,
        different_package,
        different_voltage,
        different_material,
        risk_components,
    ):
        columns = [
            "Review_ID",
            "Category",
            "Priority",
            "Source_ID",
            "Finding",
            "Confidence",
            "Recommendation",
        ]
        results = []

        for _, group in duplicate_pn.iterrows():
            priority = 1 if group["Has_Critical_Part"] == "Yes" else 3
            results.append(
                {
                    "Category": "Duplicate specification",
                    "Priority": priority,
                    "Source_ID": group["Group_ID"],
                    "Finding": f"{group['Part_Number_Count']} PNs share {group['Normalize_Key']}",
                    "Confidence": self.rule_library.get("confidence", "exact_specification"),
                    "Recommendation": group["Recommendation"],
                }
            )

        for _, pair in near_value.iterrows():
            priority = 1 if pair["Has_Critical_Part"] == "Yes" else 3
            results.append(
                {
                    "Category": "Near value",
                    "Priority": priority,
                    "Source_ID": pair["Pair_ID"],
                    "Finding": f"{pair['Value_A']} and {pair['Value_B']} are within {pair['Ratio']}:1",
                    "Confidence": self.rule_library.get("confidence", "near_value"),
                    "Recommendation": pair["Recommendation"],
                }
            )

        for variants in (different_package, different_voltage, different_material):
            for _, variant in variants.iterrows():
                priority = 1 if variant["Has_Critical_Part"] == "Yes" else 3
                results.append(
                    {
                        "Category": f"Different {variant['Compared_Attribute'].lower()}",
                        "Priority": priority,
                        "Source_ID": variant["Variant_ID"],
                        "Finding": (
                            f"{variant['Normalized_Value']} uses {variant['Attribute_Values']}"
                        ),
                        "Confidence": variant["Confidence"],
                        "Recommendation": variant["Recommendation"],
                    }
                )

        for _, risk in risk_components.iterrows():
            results.append(
                {
                    "Category": risk["Risk_Category"],
                    "Priority": 1 if risk["Severity"] == "Critical" else 2,
                    "Source_ID": risk["Risk_ID"],
                    "Finding": risk["Finding"],
                    "Confidence": self.rule_library.get("confidence", "structured_attribute"),
                    "Recommendation": risk["Recommendation"],
                }
            )

        results.sort(key=lambda item: item["Priority"])
        for index, result in enumerate(results, start=1):
            result["Review_ID"] = f"REV-{index:03d}"

        return pd.DataFrame(results, columns=columns)

    def _score(self, normalized_bom, duplicate_pn, near_value, critical_parts, risk_components):
        capacitor_specs = normalized_bom.loc[
            (normalized_bom["Component_Type"] == "C")
            & normalized_bom["Normalize_Key"].ne(""),
            "Normalize_Key",
        ].nunique()
        resistor_specs = normalized_bom.loc[
            (normalized_bom["Component_Type"] == "R")
            & normalized_bom["Normalize_Key"].ne(""),
            "Normalize_Key",
        ].nunique()
        noncritical_duplicates = len(duplicate_pn[duplicate_pn["Has_Critical_Part"] == "No"])
        noncritical_near_values = len(near_value[near_value["Has_Critical_Part"] == "No"])
        complexity_score = min(
            100,
            round(20 + capacitor_specs * 0.8 + resistor_specs * 0.5 + len(near_value) * 0.8),
        )
        health_penalty = min(30, noncritical_duplicates * 3)
        health_penalty += min(20, noncritical_near_values * 2)
        health_penalty += min(20, max(0, capacitor_specs - 30) // 2)
        health_penalty += min(15, len(risk_components) * 2)
        overall_score = max(0, 100 - health_penalty)

        rows = [
            ("Overall Score", overall_score, "100 is most consolidated; protected circuits are excluded from cost-down penalties"),
            ("BOM Complexity Score", complexity_score, "Higher score indicates more specification variety to maintain"),
            ("MLCC Specifications", capacitor_specs, "Unique normalized capacitor specifications"),
            ("Resistor Specifications", resistor_specs, "Unique normalized resistor specifications"),
            ("Duplicate PN Groups", len(duplicate_pn), "Equivalent specifications using multiple PNs"),
            ("Near Value Pairs", len(near_value), "Values within the configured similarity ratio"),
            ("Critical Component Lines", len(critical_parts), "Protected from automatic cost-down suggestions"),
            ("High Risk Findings", len(risk_components), "Critical and high-severity rule findings"),
            (
                "Average Data Quality",
                round(float(normalized_bom["Data_Quality_Score"].mean()), 1),
                "Completeness of normalized engineering attributes",
            ),
        ]
        return pd.DataFrame(rows, columns=["Metric", "Value", "Interpretation"])

    def _summary(
        self,
        normalized_bom,
        duplicate_pn,
        near_value,
        cost_down,
        critical_parts,
        score,
        different_package,
        different_voltage,
        different_material,
        avl_candidates,
        risk_components,
        merge_candidates,
    ):
        capacitor_rows = normalized_bom[normalized_bom["Component_Type"] == "C"]
        overall_score = score.loc[score["Metric"] == "Overall Score", "Value"].iloc[0]
        rows = [
            ("BOM Lines", len(normalized_bom)),
            ("Total Component Quantity", normalized_bom["Quantity_Normalized"].sum()),
            ("MLCC Lines", len(capacitor_rows)),
            ("MLCC Quantity", capacitor_rows["Quantity_Normalized"].sum()),
            ("MLCC Specifications", capacitor_rows["Normalize_Key"].replace("", pd.NA).nunique()),
            ("Duplicate PN Groups", len(duplicate_pn)),
            ("Near Value Pairs", len(near_value)),
            ("Near Resistance Pairs", len(near_value[near_value["Component_Type"] == "R"])),
            ("Near Capacitance Pairs", len(near_value[near_value["Component_Type"] == "C"])),
            ("Different Package Groups", len(different_package)),
            ("Different Voltage Groups", len(different_voltage)),
            ("Different Material Groups", len(different_material)),
            (
                "Unified AVL Ready",
                len(avl_candidates[avl_candidates["AVL_Readiness"] == "Ready for unified AVL"]),
            ),
            ("Cost Down Candidates", len(cost_down)),
            ("Top Merge Candidates", len(merge_candidates)),
            ("Estimated BOM Savings", round(float(cost_down["Estimated_BOM_Savings"].sum()), 4)),
            ("High Risk Findings", len(risk_components)),
            ("Critical Component Lines", len(critical_parts)),
            ("Detected Vendors", normalized_bom["Vendor"].replace("Unknown", pd.NA).nunique()),
            ("Average Data Quality", round(float(normalized_bom["Data_Quality_Score"].mean()), 1)),
            ("Overall Score", overall_score),
        ]
        return pd.DataFrame(rows, columns=["Metric", "Value"])

    def _report_metadata(self, row_count, project_name=""):
        rows = [
            ("Platform Version", self.VERSION),
            ("Generated UTC", datetime.now(timezone.utc).isoformat(timespec="seconds")),
            ("Project Name", project_name),
            ("Rule Source", self.rule_library.source),
            ("Rule Schema Version", self.rule_library.get("schema_version")),
            ("Resistor Near-Value Ratio", self.near_value_ratios["R"]),
            ("Capacitor Near-Value Ratio", self.near_value_ratios["C"]),
            ("Analyzed Rows", row_count),
        ]
        return pd.DataFrame(rows, columns=["Property", "Value"])

    @staticmethod
    def _minimum_positive(series):
        values = pd.to_numeric(series, errors="coerce")
        values = values[values.gt(0)]
        return float(values.min()) if not values.empty else None

    @staticmethod
    def _maximum_positive(series):
        values = pd.to_numeric(series, errors="coerce")
        values = values[values.gt(0)]
        return float(values.max()) if not values.empty else None

    @staticmethod
    def _opportunity_score(base_score, quantity, estimated_savings=0):
        quantity_bonus = min(8, max(0, float(quantity)) ** 0.5)
        savings_bonus = min(12, max(0, float(estimated_savings)) ** 0.5 * 2)
        return round(min(100, base_score + quantity_bonus + savings_bonus), 1)

    @staticmethod
    def _unique_text(series, case_insensitive=False):
        values = []
        seen = set()
        for value in series:
            text = "" if pd.isna(value) else str(value).strip()
            identity = text.casefold() if case_insensitive else text
            if text and identity not in seen:
                values.append(text)
                seen.add(identity)
        return values


__all__ = ["BOMIntelligencePlatform"]

