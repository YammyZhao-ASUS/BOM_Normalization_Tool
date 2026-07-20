from collections import OrderedDict

import pandas as pd


class SpecificationAnalyzer:
    """Compare normalized specifications across package, voltage, and material."""

    VARIANT_COLUMNS = [
        "Variant_ID",
        "Component_Type",
        "Normalized_Value",
        "Compared_Attribute",
        "Attribute_Values",
        "Specification_Context",
        "Part_Number_Count",
        "Part_Numbers",
        "Vendor_Count",
        "Vendors",
        "Total_Quantity",
        "References",
        "Has_Critical_Part",
        "Confidence",
        "Recommendation",
    ]

    AVL_COLUMNS = [
        "AVL_ID",
        "Component_Type",
        "Normalize_Key",
        "Normalized_Value",
        "Preferred_PN",
        "Part_Number_Count",
        "Part_Numbers",
        "Vendor_Count",
        "Vendors",
        "Approved_Vendors",
        "Total_Quantity",
        "Estimated_BOM_Savings",
        "Has_Critical_Part",
        "AVL_Readiness",
        "Qualification_Gap",
        "Recommendation",
    ]

    RISK_COLUMNS = [
        "Risk_ID",
        "Severity",
        "Risk_Category",
        "Rule_ID",
        "Part_Number",
        "Reference",
        "Component_Type",
        "Finding",
        "Evidence",
        "Recommendation",
    ]

    ATTRIBUTE_LABELS = {
        "Package_Identity": "Package",
        "Voltage": "Voltage",
        "Dielectric": "Material / Dielectric",
    }

    def __init__(self, rule_library):
        self.rule_library = rule_library

    def find_package_variants(self, normalized_bom):
        return self._find_variants(normalized_bom, "Package_Identity", "PKG")

    def find_voltage_variants(self, normalized_bom):
        return self._find_variants(normalized_bom, "Voltage", "VOLT")

    def find_material_variants(self, normalized_bom):
        return self._find_variants(normalized_bom, "Dielectric", "MAT")

    def build_avl_candidates(self, normalized_bom):
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["R", "C"])
            & normalized_bom["Normalize_Key"].fillna("").ne("")
            & normalized_bom["Comparison_Eligible"]
        ]
        minimum_vendors = int(self.rule_library.get("avl", "minimum_vendors", default=2))
        approved_policy = {
            vendor.casefold()
            for vendor in self.rule_library.get("avl", "approved_vendors", default=[])
        }
        blocked_policy = {
            vendor.casefold()
            for vendor in self.rule_library.get("avl", "blocked_vendors", default=[])
        }
        rows = []

        for normalize_key, group in candidates.groupby("Normalize_Key", sort=True):
            part_numbers = self._unique_text(group["Part_Number"])
            vendors = self._unique_text(group["Vendor"], excluded={"Unknown"})
            approved_vendors = (
                [vendor for vendor in vendors if vendor.casefold() in approved_policy]
                if approved_policy
                else vendors
            )
            blocked_vendors = [vendor for vendor in vendors if vendor.casefold() in blocked_policy]
            is_critical = bool(group["Is_Critical"].any())
            selectable_group = group[
                ~group["Vendor"].fillna("").astype(str).str.casefold().isin(blocked_policy)
            ]
            if approved_policy:
                selectable_group = selectable_group[
                    selectable_group["Vendor"]
                    .fillna("")
                    .astype(str)
                    .str.casefold()
                    .isin(approved_policy)
                ]
            preferred_pn = self._preferred_part_number(selectable_group)
            estimated_savings = self._estimated_bom_savings(group)

            if blocked_vendors:
                readiness = "Blocked source present"
                gap = f"Blocked vendor: {', '.join(blocked_vendors)}"
                recommendation = "Replace the blocked source or document an approved exception."
            elif is_critical:
                readiness = "Design-controlled review"
                gap = "Design owner approval required"
                recommendation = "Validate alternates electrically before updating the AVL."
            elif approved_policy and len(approved_vendors) < minimum_vendors:
                readiness = "Qualification required"
                gap = f"Only {len(approved_vendors)} approved vendor(s); target is {minimum_vendors}"
                recommendation = "Qualify additional approved manufacturers for this specification."
            elif len(vendors) < minimum_vendors:
                readiness = "Single-source gap"
                gap = f"Only {len(vendors)} known vendor(s); target is {minimum_vendors}"
                recommendation = "Source and qualify at least one alternate manufacturer."
            elif len(part_numbers) > 1:
                readiness = "Ready for unified AVL"
                gap = ""
                recommendation = "Review the suggested preferred PN and publish one controlled AVL entry."
            else:
                readiness = "AVL baseline"
                gap = "No alternate MPN recorded"
                recommendation = "Keep the current source and add a qualified alternate when available."

            representative = group.iloc[0]
            rows.append(
                {
                    "Component_Type": representative["Component_Type"],
                    "Normalize_Key": normalize_key,
                    "Normalized_Value": representative["Normalized_Value"],
                    "Preferred_PN": preferred_pn,
                    "Part_Number_Count": len(part_numbers),
                    "Part_Numbers": ", ".join(part_numbers),
                    "Vendor_Count": len(vendors),
                    "Vendors": ", ".join(vendors),
                    "Approved_Vendors": ", ".join(approved_vendors),
                    "Total_Quantity": float(group["Quantity_Normalized"].sum()),
                    "Estimated_BOM_Savings": estimated_savings,
                    "Has_Critical_Part": "Yes" if is_critical else "No",
                    "AVL_Readiness": readiness,
                    "Qualification_Gap": gap,
                    "Recommendation": recommendation,
                }
            )

        for index, row in enumerate(rows, start=1):
            row["AVL_ID"] = f"AVL-{index:04d}"

        return pd.DataFrame(rows, columns=self.AVL_COLUMNS)

    def build_risk_components(self, rule_findings):
        rows = []
        high_risk = rule_findings[rule_findings["Severity"].isin(["Critical", "High"])]

        for _, finding in high_risk.iterrows():
            rule_id = finding["Rule_ID"]
            rows.append(
                {
                    "Severity": finding["Severity"],
                    "Risk_Category": self._risk_category(rule_id),
                    "Rule_ID": rule_id,
                    "Part_Number": finding["Part_Number"],
                    "Reference": finding["Reference"],
                    "Component_Type": finding["Component_Type"],
                    "Finding": finding["Finding"],
                    "Evidence": finding["Evidence"],
                    "Recommendation": finding["Recommendation"],
                }
            )

        severity_order = {"Critical": 0, "High": 1}
        rows.sort(key=lambda item: (severity_order.get(item["Severity"], 9), item["Rule_ID"]))
        for index, row in enumerate(rows, start=1):
            row["Risk_ID"] = f"RISK-{index:04d}"

        return pd.DataFrame(rows, columns=self.RISK_COLUMNS)

    def _find_variants(self, normalized_bom, target_attribute, identifier_prefix):
        candidates = normalized_bom[
            normalized_bom["Component_Type"].isin(["R", "C"])
            & normalized_bom["Normalized_Value"].fillna("").ne("")
            & normalized_bom[target_attribute].fillna("").ne("")
            & normalized_bom["Comparison_Eligible"]
        ].copy()
        rows = []

        for component_type in ("R", "C"):
            component_rows = candidates[candidates["Component_Type"] == component_type]
            family_fields = self._family_fields(component_type, target_attribute)
            if component_rows.empty:
                continue

            for family_key, group in component_rows.groupby(family_fields, dropna=False, sort=True):
                attribute_values = self._unique_text(group[target_attribute])
                if len(attribute_values) < 2:
                    continue

                part_numbers = self._unique_text(group["Part_Number"])
                vendors = self._unique_text(group["Vendor"], excluded={"Unknown"})
                is_critical = bool(group["Is_Critical"].any())
                context = self._format_context(family_fields, family_key)
                confidence = float(
                    self.rule_library.get("confidence", "structured_attribute", default=0.92)
                )
                rows.append(
                    {
                        "Component_Type": component_type,
                        "Normalized_Value": group.iloc[0]["Normalized_Value"],
                        "Compared_Attribute": self.ATTRIBUTE_LABELS[target_attribute],
                        "Attribute_Values": ", ".join(attribute_values),
                        "Specification_Context": context,
                        "Part_Number_Count": len(part_numbers),
                        "Part_Numbers": ", ".join(part_numbers),
                        "Vendor_Count": len(vendors),
                        "Vendors": ", ".join(vendors),
                        "Total_Quantity": float(group["Quantity_Normalized"].sum()),
                        "References": ", ".join(self._unique_text(group["Reference"])),
                        "Has_Critical_Part": "Yes" if is_critical else "No",
                        "Confidence": confidence,
                        "Recommendation": self._variant_recommendation(
                            target_attribute,
                            is_critical,
                        ),
                    }
                )

        for index, row in enumerate(rows, start=1):
            row["Variant_ID"] = f"{identifier_prefix}-{index:04d}"

        return pd.DataFrame(rows, columns=self.VARIANT_COLUMNS)

    @staticmethod
    def _family_fields(component_type, target_attribute):
        fields = {
            "C": [
                "Component_Type",
                "Normalized_Value",
                "Voltage",
                "Dielectric",
                "Material",
                "Tolerance",
                "Package_Identity",
            ],
            "R": [
                "Component_Type",
                "Normalized_Value",
                "Voltage",
                "Material",
                "Tolerance",
                "Power_Rating",
                "Package_Identity",
            ],
        }[component_type]
        return [field for field in fields if field != target_attribute]

    @staticmethod
    def _format_context(fields, family_key):
        values = family_key if isinstance(family_key, tuple) else (family_key,)
        return "; ".join(
            f"{field}={value or '(blank)'}"
            for field, value in zip(fields, values)
            if field != "Component_Type"
        )

    @staticmethod
    def _variant_recommendation(target_attribute, is_critical):
        if is_critical:
            return "Protected circuit: retain design control and require owner approval."
        recommendations = {
            "Package_Identity": "Review footprint demand and standardize the package where layout permits.",
            "Voltage": "Validate derating and transient margin before consolidating voltage ratings.",
            "Dielectric": "Validate temperature, bias, tolerance, and signal requirements before substitution.",
        }
        return recommendations[target_attribute]

    @staticmethod
    def _preferred_part_number(group):
        valid_parts = group[group["Part_Number"].fillna("").astype(str).str.strip().ne("")].copy()
        if valid_parts.empty:
            return ""

        priced = valid_parts[pd.to_numeric(valid_parts["Unit_Price"], errors="coerce").gt(0)].copy()
        if not priced.empty:
            priced["Unit_Price_Numeric"] = pd.to_numeric(priced["Unit_Price"], errors="coerce")
            priced = priced.sort_values(
                ["Unit_Price_Numeric", "Quantity_Normalized"],
                ascending=[True, False],
            )
            return str(priced.iloc[0]["Part_Number"])

        demand = (
            valid_parts.groupby("Part_Number", as_index=False)["Quantity_Normalized"]
            .sum()
            .sort_values(["Quantity_Normalized", "Part_Number"], ascending=[False, True])
        )
        return str(demand.iloc[0]["Part_Number"])

    @staticmethod
    def _estimated_bom_savings(group):
        prices = pd.to_numeric(group["Unit_Price"], errors="coerce")
        valid_prices = prices[prices.gt(0)]
        if valid_prices.empty or valid_prices.nunique() < 2:
            return 0.0

        target_price = float(valid_prices.min())
        quantities = pd.to_numeric(group["Quantity_Normalized"], errors="coerce").fillna(0)
        current_prices = prices.fillna(target_price)
        return round(float(((current_prices - target_price) * quantities).clip(lower=0).sum()), 4)

    @staticmethod
    def _risk_category(rule_id):
        categories = {
            "DATA": "Data quality",
            "SUPPLY": "Supply continuity",
            "LIFE": "Lifecycle",
            "AVL": "AVL policy",
            "COST": "Cost and fragmentation",
        }
        return categories.get(str(rule_id).split("-", 1)[0], "Engineering review")

    @staticmethod
    def _unique_text(series, excluded=None):
        excluded_casefold = {str(value).casefold() for value in (excluded or set())}
        values = OrderedDict()
        for value in series:
            if pd.isna(value):
                continue
            text = str(value).strip()
            if text and text.casefold() not in excluded_casefold:
                values.setdefault(text.casefold(), text)
        return list(values.values())