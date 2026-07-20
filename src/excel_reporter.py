import os
import re
import uuid
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


class ExcelReportWriter:
    """Create an RD-focused Excel workbook for daily PN optimization work."""

    SHEETS = (
        ("Merge Candidate", "merge_candidates"),
        ("Specification Summary", "specification_summary"),
        ("Specification Detail", "specification_detail"),
        ("AVL Candidate", "avl_candidate"),
        ("Risk Review", "risk_review"),
        ("Settings", "settings"),
    )

    COLORS = {
        "navy": "17324D",
        "teal": "007F7B",
        "cyan": "DDF4F2",
        "ink": "17242F",
        "muted": "607482",
        "line": "D7E0E5",
        "paper": "F5F8FA",
        "white": "FFFFFF",
        "critical": "F8D7DA",
        "high": "FCE5CD",
        "medium": "FFF2CC",
        "low": "D9EAD3",
    }

    def write(self, reports, output_file):
        output_path = Path(output_file).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = output_path.with_name(
            f".{output_path.stem}.{uuid.uuid4().hex}.tmp{output_path.suffix or '.xlsx'}"
        )

        try:
            with pd.ExcelWriter(temporary_path, engine="openpyxl") as writer:
                workbook = writer.book
                dashboard = workbook.create_sheet("Dashboard", 0)
                report_views = self._build_report_views(reports)

                for sheet_name, report_key in self.SHEETS:
                    dataframe = report_views.get(report_key, pd.DataFrame())
                    safe_dataframe = self._sanitize_dataframe(dataframe)
                    safe_dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    self._format_data_sheet(worksheet, safe_dataframe, report_key)
                    if report_key == "settings":
                        worksheet.sheet_state = "hidden"

                self._link_summary_to_detail(workbook)
                self._build_dashboard(dashboard, reports, report_views)

            os.replace(temporary_path, output_path)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise

        return str(output_path)

    def _build_report_views(self, reports):
        return {
            "merge_candidates": self._merge_candidate_view(
                reports.get("merge_candidates", pd.DataFrame())
            ),
            "specification_summary": self._specification_summary_view(reports),
            "specification_detail": self._specification_detail_view(reports),
            "avl_candidate": self._avl_candidate_view(
                reports.get("avl_candidates", pd.DataFrame())
            ),
            "risk_review": self._risk_review_view(reports),
            "settings": self._settings_view(reports),
        }

    def _merge_candidate_view(self, merge_candidates):
        columns = [
            "Priority",
            "Merge Difficulty",
            "Difference",
            "Spec",
            "Current PN",
            "Current Qty",
            "Target PN",
            "Target Qty",
            "Reference Count",
            "Reason",
            "Action",
            "Score",
        ]
        if merge_candidates.empty:
            return pd.DataFrame(columns=columns)

        rows = []
        for _, candidate in merge_candidates.iterrows():
            difference = self._difference_label(candidate.get("Spec_Similarity", 0))
            rows.append(
                {
                    "Priority": candidate.get("Priority_Stars", "") or candidate.get("Priority", ""),
                    "Merge Difficulty": self._difficulty_label(candidate.get("Spec_Similarity", 0)),
                    "Difference": difference,
                    "Spec": self._compact_spec(candidate),
                    "Current PN": candidate.get("Current_PN", ""),
                    "Current Qty": candidate.get("Current_Qty", 0),
                    "Target PN": candidate.get("Target_PN", ""),
                    "Target Qty": candidate.get("Target_Qty", 0),
                    "Reference Count": self._reference_count(candidate.get("Item", "")),
                    "Reason": candidate.get("Reason", ""),
                    "Action": candidate.get("Recommendation", ""),
                    "Score": candidate.get("Merge_Score", 0),
                }
            )

        return pd.DataFrame(rows, columns=columns)

    def _specification_summary_view(self, reports):
        columns = ["Value", "Spec Detail", "PN Count", "Total Qty", "Target PN", "Priority", "Reason", "Detail", "Group"]
        rows = []
        for group in self._merge_tree_groups(reports):
            rows.append(
                {
                    "Value": group["value"],
                    "Spec Detail": group["spec_detail"],
                    "PN Count": group["pn_count"],
                    "Total Qty": group["total_quantity"],
                    "Target PN": group["target"]["part_number"],
                    "Priority": group["potential"],
                    "Reason": group["summary_reason"],
                    "Detail": "▶ Detail",
                    "Group": group["group_id"],
                }
            )

        for group in self._variant_summary_groups(reports, len(rows) + 1):
            rows.append(group)

        dataframe = pd.DataFrame(rows, columns=columns)
        if dataframe.empty:
            return dataframe
        return dataframe.sort_values(["Value", "Priority", "Total Qty"], ascending=[True, False, False]).reset_index(drop=True)

    def _specification_detail_view(self, reports):
        columns = ["Group", "Row Type", "Merge Tree", "PN", "Qty", "Qty Share", "Difference", "Can Merge", "Reason"]
        rows = []
        for group in self._merge_tree_groups(reports):
            target_share, candidate_share = self._target_candidate_share(group)
            rows.append(
                {
                    "Group": group["group_id"],
                    "Row Type": "Group Header",
                    "Merge Tree": f"{group['spec']}  |  {group['pn_count']} PN  |  Target {target_share}  |  Candidate {candidate_share}",
                    "PN": "",
                    "Qty": "",
                    "Qty Share": "",
                    "Difference": "",
                    "Can Merge": group["potential"],
                    "Reason": group["summary_reason"],
                }
            )
            target = group["target"]
            rows.append(
                {
                    "Group": group["group_id"],
                    "Row Type": "Target PN",
                    "Merge Tree": f"▲ Target PN  {target['part_number']}  {target['quantity']:g} pcs",
                    "PN": target["part_number"],
                    "Qty": target["quantity"],
                    "Qty Share": target_share,
                    "Difference": "",
                    "Can Merge": "Target",
                    "Reason": "Main target selected by highest quantity",
                }
            )
            for index, candidate in enumerate(group["candidates"]):
                branch = "└──" if index == len(group["candidates"]) - 1 else "├──"
                difference, reason = self._part_difference(candidate["representative"], target["representative"])
                rows.append(
                    {
                        "Group": group["group_id"],
                        "Row Type": "Candidate",
                        "Merge Tree": f"{branch} {candidate['part_number']}  {candidate['quantity']:g} pcs",
                        "PN": candidate["part_number"],
                        "Qty": candidate["quantity"],
                        "Qty Share": self._quantity_share(candidate["quantity"], group["total_quantity"]),
                        "Difference": difference,
                        "Can Merge": group["potential"],
                        "Reason": reason,
                    }
                )
            rows.append({column: "" for column in columns})

        rows.extend(self._variant_detail_groups(reports))

        return pd.DataFrame(rows, columns=columns)

    @staticmethod
    def _avl_candidate_view(avl_candidates):
        columns = ["Spec", "Preferred PN", "PN Count", "Vendor", "Qty", "Status", "Suggestion"]
        if avl_candidates.empty:
            return pd.DataFrame(columns=columns)
        dataframe = avl_candidates.rename(
            columns={
                "Normalize_Key": "Spec",
                "Preferred_PN": "Preferred PN",
                "Part_Number_Count": "PN Count",
                "Vendors": "Vendor",
                "Total_Quantity": "Qty",
                "AVL_Readiness": "Status",
                "Recommendation": "Suggestion",
            }
        )
        return dataframe.reindex(columns=columns)

    def _risk_review_view(self, reports):
        columns = ["Severity", "Risk", "PN", "Reference Count", "Finding", "Suggestion"]
        risk_components = reports.get("risk_components", pd.DataFrame())
        rows = []
        for _, risk in risk_components.iterrows():
            rows.append(
                {
                    "Severity": risk.get("Severity", ""),
                    "Risk": risk.get("Risk_Category", ""),
                    "PN": risk.get("Part_Number", ""),
                    "Reference Count": self._reference_count(risk.get("Reference", "")),
                    "Finding": risk.get("Finding", ""),
                    "Suggestion": risk.get("Recommendation", ""),
                }
            )
        return pd.DataFrame(rows, columns=columns)

    def _merge_tree_groups(self, reports):
        normalized = reports.get("normalized_bom", pd.DataFrame())
        required_columns = {"Component_Type", "Normalize_Key", "Part_Number", "Quantity_Normalized"}
        if normalized.empty or not required_columns.issubset(normalized.columns):
            return []

        candidates = normalized[
            normalized["Component_Type"].isin(["C", "R"])
            & normalized["Normalize_Key"].fillna("").ne("")
            & normalized["Part_Number"].fillna("").astype(str).str.strip().ne("")
            & ~normalized.get("Is_Second_Source", pd.Series(False, index=normalized.index))
        ]
        groups = []
        for _, group in candidates.groupby(["Component_Type", "Normalize_Key"], sort=True):
            part_groups = self._part_groups(group)
            if len(part_groups) < 2:
                continue

            target = max(
                part_groups,
                key=lambda item: (item["quantity"], item["line_count"], item["part_number"]),
            )
            candidate_parts = [part for part in part_groups if part["part_number"] != target["part_number"]]
            candidate_parts.sort(key=lambda item: (item["quantity"], item["part_number"]))
            potential = self._group_potential(candidate_parts)
            value, spec_detail = self._split_spec(group.iloc[0])
            groups.append(
                {
                    "value": value,
                    "spec_detail": spec_detail,
                    "spec": self._compact_spec(group.iloc[0]),
                    "pn_count": len(part_groups),
                    "total_quantity": float(group["Quantity_Normalized"].sum()),
                    "target": target,
                    "candidates": candidate_parts,
                    "potential": potential,
                    "summary_reason": self._group_summary_reason(candidate_parts, target),
                }
            )

        groups.sort(key=lambda item: (item["potential"], item["total_quantity"]), reverse=True)
        for index, group in enumerate(groups, start=1):
            group["group_id"] = f"Group{index:03d}"
        return groups

    def _part_groups(self, group):
        part_groups = []
        for part_number, part_group in group.groupby("Part_Number", sort=False):
            representative = part_group.iloc[0]
            part_groups.append(
                {
                    "part_number": str(part_number),
                    "quantity": float(part_group["Quantity_Normalized"].sum()),
                    "line_count": len(part_group),
                    "vendor": ", ".join(self._unique_text(part_group.get("Vendor", pd.Series(dtype=object)), case_insensitive=True)),
                    "representative": representative,
                }
            )
        return part_groups

    @staticmethod
    def _group_potential(candidate_parts):
        if not candidate_parts:
            return ""
        minimum_quantity = min(part["quantity"] for part in candidate_parts)
        if minimum_quantity <= 2:
            return "★★★★★"
        if minimum_quantity <= 5:
            return "★★★★"
        if minimum_quantity <= 10:
            return "★★★"
        return "★★"

    def _group_summary_reason(self, candidate_parts, target):
        if not candidate_parts:
            return "No merge candidate"
        differences = []
        for candidate in candidate_parts:
            difference, _ = self._part_difference(candidate["representative"], target["representative"])
            if difference and difference not in differences:
                differences.append(difference)
        if differences == ["🟢 Same spec"]:
            return "Same Spec"
        if differences == ["🟢 Vendor"]:
            return "Vendor Only"
        return " / ".join(differences) or "Same Spec"

    def _variant_summary_groups(self, reports, start_index):
        rows = []
        for report_key, status, label in (
            ("different_package", "★★★", "🟡 Package"),
            ("different_voltage", "★★", "🟠 Voltage"),
            ("different_material", "★", "🔴 Material"),
        ):
            variants = reports.get(report_key, pd.DataFrame())
            for _, variant in variants.iterrows():
                values = str(variant.get("Attribute_Values", ""))
                rows.append(
                    {
                        "Value": variant.get("Normalized_Value", ""),
                        "Spec Detail": self._compact_attribute_values(values),
                        "PN Count": variant.get("Part_Number_Count", 0),
                        "Total Qty": variant.get("Total_Quantity", 0),
                        "Target PN": "Review required",
                        "Priority": status,
                        "Reason": f"{label} {self._compact_attribute_values(values)}",
                        "Detail": "▶ Detail",
                        "Group": f"Review{len(rows) + 1:03d}",
                    }
                )
        return rows

    def _variant_detail_groups(self, reports):
        rows = []
        columns = ["Group", "Row Type", "Merge Tree", "PN", "Qty", "Qty Share", "Difference", "Can Merge", "Reason"]
        group_index = 1
        for report_key, label in (
            ("different_package", "🟡 Package"),
            ("different_voltage", "🟠 Voltage"),
            ("different_material", "🔴 Material"),
        ):
            variants = reports.get(report_key, pd.DataFrame())
            for _, variant in variants.iterrows():
                group_id = f"Review{group_index:03d}"
                rows.append(
                    {
                        "Group": group_id,
                        "Row Type": "Cannot Merge",
                        "Merge Tree": f"{group_id}  {variant.get('Normalized_Value', '')}  |  {variant.get('Part_Number_Count', 0)} PN",
                        "PN": "",
                        "Qty": variant.get("Total_Quantity", 0),
                        "Qty Share": "",
                        "Difference": label,
                        "Can Merge": "No",
                        "Reason": f"{label} {self._compact_attribute_values(variant.get('Attribute_Values', ''))}",
                    }
                )
                rows.append({column: "" for column in columns})
                group_index += 1
        return rows

    def _part_difference(self, current, target):
        checks = (
            ("Normalized_Value", "🔴 Value", "Value"),
            ("Dielectric", "🔴 Material", "Material"),
            ("Package_Identity", "🟡 Package", "Package"),
            ("Voltage", "🟠 Voltage", "Voltage"),
            ("Tolerance", "🟡 Tolerance", "Tolerance"),
        )
        for field, label, name in checks:
            current_value = str(current.get(field, "") or "")
            target_value = str(target.get(field, "") or "")
            if current_value != target_value:
                return label, f"{name} {current_value or '(blank)'} -> {target_value or '(blank)'}"

        current_vendor = str(current.get("Vendor", "") or "").casefold()
        target_vendor = str(target.get("Vendor", "") or "").casefold()
        if current_vendor != target_vendor:
            return "🟢 Vendor", "Vendor different"
        return "🟢 Same spec", "Same specification"

    def _settings_view(self, reports):
        rows = []
        for section, dataframe in (
            ("Summary", reports.get("summary", pd.DataFrame())),
            ("Metadata", reports.get("report_metadata", pd.DataFrame())),
            ("Rules", reports.get("rule_library", pd.DataFrame())),
        ):
            if dataframe.empty:
                continue
            for _, row in dataframe.iterrows():
                rows.append(
                    {
                        "Section": section,
                        "Name": row.get("Metric", row.get("Property", row.get("Rule_Path", ""))),
                        "Value": row.get("Value", ""),
                        "Source": row.get("Source", ""),
                    }
                )
        return pd.DataFrame(rows, columns=["Section", "Name", "Value", "Source"])

    def _build_dashboard(self, worksheet, reports, report_views):
        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A17"
        worksheet.sheet_view.zoomScale = 90
        worksheet.merge_cells("A1:H2")
        title = worksheet["A1"]
        title.value = "PN Optimization Dashboard"
        title.fill = PatternFill("solid", fgColor=self.COLORS["navy"])
        title.font = Font(name="Aptos Display", size=24, bold=True, color=self.COLORS["white"])
        title.alignment = Alignment(vertical="center")

        for row in worksheet[1:2]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=self.COLORS["navy"])

        worksheet.merge_cells("A3:H3")
        worksheet["A3"] = "Daily RD merge actions: current PN -> target PN"
        worksheet["A3"].font = Font(name="Aptos", size=11, color=self.COLORS["muted"])
        worksheet["A3"].alignment = Alignment(vertical="center")

        summary = self._metric_lookup(reports.get("summary", pd.DataFrame()))
        normalized = reports.get("normalized_bom", pd.DataFrame())
        merge_view = report_views.get("merge_candidates", pd.DataFrame())
        merge_raw = reports.get("merge_candidates", pd.DataFrame())
        project = self._project_name(reports)
        total_pn = int(normalized["Part_Number"].replace("", pd.NA).nunique()) if "Part_Number" in normalized else 0
        merge_pn_reduction = int(merge_raw["Current_PN"].replace("", pd.NA).nunique()) if "Current_PN" in merge_raw else 0
        merge_rate = f"{round(merge_pn_reduction / total_pn * 100):g}%" if total_pn else "0%"
        cards = (
            ("Project", project, "Source BOM"),
            ("Total BOM", summary.get("BOM Lines", 0), "Analyzed rows"),
            ("Candidate", len(merge_view), "PN merge actions"),
            ("Priority 1", int((merge_raw.get("Priority", pd.Series(dtype=object)) == "Priority 1").sum()), "Do first"),
            ("Estimated PN Reduction", merge_pn_reduction, "Current PN removed"),
            ("Estimated Merge Rate", merge_rate, "PN reduction / total PN"),
        )
        card_ranges = (
            "A5:B8",
            "C5:D8",
            "E5:F8",
            "A9:B12",
            "C9:D12",
            "E9:F12",
        )
        for card, cell_range in zip(cards, card_ranges):
            self._draw_kpi(worksheet, cell_range, *card)

        self._write_top_merge_table(worksheet, merge_view)

        metadata = self._metric_lookup(
            reports.get("report_metadata", pd.DataFrame()),
            key_column="Property",
        )
        worksheet.merge_cells("A44:H44")
        worksheet["A44"] = (
            f"Generated: {metadata.get('Generated UTC', '')}  |  "
            f"Rules: {metadata.get('Rule Source', '')}  |  "
            f"Platform: {metadata.get('Platform Version', '')}"
        )
        worksheet["A44"].font = Font(name="Aptos", size=9, color=self.COLORS["muted"])
        worksheet["A44"].alignment = Alignment(vertical="center", wrap_text=True)

        for column in "ABCDEFGH":
            worksheet.column_dimensions[column].width = 18
        worksheet.column_dimensions["D"].width = 24
        worksheet.column_dimensions["E"].width = 24
        worksheet.column_dimensions["F"].width = 12

    def _write_top_merge_table(self, worksheet, merge_view):
        worksheet.merge_cells("A15:H15")
        worksheet["A15"] = "★★★★★ Priority 1"
        worksheet["A15"].font = Font(name="Aptos Display", size=16, bold=True, color=self.COLORS["teal"])
        worksheet["A15"].alignment = Alignment(vertical="center")

        headers = ["Spec", "Current PN", "Qty", "Target PN", "Target Qty", "Difference", "Difficulty", "Action"]
        for column_index, header in enumerate(headers, start=1):
            cell = worksheet.cell(16, column_index, header)
            cell.fill = PatternFill("solid", fgColor=self.COLORS["navy"])
            cell.font = Font(name="Aptos", size=10, bold=True, color=self.COLORS["white"])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        if merge_view.empty:
            worksheet.cell(17, 1, "No Priority 1 merge candidate found.")
            return

        top_rows = merge_view[merge_view["Priority"].eq("★★★★★")].head(15)
        if top_rows.empty:
            top_rows = merge_view.head(15)

        for row_index, (_, row) in enumerate(top_rows.iterrows(), start=17):
            values = [
                row.get("Spec", ""),
                row.get("Current PN", ""),
                row.get("Current Qty", ""),
                row.get("Target PN", ""),
                row.get("Target Qty", ""),
                row.get("Difference", ""),
                row.get("Merge Difficulty", ""),
                row.get("Action", ""),
            ]
            for column_index, value in enumerate(values, start=1):
                cell = worksheet.cell(row_index, column_index, self._safe_excel_value(value))
                cell.font = Font(name="Aptos", size=10, color=self.COLORS["ink"])
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = Border(bottom=Side(style="hair", color=self.COLORS["line"]))
                if row_index % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor=self.COLORS["paper"])

    def _draw_kpi(self, worksheet, cell_range, label, value, caption):
        start_cell, end_cell = cell_range.split(":")
        start = worksheet[start_cell]
        end = worksheet[end_cell]
        min_row, max_row = start.row, end.row
        min_column, max_column = start.column, end.column

        for row in worksheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_column,
            max_col=max_column,
        ):
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=self.COLORS["white"])
                cell.border = self._thin_border()

        worksheet.merge_cells(
            start_row=min_row,
            start_column=min_column,
            end_row=min_row,
            end_column=max_column,
        )
        worksheet.merge_cells(
            start_row=min_row + 1,
            start_column=min_column,
            end_row=min_row + 2,
            end_column=max_column,
        )
        worksheet.merge_cells(
            start_row=max_row,
            start_column=min_column,
            end_row=max_row,
            end_column=max_column,
        )

        label_cell = worksheet.cell(min_row, min_column)
        label_cell.value = label
        label_cell.font = Font(name="Aptos", size=10, bold=True, color=self.COLORS["muted"])
        label_cell.alignment = Alignment(vertical="center")

        value_cell = worksheet.cell(min_row + 1, min_column)
        value_cell.value = value
        value_cell.font = Font(name="Aptos Display", size=22, bold=True, color=self.COLORS["teal"])
        value_cell.alignment = Alignment(vertical="center")

        caption_cell = worksheet.cell(max_row, min_column)
        caption_cell.value = caption
        caption_cell.font = Font(name="Aptos", size=9, color=self.COLORS["muted"])
        caption_cell.alignment = Alignment(vertical="center")

    def _write_chart_sources(self, worksheet, reports, summary):
        findings = [
            ("Duplicate PN", summary.get("Duplicate PN Groups", 0)),
            ("Near values", summary.get("Near Value Pairs", 0)),
            ("Package variants", summary.get("Different Package Groups", 0)),
            ("Voltage variants", summary.get("Different Voltage Groups", 0)),
            ("Material variants", summary.get("Different Material Groups", 0)),
        ]
        worksheet["J1"] = "Finding"
        worksheet["K1"] = "Count"
        for row_index, (label, count) in enumerate(findings, start=2):
            worksheet.cell(row_index, 10, label)
            worksheet.cell(row_index, 11, count)

        worksheet["M1"] = "Severity"
        worksheet["N1"] = "Count"
        rule_findings = reports.get("ai_rule_findings", pd.DataFrame())
        severity_counts = (
            rule_findings["Severity"].value_counts().to_dict()
            if "Severity" in rule_findings
            else {}
        )
        for row_index, severity in enumerate(("Critical", "High", "Medium", "Low"), start=2):
            worksheet.cell(row_index, 13, severity)
            worksheet.cell(row_index, 14, int(severity_counts.get(severity, 0)))

        worksheet["P1"] = "Vendor"
        worksheet["Q1"] = "Quantity"
        normalized = reports.get("normalized_bom", pd.DataFrame())
        if {"Vendor", "Quantity_Normalized"}.issubset(normalized.columns):
            vendor_quantities = (
                normalized[normalized["Vendor"].ne("Unknown")]
                .groupby("Vendor")["Quantity_Normalized"]
                .sum()
                .sort_values(ascending=False)
                .head(8)
            )
            for row_index, (vendor, quantity) in enumerate(vendor_quantities.items(), start=2):
                worksheet.cell(row_index, 16, self._safe_excel_value(vendor))
                worksheet.cell(row_index, 17, float(quantity))

    @staticmethod
    def _management_action_text(summary):
        risk_count = summary.get("High Risk Findings", 0)
        avl_count = summary.get("Unified AVL Ready", 0)
        merge_count = summary.get("Top Merge Candidates", 0)
        return (
            "MANAGEMENT ACTIONS\n\n"
            f"1. Resolve {risk_count} high-risk finding(s).\n\n"
            f"2. Approve {avl_count} unified AVL candidate(s).\n\n"
            f"3. Review {merge_count} top merge candidate(s).\n\n"
            "Engineering validation remains mandatory for protected circuits."
        )

    def _add_findings_chart(self, worksheet):
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Normalization Findings"
        chart.y_axis.title = "Groups / pairs"
        chart.height = 7.2
        chart.width = 13.2
        chart.add_data(Reference(worksheet, min_col=11, min_row=1, max_row=6), titles_from_data=True)
        chart.set_categories(Reference(worksheet, min_col=10, min_row=2, max_row=6))
        chart.legend = None
        worksheet.add_chart(chart, "A14")

    def _add_risk_chart(self, worksheet):
        chart = PieChart()
        chart.style = 10
        chart.title = "Rule Findings by Severity"
        chart.height = 7.2
        chart.width = 10.2
        chart.add_data(Reference(worksheet, min_col=14, min_row=1, max_row=5), titles_from_data=True)
        chart.set_categories(Reference(worksheet, min_col=13, min_row=2, max_row=5))
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showPercent = True
        chart.dataLabels.showLeaderLines = True
        worksheet.add_chart(chart, "E14")

    def _add_vendor_chart(self, worksheet):
        populated_rows = max(
            1,
            sum(1 for row in range(2, 10) if worksheet.cell(row, 16).value is not None),
        )
        chart = BarChart()
        chart.type = "bar"
        chart.style = 11
        chart.title = "Top Vendor Quantity Exposure"
        chart.x_axis.title = "Quantity"
        chart.height = 7.8
        chart.width = 17.5
        chart.add_data(
            Reference(worksheet, min_col=17, min_row=1, max_row=populated_rows + 1),
            titles_from_data=True,
        )
        chart.set_categories(
            Reference(worksheet, min_col=16, min_row=2, max_row=populated_rows + 1)
        )
        chart.legend = None
        worksheet.add_chart(chart, "A29")

    @staticmethod
    def _compact_spec(row):
        parts = [
            row.get("Normalized_Value", ""),
            row.get("Voltage", ""),
            row.get("Dielectric", ""),
            row.get("Tolerance", ""),
            row.get("Size", ""),
        ]
        text = " ".join(str(part).strip() for part in parts if str(part).strip())
        return text or row.get("Normalize_Key", "")

    @staticmethod
    def _split_spec(row):
        value = str(row.get("Normalized_Value", "") or "").strip()
        details = [
            row.get("Voltage", ""),
            row.get("Dielectric", ""),
            row.get("Tolerance", ""),
            row.get("Size", ""),
        ]
        spec_detail = " ".join(str(part).strip() for part in details if str(part).strip())
        return value or str(row.get("Normalize_Key", "") or ""), spec_detail

    @staticmethod
    def _compact_attribute_values(attribute_values):
        values = [value.strip() for value in str(attribute_values or "").split(",") if value.strip()]
        return "⇄".join(values) if values else ""

    @staticmethod
    def _quantity_share(quantity, total_quantity):
        if not total_quantity:
            return "0 (0%)"
        percent = round(float(quantity) / float(total_quantity) * 100)
        return f"{float(quantity):g} ({percent:g}%)"

    def _target_candidate_share(self, group):
        target_quantity = group["target"]["quantity"]
        candidate_quantity = sum(candidate["quantity"] for candidate in group["candidates"])
        total_quantity = group["total_quantity"]
        return (
            self._quantity_share(target_quantity, total_quantity),
            self._quantity_share(candidate_quantity, total_quantity),
        )

    @staticmethod
    def _difficulty_label(spec_similarity):
        try:
            score = float(spec_similarity)
        except (TypeError, ValueError):
            score = 0
        if score >= 95:
            return "★★★★★ Easy"
        if score >= 70:
            return "★★★★ Medium"
        if score >= 40:
            return "★★★ Hard"
        return "★ High risk"

    @staticmethod
    def _difference_label(spec_similarity):
        try:
            score = float(spec_similarity)
        except (TypeError, ValueError):
            score = 0
        if score == 100:
            return "🟢 Same spec"
        if score == 95:
            return "🟢 Vendor"
        if score == 70:
            return "🟠 Voltage"
        if score == 60:
            return "🟡 Tolerance"
        if score == 40:
            return "🟡 Package"
        return "🔴 Material / Value"

    @staticmethod
    def _reference_count(reference_text):
        if not isinstance(reference_text, str) or not reference_text.strip():
            return 0
        return len([reference for reference in re.split(r"[,;\s]+", reference_text.strip()) if reference])

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

    @staticmethod
    def _project_name(reports):
        metadata = ExcelReportWriter._metric_lookup(
            reports.get("report_metadata", pd.DataFrame()),
            key_column="Property",
        )
        project_name = str(metadata.get("Project Name", "")).strip()
        if project_name:
            return project_name
        input_file = str(metadata.get("Input File", "")).strip()
        if input_file:
            return Path(input_file).stem
        return "BOM Project"

    def _link_summary_to_detail(self, workbook):
        if "Specification Summary" not in workbook or "Specification Detail" not in workbook:
            return

        summary = workbook["Specification Summary"]
        detail = workbook["Specification Detail"]
        detail_rows = {
            detail.cell(row_index, 1).value: row_index
            for row_index in range(2, detail.max_row + 1)
            if detail.cell(row_index, 2).value in {"Group Header", "Cannot Merge"}
        }
        headers = [cell.value for cell in summary[1]]
        try:
            detail_column = headers.index("Detail") + 1
            group_column = headers.index("Group") + 1
        except ValueError:
            return

        for row_index in range(2, summary.max_row + 1):
            group_id = summary.cell(row_index, group_column).value
            target_row = detail_rows.get(group_id, 1)
            cell = summary.cell(row_index, detail_column)
            cell.hyperlink = f"#'Specification Detail'!A{target_row}"
            cell.style = "Hyperlink"

    def _format_data_sheet(self, worksheet, dataframe, report_key):
        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 30

        header_fill = PatternFill("solid", fgColor=self.COLORS["navy"])
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = Font(name="Aptos", size=10, bold=True, color=self.COLORS["white"])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for column_index, column_name in enumerate(dataframe.columns, start=1):
            values = [str(column_name)]
            values.extend(str(value) for value in dataframe[column_name].head(250).fillna(""))
            width = min(max(max(len(value) for value in values) + 2, 11), 48)
            worksheet.column_dimensions[worksheet.cell(1, column_index).column_letter].width = width

        for row_index, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
            if row_index % 2 == 0:
                for cell in row:
                    cell.fill = PatternFill("solid", fgColor=self.COLORS["paper"])
            for cell in row:
                cell.font = Font(name="Aptos", size=10, color=self.COLORS["ink"])
                cell.alignment = Alignment(vertical="top", wrap_text=True)
                cell.border = Border(bottom=Side(style="hair", color=self.COLORS["line"]))

        self._add_conditional_formatting(worksheet, dataframe)
        if report_key == "specification_summary":
            self._format_specification_summary(worksheet, dataframe)
        if report_key == "specification_detail":
            self._format_specification_detail(worksheet, dataframe)
        if report_key in {"summary", "bom_score", "report_metadata"}:
            for cell in worksheet["A"]:
                cell.font = Font(name="Aptos", size=10, bold=True, color=self.COLORS["ink"])

    def _format_specification_summary(self, worksheet, dataframe):
        if dataframe.empty:
            return
        headers = [cell.value for cell in worksheet[1]]
        if "Group" in headers:
            group_letter = worksheet.cell(1, headers.index("Group") + 1).column_letter
            worksheet.column_dimensions[group_letter].hidden = True
        if "Value" in headers:
            value_letter = worksheet.cell(1, headers.index("Value") + 1).column_letter
            worksheet.column_dimensions[value_letter].width = 16
            for cell in worksheet[value_letter][1:]:
                cell.font = Font(name="Aptos Display", size=12, bold=True, color=self.COLORS["teal"])
        if "Reason" in headers:
            reason_letter = worksheet.cell(1, headers.index("Reason") + 1).column_letter
            worksheet.column_dimensions[reason_letter].width = 22

    def _format_specification_detail(self, worksheet, dataframe):
        if dataframe.empty:
            return
        headers = [cell.value for cell in worksheet[1]]
        if "Group" in headers:
            group_letter = worksheet.cell(1, headers.index("Group") + 1).column_letter
            worksheet.column_dimensions[group_letter].hidden = True
        try:
            row_type_column = headers.index("Row Type") + 1
            merge_tree_column = headers.index("Merge Tree") + 1
        except ValueError:
            return

        fills = {
            "Group Header": PatternFill("solid", fgColor=self.COLORS["cyan"]),
            "Target PN": PatternFill("solid", fgColor=self.COLORS["low"]),
            "Cannot Merge": PatternFill("solid", fgColor="E7E9EC"),
        }
        for row_index in range(2, worksheet.max_row + 1):
            row_type = worksheet.cell(row_index, row_type_column).value
            fill = fills.get(row_type)
            if fill:
                for cell in worksheet[row_index]:
                    cell.fill = fill
            if row_type in {"Group Header", "Target PN"}:
                worksheet.cell(row_index, merge_tree_column).font = Font(
                    name="Aptos",
                    size=10,
                    bold=True,
                    color=self.COLORS["ink"],
                )

    def _add_conditional_formatting(self, worksheet, dataframe):
        if dataframe.empty:
            return

        column_lookup = {
            str(cell.value): cell.column_letter
            for cell in worksheet[1]
            if cell.value is not None
        }
        last_row = worksheet.max_row

        severity_column = column_lookup.get("Severity")
        if severity_column:
            fills = {
                "Critical": self.COLORS["critical"],
                "High": self.COLORS["high"],
                "Medium": self.COLORS["medium"],
                "Low": self.COLORS["low"],
            }
            for severity, color in fills.items():
                worksheet.conditional_formatting.add(
                    f"{severity_column}2:{severity_column}{last_row}",
                    FormulaRule(
                        formula=[f'${severity_column}2="{severity}"'],
                        fill=PatternFill("solid", fgColor=color),
                    ),
                )

        for column_name in ("Opportunity_Score", "Data_Quality_Score", "Confidence"):
            column_letter = column_lookup.get(column_name)
            if column_letter:
                worksheet.conditional_formatting.add(
                    f"{column_letter}2:{column_letter}{last_row}",
                    ColorScaleRule(
                        start_type="min",
                        start_color="F8696B",
                        mid_type="percentile",
                        mid_value=50,
                        mid_color="FFEB84",
                        end_type="max",
                        end_color="63BE7B",
                    ),
                )

        critical_column = column_lookup.get("Has_Critical_Part")
        if critical_column:
            worksheet.conditional_formatting.add(
                f"{critical_column}2:{critical_column}{last_row}",
                FormulaRule(
                    formula=[f'${critical_column}2="Yes"'],
                    fill=PatternFill("solid", fgColor=self.COLORS["critical"]),
                ),
            )

    def _sanitize_dataframe(self, dataframe):
        safe_dataframe = dataframe.copy()
        safe_dataframe.columns = self._safe_column_labels(safe_dataframe.columns)
        for column in safe_dataframe.columns:
            safe_dataframe[column] = safe_dataframe[column].map(self._safe_excel_value)
        return safe_dataframe

    def _safe_column_labels(self, columns):
        counts = {}
        labels = []
        for index, column in enumerate(columns, start=1):
            label = str(self._safe_excel_value(str(column))).strip()
            if not label:
                label = f"Column_{index}"
            counts[label] = counts.get(label, 0) + 1
            labels.append(label if counts[label] == 1 else f"{label}_{counts[label]}")
        return labels

    @staticmethod
    def _safe_excel_value(value):
        if not isinstance(value, str):
            return value
        formula_candidate = value.lstrip("\t\r\n\ufeff")
        if formula_candidate.startswith(("=", "+", "-", "@")):
            return f"'{value}"
        return value

    @staticmethod
    def _metric_lookup(dataframe, key_column="Metric"):
        if dataframe.empty or key_column not in dataframe or "Value" not in dataframe:
            return {}
        return dict(zip(dataframe[key_column], dataframe["Value"]))

    def _thin_border(self):
        side = Side(style="thin", color=self.COLORS["line"])
        return Border(left=side, right=side, top=side, bottom=side)