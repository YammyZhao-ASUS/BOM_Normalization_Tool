import os
import re
import uuid
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo


class ExcelReportWriter:
    """Create an audit-friendly Excel workbook with an executive dashboard."""

    SHEETS = (
        ("Summary", "summary"),
        ("Normalized BOM", "normalized_bom"),
        ("Duplicate PN", "duplicate_pn"),
        ("Near Resistance", "near_resistance"),
        ("Near Capacitance", "near_capacitance"),
        ("Near Value", "near_value"),
        ("Different Package", "different_package"),
        ("Different Voltage", "different_voltage"),
        ("Different Material", "different_material"),
        ("AVL Candidates", "avl_candidates"),
        ("Cost Down Candidate", "cost_down"),
        ("Risk Components", "risk_components"),
        ("AI Rule Findings", "ai_rule_findings"),
        ("Review Needed", "review_needed"),
        ("Vendor Distribution", "vendor_distribution"),
        ("Statistics", "statistics"),
        ("Critical Parts", "critical_parts"),
        ("BOM Score", "bom_score"),
        ("Rule Library", "rule_library"),
        ("Report Metadata", "report_metadata"),
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

                for sheet_name, report_key in self.SHEETS:
                    dataframe = reports.get(report_key, pd.DataFrame())
                    safe_dataframe = self._sanitize_dataframe(dataframe)
                    safe_dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    self._format_data_sheet(worksheet, safe_dataframe, report_key)

                self._build_dashboard(dashboard, reports)

            os.replace(temporary_path, output_path)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise

        return str(output_path)

    def _build_dashboard(self, worksheet, reports):
        worksheet.sheet_view.showGridLines = False
        worksheet.freeze_panes = "A4"
        worksheet.sheet_view.zoomScale = 90
        worksheet.merge_cells("A1:H2")
        title = worksheet["A1"]
        title.value = "BOM Intelligence Dashboard"
        title.fill = PatternFill("solid", fgColor=self.COLORS["navy"])
        title.font = Font(name="Aptos Display", size=24, bold=True, color=self.COLORS["white"])
        title.alignment = Alignment(vertical="center")

        for row in worksheet[1:2]:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=self.COLORS["navy"])

        worksheet.merge_cells("A3:H3")
        worksheet["A3"] = "Engineering normalization, AVL readiness, cost-down, and supply-risk review"
        worksheet["A3"].font = Font(name="Aptos", size=11, color=self.COLORS["muted"])
        worksheet["A3"].alignment = Alignment(vertical="center")

        summary = self._metric_lookup(reports.get("summary", pd.DataFrame()))
        cards = (
            ("Overall Score", summary.get("Overall Score", 0), "Health / 100"),
            ("BOM Lines", summary.get("BOM Lines", 0), "Analyzed rows"),
            ("Unified AVL Ready", summary.get("Unified AVL Ready", 0), "Specification groups"),
            ("Cost Down Candidates", summary.get("Cost Down Candidates", 0), "Ranked opportunities"),
            ("High Risk Findings", summary.get("High Risk Findings", 0), "Requires action"),
            ("Average Data Quality", summary.get("Average Data Quality", 0), "Completeness / 100"),
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

        worksheet.merge_cells("G5:H12")
        worksheet["G5"] = self._management_action_text(summary)
        worksheet["G5"].fill = PatternFill("solid", fgColor=self.COLORS["cyan"])
        worksheet["G5"].font = Font(name="Aptos", size=11, color=self.COLORS["ink"])
        worksheet["G5"].alignment = Alignment(vertical="top", wrap_text=True, indent=1)
        worksheet["G5"].border = self._thin_border()

        self._write_chart_sources(worksheet, reports, summary)
        self._add_findings_chart(worksheet)
        self._add_risk_chart(worksheet)
        self._add_vendor_chart(worksheet)

        metadata = self._metric_lookup(
            reports.get("report_metadata", pd.DataFrame()),
            key_column="Property",
        )
        worksheet.merge_cells("A44:H45")
        worksheet["A44"] = (
            f"Generated: {metadata.get('Generated UTC', '')}  |  "
            f"Rules: {metadata.get('Rule Source', '')}  |  "
            f"Platform: {metadata.get('Platform Version', '')}"
        )
        worksheet["A44"].font = Font(name="Aptos", size=9, color=self.COLORS["muted"])
        worksheet["A44"].alignment = Alignment(vertical="center", wrap_text=True)

        for column in "ABCDEFGH":
            worksheet.column_dimensions[column].width = 15
        worksheet.column_dimensions["A"].width = 18
        worksheet.column_dimensions["H"].width = 18
        for column in range(10, 18):
            worksheet.column_dimensions[worksheet.cell(1, column).column_letter].hidden = True

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
        cost_count = summary.get("Cost Down Candidates", 0)
        return (
            "MANAGEMENT ACTIONS\n\n"
            f"1. Resolve {risk_count} high-risk finding(s).\n\n"
            f"2. Approve {avl_count} unified AVL candidate(s).\n\n"
            f"3. Review {cost_count} ranked cost-down opportunity(ies).\n\n"
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

        if worksheet.max_row > 1 and worksheet.max_column > 0:
            table_name = self._table_name(report_key)
            table = Table(displayName=table_name, ref=worksheet.dimensions)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            worksheet.add_table(table)

        self._add_conditional_formatting(worksheet, dataframe)
        if report_key in {"summary", "bom_score", "report_metadata"}:
            for cell in worksheet["A"]:
                cell.font = Font(name="Aptos", size=10, bold=True, color=self.COLORS["ink"])

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

    @staticmethod
    def _table_name(report_key):
        sanitized = re.sub(r"[^A-Za-z0-9_]", "_", f"tbl_{report_key}")
        return sanitized[:255]

    def _thin_border(self):
        side = Side(style="thin", color=self.COLORS["line"])
        return Border(left=side, right=side, top=side, bottom=side)