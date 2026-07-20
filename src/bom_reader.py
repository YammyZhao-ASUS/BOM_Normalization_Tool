from pathlib import Path

import pandas as pd


class BOMReadError(ValueError):
    """Raised when an input file cannot be interpreted as a BOM table."""


class BOMReader:
    """Read Excel or CSV BOM data and identify the most likely header row."""

    SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".csv", ".txt"}
    CSV_ENCODINGS = ("utf-8-sig", "utf-8", "gb18030", "big5")

    def __init__(self, filename, column_candidates=None):
        self.filename = Path(filename).expanduser().resolve()
        self.column_candidates = column_candidates or {}
        self.selected_sheet = ""
        self.header_row = 0
        self.source_type = ""

    def load(self, sheet_name=None):
        if not self.filename.is_file():
            raise FileNotFoundError(f"BOM file not found: {self.filename}")

        suffix = self.filename.suffix.casefold()
        if suffix not in self.SUPPORTED_SUFFIXES:
            supported = ", ".join(sorted(self.SUPPORTED_SUFFIXES))
            raise BOMReadError(f"Unsupported BOM file type {suffix!r}. Supported: {supported}")

        if suffix in {".csv", ".txt"}:
            if sheet_name not in (None, ""):
                raise BOMReadError("sheet_name is only valid for Excel workbooks.")
            dataframe = self._load_delimited()
            self.source_type = "Delimited text"
            self.selected_sheet = "(not applicable)"
        else:
            dataframe = self._load_excel(sheet_name)
            self.source_type = "Excel workbook"

        dataframe = dataframe.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if dataframe.empty:
            raise BOMReadError("The selected BOM table contains no data rows.")

        dataframe.columns = self._deduplicate_columns(dataframe.columns)
        return dataframe.reset_index(drop=True)

    def metadata(self):
        return {
            "Input File": str(self.filename),
            "Input Type": self.source_type,
            "Selected Sheet": self.selected_sheet,
            "Detected Header Row": self.header_row + 1,
        }

    def _load_excel(self, requested_sheet):
        excel_file = None
        try:
            excel_file = pd.ExcelFile(self.filename)
            sheet_names = excel_file.sheet_names
            if not sheet_names:
                raise BOMReadError("The Excel workbook contains no worksheets.")

            candidates = self._resolve_requested_sheets(sheet_names, requested_sheet)
            ranked_sheets = []
            for order, sheet in enumerate(candidates):
                preview = pd.read_excel(
                    excel_file,
                    sheet_name=sheet,
                    header=None,
                    nrows=30,
                    dtype=object,
                )
                header_row, score = self._find_header_row(preview)
                nonempty_rows = int(
                    preview.iloc[header_row + 1 :].dropna(how="all").shape[0]
                )
                ranked_sheets.append((score, nonempty_rows, -order, sheet, header_row))

            selected_score, _, _, selected_sheet, header_row = max(ranked_sheets)
            self._require_header_confidence(selected_score, selected_sheet)
            self.selected_sheet = selected_sheet
            self.header_row = header_row
            return pd.read_excel(
                excel_file,
                sheet_name=selected_sheet,
                header=header_row,
                dtype=object,
            )
        except ImportError as error:
            if self.filename.suffix.casefold() == ".xls":
                raise BOMReadError(
                    "Legacy .xls input requires the optional 'xlrd' package."
                ) from error
            raise
        except ValueError as error:
            if isinstance(error, BOMReadError):
                raise
            raise BOMReadError(f"Unable to open Excel workbook: {error}") from error
        finally:
            if excel_file is not None:
                excel_file.close()

    def _load_delimited(self):
        last_error = None
        for encoding in self.CSV_ENCODINGS:
            try:
                preview = pd.read_csv(
                    self.filename,
                    header=None,
                    nrows=30,
                    encoding=encoding,
                    sep=None,
                    engine="python",
                    dtype=object,
                )
                header_row, score = self._find_header_row(preview)
                self._require_header_confidence(score, self.filename.name)
                self.header_row = header_row
                return pd.read_csv(
                    self.filename,
                    header=header_row,
                    encoding=encoding,
                    sep=None,
                    engine="python",
                    dtype=object,
                )
            except UnicodeDecodeError as error:
                last_error = error
            except pd.errors.ParserError as error:
                last_error = error

        raise BOMReadError(f"Unable to parse delimited BOM file: {last_error}")

    def _find_header_row(self, preview):
        if preview.empty:
            return 0, 0

        aliases = self._column_aliases()
        best_row = 0
        best_score = -1
        for row_index, row in preview.iterrows():
            score = 0
            matched_fields = set()
            for value in row.dropna():
                normalized = self._normalize_label(value)
                if not normalized:
                    continue
                for field, field_aliases in aliases.items():
                    if normalized in field_aliases:
                        score += 4
                        matched_fields.add(field)
                        break
                    if any(
                        len(alias) >= 4 and (alias in normalized or normalized in alias)
                        for alias in field_aliases
                    ):
                        score += 1
                        matched_fields.add(field)
                        break

            score += len(matched_fields) * 2
            if score > best_score:
                best_row = int(row_index)
                best_score = score

        return best_row, max(best_score, 0)

    @staticmethod
    def _require_header_confidence(score, source_name):
        if score < 12:
            raise BOMReadError(
                f"Could not identify a BOM header in {source_name!r}. "
                "At least two recognized BOM fields are required."
            )

    def _column_aliases(self):
        fallback = {
            "reference": ("reference", "refdes", "part reference", "插件位置"),
            "part_number": ("part number", "mpn", "item number"),
            "description": ("description", "component name", "item description"),
            "quantity": ("quantity", "qty", "comp quantity"),
        }
        source = self.column_candidates or fallback
        return {
            field: {self._normalize_label(alias) for alias in aliases}
            for field, aliases in source.items()
        }

    @staticmethod
    def _resolve_requested_sheets(sheet_names, requested_sheet):
        if requested_sheet in (None, ""):
            return list(sheet_names)

        if isinstance(requested_sheet, int) or str(requested_sheet).isdigit():
            index = int(requested_sheet)
            if index < 0 or index >= len(sheet_names):
                raise BOMReadError(
                    f"Worksheet index {index} is out of range for {len(sheet_names)} sheets."
                )
            return [sheet_names[index]]

        if requested_sheet not in sheet_names:
            available = ", ".join(sheet_names)
            raise BOMReadError(
                f"Worksheet {requested_sheet!r} was not found. Available: {available}"
            )
        return [requested_sheet]

    @staticmethod
    def _deduplicate_columns(columns):
        counts = {}
        result = []
        for index, column in enumerate(columns, start=1):
            label = str(column).strip()
            if not label or label.casefold().startswith("unnamed:"):
                label = f"Source_Column_{index}"
            counts[label] = counts.get(label, 0) + 1
            result.append(label if counts[label] == 1 else f"{label}_{counts[label]}")
        return result

    @staticmethod
    def _normalize_label(value):
        return " ".join(str(value).strip().casefold().replace("_", " ").split())
