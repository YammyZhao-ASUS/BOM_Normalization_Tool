import queue
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from .basic_bom_tool import BasicBOMTool
except ImportError:
    from basic_bom_tool import BasicBOMTool


BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = BASE_DIR / "input" / "bom.xlsx"
DEFAULT_OUTPUT = BASE_DIR / "output" / "BOM_Basic_Report.xlsx"


class BOMToolApp:
    """Minimal desktop workflow for creating a basic BOM report."""

    def __init__(self, root):
        self.root = root
        self.root.title("BOM Basic Tool")
        self.root.minsize(600, 280)

        self.input_path = tk.StringVar(
            value=str(DEFAULT_INPUT) if DEFAULT_INPUT.is_file() else ""
        )
        self.output_path = tk.StringVar(value=str(DEFAULT_OUTPUT))
        self.status = tk.StringVar(value="Select a BOM file, then generate the report.")
        self.events = queue.Queue()

        self._build_layout()
        self.root.after(100, self._process_events)

    def _build_layout(self):
        container = ttk.Frame(self.root, padding=24)
        container.grid(sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        ttk.Label(
            container,
            text="BOM Basic Tool",
            font=("Segoe UI", 18, "bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(
            container,
            text="Creates a simple report with normalized values, same specifications, and near values.",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 22))

        self._add_path_row(container, 2, "BOM file", self.input_path, self._choose_input)
        self._add_path_row(container, 3, "Report file", self.output_path, self._choose_output)

        self.run_button = ttk.Button(
            container,
            text="Generate Report",
            command=self._start_analysis,
        )
        self.run_button.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(22, 10))
        self.progress = ttk.Progressbar(container, mode="indeterminate")
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew")
        ttk.Label(container, textvariable=self.status).grid(
            row=6,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(10, 0),
        )

    @staticmethod
    def _add_path_row(container, row, label, variable, command):
        ttk.Label(container, text=label).grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(container, textvariable=variable).grid(
            row=row,
            column=1,
            sticky="ew",
            padx=(14, 8),
            pady=6,
        )
        ttk.Button(container, text="Browse", command=command).grid(
            row=row,
            column=2,
            pady=6,
        )

    def _choose_input(self):
        selected = filedialog.askopenfilename(
            title="Select BOM file",
            filetypes=[
                ("BOM files", "*.xlsx *.xlsm *.xls *.csv"),
                ("All files", "*.*"),
            ],
        )
        if selected:
            self.input_path.set(selected)
            self.output_path.set(
                str(DEFAULT_OUTPUT.parent / f"{Path(selected).stem}_Basic_Report.xlsx")
            )

    def _choose_output(self):
        selected = filedialog.asksaveasfilename(
            title="Save basic BOM report",
            initialfile=Path(self.output_path.get()).name,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if selected:
            self.output_path.set(selected)

    def _start_analysis(self):
        input_file = Path(self.input_path.get()).expanduser()
        output_file = Path(self.output_path.get()).expanduser()
        if not input_file.is_file():
            messagebox.showerror("BOM Basic Tool", "Choose an existing BOM file.")
            return
        if output_file.suffix.casefold() != ".xlsx":
            messagebox.showerror("BOM Basic Tool", "The report file must use the .xlsx extension.")
            return

        self.run_button.state(["disabled"])
        self.progress.start(12)
        self.status.set("Generating basic BOM report...")
        threading.Thread(
            target=self._run_analysis,
            args=(input_file, output_file),
            daemon=True,
        ).start()

    def _run_analysis(self, input_file, output_file):
        try:
            tool = BasicBOMTool()
            reports = tool.analyze_file(input_file)
            created_file = tool.write_excel_report(reports, output_file)
            self.events.put(("success", created_file, reports["summary"]))
        except Exception as error:
            self.events.put(("error", str(error), None))

    def _process_events(self):
        try:
            event, payload, summary = self.events.get_nowait()
        except queue.Empty:
            self.root.after(100, self._process_events)
            return

        self.progress.stop()
        self.run_button.state(["!disabled"])
        if event == "success":
            self.status.set(f"Completed: {payload}")
            metrics = summary.set_index("Metric")["Value"]
            messagebox.showinfo(
                "BOM Basic Tool",
                "Report created.\n\n"
                f"BOM lines: {metrics['BOM Lines']}\n"
                f"Same specifications: {metrics['Same Specification Groups']}\n"
                f"Near values: {metrics['Near Value Pairs']}\n\n"
                f"{payload}",
            )
        else:
            self.status.set("Failed. Review the error and try again.")
            messagebox.showerror("BOM Basic Tool", payload)

        self.root.after(100, self._process_events)


def main():
    root = tk.Tk()
    BOMToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()