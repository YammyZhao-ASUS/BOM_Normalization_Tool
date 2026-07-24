# BOM Normalization Tool

An intelligent BOM analysis tool designed for hardware engineers to identify duplicate components, normalize resistor/capacitor specifications, and assist BOM consolidation through an interactive review workflow.

The tool analyzes BOM Excel/CSV files, normalizes component values, groups equivalent parts, identifies near-value candidates, and generates an engineering review report.

---

# Features

Current supported features include:

- ✅ Resistor and Capacitor normalization
- ✅ Multiple unit conversion (Ω, KΩ, MΩ, pF, nF, uF...)
- ✅ Same Specification detection
- ✅ Near Value analysis
- ✅ Merge candidate recommendation
- ✅ Interactive Merge Workspace
- ✅ Engineering review workflow
- ✅ Excel report generation

Generated report includes:

| Worksheet | Description |
|-----------|-------------|
| Summary | Overall BOM statistics and analysis summary |
| Normalized BOM | Normalized resistor/capacitor values |
| Same Specification | Components with identical specifications but different part numbers |
| Near Values | Similar resistor/capacitor values requiring engineering review |
| Capacitor Summary | Capacitor merge recommendation summary |
| Capacitor Merge Workspace | Interactive capacitor comparison and review workspace |
| Resistor Summary | Resistor merge recommendation summary |
| Resistor Merge Workspace | Interactive resistor comparison and review workspace |

---

# Workflow
```
Import BOM
      │
      ▼
Normalize Components
      │
      ▼
Analyze Specifications
      │
      ▼
Generate Merge Candidates
      │
      ▼
Engineering Review
      │
      ▼
Export Excel Report
```

The tool **does not automatically modify the BOM**.

All merge recommendations require engineering review before implementation.

---

# Installation

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

---

# Usage

## GUI

Run:

```powershell
python src/gui.py
```

or simply double-click:

```
Run_BOM_Tool.bat
```

Select:

- Input BOM file
- Output folder

Then click:

```
Generate Report
```

---

## Command Line

```powershell
python src/main.py input/bom.xlsx --output output/BOM_Report.xlsx
```

---

# Supported BOM Columns

The tool automatically recognizes common BOM column names.

| Purpose | Example Columns |
|---------|-----------------|
| Part Number | Part Number, MPN, Item Number |
| Reference | Part Reference, Reference, RefDes |
| Description | Component_Name, Description, Item Description |
| Quantity | Quantity, Qty |
| Manufacturer | Vendor, Manufacturer |

---

# Recommended Description Format

Best recognition accuracy is achieved using descriptions such as:

```text
MLCC 100NF 16V 0402 X7R 10%

RES 10K OHM 1/16W 0402 1% THICK FILM
```

---

# Engineering Review

This tool is intended to assist engineering decisions rather than replace them.

Current recommendations include:

- Same Specification
- Near Value
- Merge Candidate
- BOM Action (Merge / Review / Keep)

Final BOM changes should always be confirmed by hardware engineers.

---

# Project Structure

```
BOM_Normalization_Tool/

├── docs/
│   ├── PRD/
│   ├── Design/
│   └── Architecture/
│
├── src/
├── input/
├── output/
├── requirements.txt
├── README.md
└── Run_BOM_Tool.bat
```

---

# Documentation

Project documentation is located under the **docs/** directory.

Suggested structure:

```
docs/

├── PRD/
├── Design/
├── Architecture/
└── Release/
```

Design documents describe feature implementation and UI workflow.

PRD documents describe product requirements and engineering objectives.

---

# Roadmap

Planned improvements include:

- Unified Review Framework
- Merge Target recommendation
- Rule-based merge validation
- Engineering knowledge integration
- Multi-component support (Inductor, Ferrite Bead, Crystal, IC...)
- AI-assisted merge recommendation

---

# Testing

Run unit tests:

```powershell
python -m unittest discover -s src -p "test_*.py"
```

---

# Disclaimer

This tool provides engineering recommendations only.

Merge suggestions are generated based on normalization rules and similarity analysis.

All engineering decisions should be reviewed and verified before modifying production BOM data.