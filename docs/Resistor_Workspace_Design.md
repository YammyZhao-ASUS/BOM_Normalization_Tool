# Design Document
## BOM Normalization Tool – Resistor Merge Workspace

**Version:** V1.0  
**Status:** Implemented  
**Priority:** High  
**Owner:** BOM Normalization Tool

---

# 1. Background

The Capacitor module provides a complete review workflow through the **Capacitor Merge Workspace**, allowing engineers to:

- Review merge candidates
- Select BOM Action
- Track review status
- Synchronize decisions back to Summary

Previously, the Resistor module only provided a hyperlink from **Resistor Summary** to **Resistor Detail**, resulting in a different user experience.

To improve usability and maintain consistency across the tool, the Resistor workflow should be redesigned to match the Capacitor Merge Workspace.

---

# 2. Objectives

This design focuses on **workflow consistency**, not algorithm changes.

Objectives:

- Unify the user experience between Capacitor and Resistor.
- Reduce learning effort for engineers.
- Improve review efficiency.
- Reuse existing UI components and logic.
- Minimize duplicated implementation.
- Build a scalable review framework for future component categories.

---

# 3. Previous Workflow

```
Resistor Summary
        │
        ▼
Click Detail Hyperlink
        │
        ▼
Read-only Resistor Detail
        │
        ▼
Manual Review
```

### Current Issues

- Summary only provides hyperlinks.
- Review decisions are not stored.
- Summary cannot display review status.
- Workflow is inconsistent with Capacitor.
- Engineers must switch between pages without any visible review progress.

---

# 4. Proposed Workflow

```
Resistor Summary
        │
        ▼
Click Target PN
        │
        ▼
Resistor Merge Workspace
        │
        ▼
Select BOM ACTION
(Merge / Review / Keep)
        │
        ▼
Summary updates automatically
```

The workflow should behave exactly like the existing Capacitor Merge Workspace.

---

# 5. Functional Design

## 5.1 Convert Resistor Detail into Resistor Merge Workspace

The previous Resistor Detail sheet was read-only.

Convert it into an interactive review workspace by adding an editable **BOM ACTION** column.

### Available Actions

| Action | Meaning |
|---------|---------|
| Merge | Candidate approved for merge |
| Review | Requires further engineering review |
| Keep | Keep current resistor without merging |

### Visual Style

Use one shared BOM ACTION color scheme for all component workspaces. Color represents action status, not component category:

| Status | Color |
|---------|-------|
| 🟢 Merge | Green |
| 🟡 Review | Yellow |
| ⚪ Keep | Gray / White |

The dropdown style, conditional formatting, icons, and status colors must be identical for Capacitor Merge Workspace, Resistor Merge Workspace, and future component merge workspaces.

---

## 5.2 Summary Synchronization

The Summary page should always display the latest review result.

Current Summary:

| Target PN | Detail | Candidate Count |

Proposed Summary:

| Target PN | BOM ACTION | Candidate Count |

Whenever BOM ACTION changes inside the Detail Workspace, the Summary page should update automatically.

Example:

Detail:

```
Target PN : 10K

BOM ACTION

▼ Merge
```

↓

Summary:

```
10K

🟢 Merge
```

No manual editing should be required.

---

## 5.3 Keep Existing Navigation

The hyperlink navigation should remain.

Users should still be able to jump directly from Summary to the corresponding location in Resistor Merge Workspace.

Suggested clickable fields:

- Target PN
- Candidate Count

Navigation behavior should not change.

---

## 5.4 Merge Target (Recommended)

To simplify future BOM updates, add another optional editable column.

| Column | Description |
|---------|-------------|
| Merge Target | Existing resistor selected as merge destination |

Example:

| BOM ACTION | Merge Target |
|------------|--------------|
| Merge | RC0402_10K_1% |

Summary can display:

```
🟢 Merge
→ RC0402_10K_1%
```

instead of only:

```
Merge
```

This allows engineers to immediately understand which existing resistor should become the final standardized component.

---

# 6. UI Consistency

The Resistor Workspace should reuse as much of the existing Capacitor Workspace implementation as possible.

Reuse:

- Dropdown controls
- Conditional formatting
- Icons
- Colors
- Cell layout
- Review workflow
- Status synchronization logic

No separate UI design should be introduced.

---

# 7. Future Extensibility

This review framework should become the standard architecture for all component categories.

Current:

- Capacitor ✅
- Resistor ✅

Future:

- Inductor
- Ferrite Bead
- Crystal
- IC
- Connector
- Diode
- MOSFET

Every component should follow the same workflow:

```
Summary
    │
    ▼
Detail Workspace
    │
    ▼
Review Decision
    │
    ▼
Automatic Synchronization
```

This provides a consistent engineering experience across the entire BOM Normalization Tool.

---

# 8. Expected Benefits

| Area | Benefit |
|------|----------|
| User Experience | One consistent workflow across all component types |
| Review Efficiency | Faster engineering decisions |
| Visibility | Summary always reflects current review status |
| Maintainability | Maximum reuse of existing Capacitor implementation |
| Scalability | Future components require minimal UI redesign |

---

# 9. Non-Goals

This proposal does **not** modify:

- Resistor normalization algorithm
- Similarity scoring
- Candidate ranking
- Merge recommendation logic
- Report generation logic

This document only covers **UI workflow improvements**.

---

# 10. Implementation Priority

**Priority:** High (Recommended for V1.0)

Reasons:

- Significant usability improvement
- Minimal implementation risk
- Reuses mature Capacitor components
- No algorithm changes required
- Establishes the standard review workflow for future development

---

# 11. Design Principle

## One Review Framework for All Component Types

The BOM Normalization Tool should provide a single, unified review experience regardless of component category.

Standard workflow:

```
Summary
    │
    ▼
Detail Workspace
    │
    ▼
BOM ACTION
    │
    ▼
Automatic Synchronization
```

This design principle reduces duplicated code, improves usability, and creates a scalable architecture that future component modules can adopt with minimal effort.

---

# 12. Development Notes

Implementation should prioritize **reusing the existing Capacitor Merge Workspace** wherever possible.

Recommended approach:

- Reuse the existing dropdown component.
- Reuse conditional formatting rules.
- Reuse synchronization logic.
- Reuse worksheet layout where applicable.
- Adapt only the data source and resistor-specific fields.

The goal is **not to build a second independent review system**, but to extend the existing review framework so that both Capacitor and Resistor share the same interaction model and maintenance strategy.