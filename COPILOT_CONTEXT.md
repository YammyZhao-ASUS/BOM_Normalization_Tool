# COPILOT_CONTEXT.md

# BOM Normalization Tool - AI Development Context

> This file provides permanent development context for AI coding assistants
> (GitHub Copilot, Copilot Agent, Claude Code, ChatGPT, TRAE, etc.)
>
> Always read this document before generating or modifying code.

---

# Project Overview

Project Name

BOM Normalization Tool

Purpose

This tool is developed for Hardware R&D engineers to analyze electronic BOMs.

The goal is NOT simply to normalize component values.

The primary goal is to help Hardware Engineers reduce BOM diversity by identifying safe merge opportunities while keeping engineering risks under control.

The tool generates analysis reports only.

It NEVER modifies the original BOM automatically.

Final decisions are always made by Hardware RD.

---

# Primary Design Philosophy

Always remember:

This project is an Engineering Decision Support Tool.

NOT an Auto Merge Tool.

NOT an ERP Tool.

NOT a Database System.

The tool provides suggestions.

Hardware engineers make final decisions.

Whenever there is uncertainty,

Always send the result to Review Required instead of making assumptions.

Rule:

Safety > Automation

Correctness > Intelligence

Readability > Clever Code

Maintainability > Short Code

---

# Target Users

Hardware Engineers

PCB Designers

EE Engineers

Component Engineers

NPI Engineers

Manufacturing Engineers

Users are NOT professional software developers.

Generated outputs must be easy to understand.

---

# Coding Principles

Always follow these principles.

## 1. Do not rewrite working modules

Avoid large-scale refactoring.

Modify only related modules.

Minimize impact.

---

## 2. Backward Compatibility

Never break existing workflow.

Existing Excel format should continue working.

Existing reports should continue working.

Avoid changing public function names.

Avoid changing output sheet names.

Avoid changing output column names.

---

## 3. Modular Design

Each module should have only one responsibility.

Example:

bom_reader.py

Only read BOM.

normalizer.py

Only normalize values.

rule_engine.py

Only evaluate rules.

excel_reporter.py

Only generate Excel.

Do not mix responsibilities.

---

## 4. Keep Functions Small

Preferred

One function

One purpose

Avoid giant functions.

---

## 5. Avoid Hard Coding

Configuration belongs in

config/

instead of source code.

---

# Existing Project Structure

src/

contains all core modules.

Typical modules include

- bom_reader.py
- normalizer.py
- rule_engine.py
- similarity_analyzer.py
- component_detector.py
- excel_reporter.py
- desktop_app.py

Do not merge modules together.

---

# Rule Engine Principles

The Rule Engine is the core of the project.

Rules should be:

Transparent

Traceable

Easy to modify

Easy to add

Never hardcode business logic inside unrelated modules.

Every new rule should belong to Rule Engine.

---

# Merge Philosophy

The objective is NOT finding the most similar components.

The objective is reducing BOM complexity with minimum engineering risk.

Merge priority should always follow

1.

Lowest quantity components first.

2.

Same value.

3.

Same package.

4.

Same voltage.

5.

Same tolerance.

6.

Same dielectric.

Unknown functions should never be merged automatically.

---

# Excel Principles

Excel output is the product.

Users spend much more time reading Excel than reading code.

Therefore,

Excel readability is extremely important.

Avoid changing:

Sheet names

Column names

Column order

Output format

unless explicitly requested.

---

# Naming Convention

Python File

snake_case.py

Function

snake_case()

Class

PascalCase

Constant

UPPER_CASE

Variables

Meaningful names only.

Avoid:

tmp

aaa

bbb

test1

value2

---

# Comments

Write comments for

Complex logic

Algorithms

Engineering rules

Do not comment obvious Python syntax.

Bad

i += 1

Good

Normalize capacitor values into a unified format before comparison.

---

# Error Handling

Never crash because of

Unknown package

Unknown value

Empty cell

Unexpected description

Instead

Skip

Log

Continue

Generate warning if necessary.

---

# Performance

Typical BOM

500~5000 components

Optimization is welcome

But readability is more important than micro-optimization.

---

# Unit Tests

When modifying business logic,

Update corresponding tests.

Do not remove existing tests.

---

# Output Stability

Output format should remain stable.

Avoid introducing unnecessary changes.

The same BOM should generate identical reports.

---

# AI Development Rules

Before writing code

Please answer internally:

1.

Which module should be modified?

2.

Does this change affect Rule Engine?

3.

Will this break existing reports?

4.

Will this change Excel format?

5.

Can this feature be implemented without refactoring unrelated modules?

If the answer is uncertain,

Choose the safer implementation.

---

# When Adding New Features

Prefer

Add

instead of

Rewrite

Example

GOOD

Add new Rule 012.

BAD

Rewrite entire Rule Engine.

---

# When Unsure

Never guess engineering intent.

Instead

Generate

Review Required

so Hardware Engineers can decide.

---

# Project Vision

Current

Normalize BOM

↓

Find Merge Opportunities

↓

Reduce Component Diversity

↓

Knowledge-Based Rule Engine

↓

Engineering Intelligence Platform

↓

AI Hardware Design Assistant

Every future enhancement should move toward this vision.

---

# Reference Documents

Before making changes,

always refer to

docs/BOM_Normalization_PRD.md

If this document conflicts with the PRD,

the PRD takes precedence.

---

# End of Context