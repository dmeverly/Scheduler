# Scheduler (Workflow Integration)

**Author**: David Everly  
**Language**: Python  
**Domain**: Workflow Engineering / Operational Systems  
**Status**: Stable

---

## Description

This project implements the **execution layer** of a scheduling system, responsible for transforming abstract schedule
templates into concrete, calendar-aligned artifacts suitable for operational use.

Where upstream components focus on constraint reasoning and feasibility, this system prioritizes:

- Correctness
- Deterministic behavior
- Formatting integrity
- Repeatability

The goal is to ensure that algorithmic outputs can be reliably consumed by end users without manual correction.

---

## Why This Exists

Optimization systems often stop at producing a “solution” without addressing how that solution is operationalized.

This project bridges that gap by:

- Mapping abstract scheduling templates onto real calendar structures
- Enforcing consistent formatting and layout
- Producing artifacts directly usable in operational workflows
- Providing regression-tested calendar alignment logic

The result is a clean separation between **planning logic** and **execution logic**.

---

## Core Responsibilities

- Load predefined scheduling templates
- Collect minimal runtime configuration
- Translate template representations into calendar-aligned schedules
- Rotate a 14-week template across calendar months
- Produce formatted spreadsheet artifacts for operational use

This system intentionally does **not** perform optimization or constraint solving.

---

## Calendar Alignment

The scheduler uses **Monday-based week alignment** (ISO-style weeks).

Key behaviors:

- Deterministic date-to-cell mapping
- Correct template week rollover across month boundaries
- Accurate placement of the first and last day of each month
- Stable formatting regardless of month start day

This alignment logic is statically tested to prevent regressions.

---

## Architecture Summary

- **Input**: Structured scheduling templates (`Template.xlsx`)
- **Processing**: Deterministic mapping and validation logic
- **Output**: Calendar-aligned schedule artifacts (`Schedule_YYYY-MM-DD.xlsx`)

The system guarantees that identical inputs produce identical outputs.

---

## Reliability & Testing

The project includes static unit tests to validate:

- Monday → Sunday header alignment
- Correct date-to-column mapping
- Template week rollover logic
- Proper month boundary handling
- Accurate week-count calculation per month

This enables safe refactoring and rapid iteration when scheduling standards change (e.g., Sunday → Monday week start
migration).

Run tests:

```bash
python -m unittest test.py
```

---

## Usage

Run:

Can be executed from CLI using:

```bash
python scheduler.py <template-starting-week><template-starting-month><year>
```

If parameters are omitted, user is prompted for input runtime

---

## Dependencies

- openpyxl
- pandas

---

## Limitations and Non-Goals

- No optimization or scheduling intelligence
- No dynamic or interactive scheduling
- No persistence beyond file outputs
- These constraints are intentional and reflect the system’s role as a downstream execution layer.

---

## Relationship to Other Projects

This project operates downstream of:

- Schedule Templater, which generates feasible scheduling templates under complex constraints

Together, these projects illustrate a multi-stage system where planning and execution are explicitly decoupled.

---

## Disclaimer

This project was developed independently on personal time and is not affiliated with or endorsed by any employer.
