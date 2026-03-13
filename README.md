# WMS Dead Stock Report Generator

Automated dead stock analysis and reporting for WMS GROUP.

## Overview

This tool generates professional, modern PDF reports analyzing dead stock inventory for WMS GROUP's 7 branches.

**Supplier:** PARTS INCORPORATED (Code: AZ)  
**Definition:** Stock on hand with no sale since before 2024-01-01  
**Total Value:** 1,429 parts • R391,960.57 cost value

## Files

| File | Description |
|------|-------------|
| `PRD.md` | Product Requirements Document |
| `generate_modern_pdf.py` | PDF generator script |
| `dead_stock_*.csv` | Branch-specific data exports |
| `dead_stock_report_modern.pdf` | Generated report |

## Features

- **Modern Pitch-Deck Design** — Tech startup aesthetic
- **Visual Analytics** — Bar charts, pie charts, metric cards
- **Product Group Breakdown** — Each PGROUP as separate section
- **Part-Level Detail** — Order numbers, costs, last sold dates
- **Multi-Branch** — All 7 branches analyzed separately

## Usage

```bash
python3 generate_modern_pdf.py
```

Output: `dead_stock_report_modern.pdf`

## Requirements

- Python 3.12+
- reportlab (`pip install reportlab`)
- PostgreSQL access to d3mvdb

## Branch Summary

| Branch | Parts | Cost Value |
|--------|-------|------------|
| OLIFANTS | 374 | R99,709 |
| WADEVILLE | 231 | R61,635 |
| CHLOORKOP | 229 | R50,919 |
| DAXINA | 219 | R58,402 |
| TEMBISA | 145 | R35,684 |
| EVATON | 117 | R29,508 |
| DIEPSLOOT | 114 | R44,997 |

## Top Product Groups

| PGROUP | Description | Cost Value |
|--------|-------------|------------|
| FSP | Front Stock PIA | R51,788 |
| F02 | Air Filters | R44,821 |
| FAI | FAI | R44,765 |
| B06 | Brake Pads | R32,403 |
| G01 | Shocks | R31,736 |

---

**Generated:** 2026-03-13  
**Organization:** WMS GROUP
