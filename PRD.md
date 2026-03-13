# PRD: WMS Dead Stock Report Generator

## Overview

**Product:** Automated dead stock analysis and reporting system for WMS GROUP  
**Purpose:** Identify, quantify, and report on slow-moving inventory to enable supplier returns and stock optimization  
**Target User:** Naeem Asvat, Managing Director, WMS GROUP

---

## Problem Statement

WMS GROUP carries dead stock (no sale since before 2024-01-01) across 7 branches for supplier PARTS INCORPORATED (AZ). Manual analysis is time-consuming. Need automated, professional reports for supplier negotiations and stock management.

---

## Goals

**Primary:**
- Generate comprehensive dead stock reports by supplier, branch, and product group
- Enable supplier returns with order reference numbers
- Quantify financial impact (cost value, retail value)

**Secondary:**
- Track dead stock trends over time
- Identify patterns by product group
- Support credit note negotiations

---

## Success Metrics

| Metric | Target |
|--------|--------|
| Report generation time | < 2 minutes |
| Data accuracy | 100% match to d3mvdb |
| Supplier return rate | 50%+ of dead stock value |
| User satisfaction | Naeem approves design |

---

## Features

### Core Features

**A. Data Extraction**
- Connect to d3mvdb PostgreSQL database
- Filter by supplier code (AZ)
- Identify dead stock (qoh > 0, lsold < 2024-01-01)
- Pull last order reference from patranhist

**B. Report Generation**
- PDF with modern pitch-deck design
- CSV exports by branch
- Breakdown by product group
- Part-level detail with order numbers

**C. Visualizations**
- Cost value by branch (bar chart)
- Product group distribution (pie chart)
- Financial summary (infographic cards)
- Product group breakdown tables

### Output Requirements

**PDF Report:**
1. Cover page with key metrics
2. Executive summary with charts
3. Product group analysis
4. Branch-by-branch detail (7 sections)
5. Each PGROUP as sub-section with part-level table

**CSV Exports:**
- 1 file per branch
- Columns: pgroup, suppname, partno, sdesc, lcost, retail, qoh, cost_value, lsold, last_order_ref, last_receipt_date

---

## Technical Requirements

**Database:**
- Host: db-postgresql-lon1-sparescentre-do-user-4750606-0.b.db.ondigitalocean.com
- Port: 25060
- Database: d3mvdb
- User: doadmin

**Python Libraries:**
- reportlab (PDF generation)
- psycopg2 (PostgreSQL)
- csv (data export)

**Design System:**
- Modern tech startup aesthetic
- Gradient backgrounds
- Clean typography
- Card-style layouts
- Infographic elements

---

## Constraints

- No external API keys (use database directly)
- No sensitive data in repo (passwords in .gitignore)
- Must work offline after initial setup
- Output must be print-ready

---

## Timeline

| Phase | Deliverable | Date |
|-------|-------------|------|
| 1 | PRD + Repo setup | 2026-03-13 |
| 2 | Basic PDF with data | 2026-03-13 |
| 3 | Modern design upgrade | 2026-03-13 |
| 4 | GitHub push + documentation | 2026-03-13 |

---

## Acceptance Criteria

**Must Have:**
- [x] Connect to d3mvdb
- [x] Extract dead stock for PARTS INCORPORATED
- [x] Generate PDF with charts
- [x] Generate CSVs by branch
- [x] Include order reference numbers
- [ ] Modern pitch-deck design
- [ ] Part-level tables by PGROUP
- [ ] GitHub repo with PRD

**Should Have:**
- [ ] Multiple supplier support
- [ ] Date range selector
- [ ] Historical comparison

**Won't Have (v1):**
- Web interface
- Email automation
- Real-time updates

---

## Maintenance

**Scheduled runs:**
- Monthly dead stock report
- Quarterly supplier review

**Updates needed:**
- New supplier codes
- Date threshold changes
- Product group mappings

---

## Changelog

| Date | Version | Change |
|------|---------|--------|
| 2026-03-13 | 1.0 | Initial PRD |
