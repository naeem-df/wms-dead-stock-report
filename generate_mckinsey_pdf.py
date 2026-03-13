#!/usr/bin/env python3
"""Generate McKinsey-style Dead Stock PDF Report - Black & White, Portrait"""

import csv
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, KeepTogether
)
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# McKinsey-style: Pure black and white
BLACK = colors.Color(0, 0, 0)
DARK_GRAY = colors.Color(0.2, 0.2, 0.2)
MID_GRAY = colors.Color(0.5, 0.5, 0.5)
LIGHT_GRAY = colors.Color(0.9, 0.9, 0.9)
WHITE = colors.Color(1, 1, 1)

# PGROUP descriptions
PGROUP_MAP = {
    'A03': 'BALLJOINTS', 'B01': 'EXHAUST', 'B02': 'IGNITION LEADS',
    'b06': 'BRAKE PADS', 'B06': 'BRAKE PADS', 'B09': 'BRAKE SHOES',
    'B19': 'BRAKE PARTS', 'BO': 'B/O BUY-OUT', 'BP': 'BODY PARTS',
    'BP2': 'BP TAILLAMPS', 'C05': 'NEADLE & SEATS', 'c06': 'RADIATOR',
    'C06': 'RADIATOR CAPS', 'c07': 'THERMOSTATS', 'C07': 'THERMOSTATS',
    'C08': 'F/BELTS (V NOS.)', 'C13': 'CLUTCH KITS', 'C14': 'CABLES SPEEDO',
    'c18': 'CV JOINTS', 'C18': 'CV JOINTS', 'd02': 'U JOINT & YOKES',
    'D02': 'U JOINT & YOKES', 'd03': 'DRUMS', 'D03': 'DRUMS',
    'd06': 'HUB,DISC', 'D06': 'HUB,DISC', 'd09': 'DISTRIBUTOR CAPS',
    'D09': 'DISTRIBUTOR CAPS', 'e02': 'ECHLIN POINTS & COND.',
    'E02': 'ECHLIN POINTS & COND.', 'e04': 'GLOBES', 'E04': 'GLOBES',
    'E06': 'FLASHERS & RELAYS', 'e08': 'SWITCHES TEMP.', 'E08': 'SWITCHES TEMP.',
    'E09': 'OIL SENDER UNITS', 'e10': 'SWITCH STOP LIGHT', 'E10': 'SWITCH STOP LIGHT',
    'e11': 'SWITCHES REV LIGHT', 'E11': 'SWITCHES REV LIGHT',
    'e12': 'ARM,RECT,& REGULATORS', 'E12': 'ARM,RECT,& REGULATORS',
    'E14': 'STA,& ALT REPAIR KITS', 'e19': 'BENDIX DRIVES', 'E19': 'BENDIX DRIVES',
    'E20': 'SOLENOIDS', 'E21': 'STARTOR & ALTERNATOR', 'e23': 'FAN SWITCH',
    'E23': 'FAN SWITCH', 'E25': 'COILS', 'E28': 'GLOW PLUGS',
    'F01': 'FUEL PUMPS', 'f02': 'AIR FILTERS', 'F02': 'AIR FILTERS',
    'F03': 'FRONT COUNTER', 'f05': 'CARB KITS', 'F05': 'CARB KITS',
    'F06': 'FS (SLOW MOVERS)', 'F07': 'FUEL FILTERS', 'f10': 'OIL & FUEL (CARTRIDGE)',
    'F10': 'OIL & FUEL (CARTRIDGE)', 'f11': 'OIL FILTERS', 'F11': 'OIL FILTERS',
    'fai': 'FAI', 'FAI': 'FAI', 'FSD': 'DEANS', 'FSK': 'KWH',
    'fsp': 'PIA', 'fsP': 'PIA', 'FSP': 'FRONT STOCK PIA', 'FSS': 'SHIELD',
    'FSTI': 'STICKERS', 'G01': 'SHOCKS', 'H05': 'HYDRAULIC COMPLETE CYLINDERS',
    'M01': 'MOUNTING ENG & G/BOX', 'M04': 'F/BELT (MICRO)', 'O03': 'OIL SEALS',
    'O06': 'OIL CAPS', 'P20': 'HYDRAULIC KITS', 'PROM': 'PROMO',
    'Q1': 'QUANTUM', 'R10': 'RADIATOR CAPS', 'RAL1': 'RALLY',
    'S06': 'SPARK PLUG (CHAMPION)', 's07': 'SPARK PLUGS (NGK)', 'S07': 'SPARK PLUGS (NGK)',
    'S08': 'SPARK PLUGS (BOSCH)', 's11': 'TIERODS,DRAGLINK', 'S11': 'TIERODS,DRAGLINK',
    'T01': 'SPANNERS & SOCKETS', 'T02': 'THRUST BEARINGS', 'T05': 'TIMING CHAINS',
    'T13': 'TIMING BELTS', 't15': 'TIMING BELT TENSIONERS', 'V02': 'V W SPECIAL PARTS LIST',
    'W05': 'WHEEL BEARING KITS', 'w07': 'WATER PUMPS', 'W07': 'WATER PUMPS',
    'W09': 'WIPER BLADES', 'W16': 'WINDOW TINT & STRIPE'
}

def load_all_data(base_path):
    """Load all CSV data"""
    branches = ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']
    all_data = {}
    for branch in branches:
        csv_path = os.path.join(base_path, f'dead_stock_{branch}.csv')
        if os.path.exists(csv_path):
            data = []
            with open(csv_path, 'r') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    data.append(row)
            all_data[branch] = data
    return all_data

def calculate_branch_totals(branch_data):
    """Calculate totals for a branch"""
    total_cost = sum(float(row.get('cost_value', 0) or 0) for row in branch_data)
    total_qty = sum(int(row.get('qoh', 0) or 0) for row in branch_data)
    return total_cost, total_qty

def create_simple_bar_chart(data_dict, title, width=450, height=150):
    """Create simple B&W bar chart"""
    drawing = Drawing(width, height)
    
    labels = list(data_dict.keys())
    values = list(data_dict.values())
    
    bc = VerticalBarChart()
    bc.x = 50
    bc.y = 30
    bc.height = 90
    bc.width = width - 100
    bc.data = [values]
    bc.strokeColor = BLACK
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = max(values) * 1.15 if values else 100
    bc.valueAxis.strokeColor = MID_GRAY
    bc.valueAxis.labels.fontName = 'Helvetica'
    bc.valueAxis.labels.fontSize = 8
    bc.valueAxis.labels.fillColor = DARK_GRAY
    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.dx = 5
    bc.categoryAxis.labels.dy = -2
    bc.categoryAxis.labels.angle = 45
    bc.categoryAxis.labels.fontName = 'Helvetica'
    bc.categoryAxis.labels.fontSize = 8
    bc.categoryAxis.labels.fillColor = DARK_GRAY
    bc.categoryAxis.categoryNames = [l[:3].upper() for l in labels]
    bc.bars[0].fillColor = DARK_GRAY
    bc.bars[0].strokeColor = BLACK
    
    drawing.add(bc)
    return drawing

class PageHeaderCanvas:
    """Canvas with McKinsey-style page headers"""
    def __init__(self, doc, branch_name=None):
        self.doc = doc
        self.branch_name = branch_name
    
    def __call__(self, canvas, doc):
        canvas.saveState()
        
        # Header line
        canvas.setStrokeColor(BLACK)
        canvas.setLineWidth(0.5)
        canvas.line(doc.leftMargin, doc.height + doc.topMargin - 10, 
                   doc.width + doc.leftMargin, doc.height + doc.topMargin - 10)
        
        # Header text - McKinsey style
        canvas.setFont('Times-Bold', 10)
        canvas.setFillColor(DARK_GRAY)
        canvas.drawString(doc.leftMargin, doc.height + doc.topMargin - 5, 
                         "WMS GROUP | Dead Stock Analysis")
        
        if self.branch_name:
            canvas.setFont('Times-Bold', 12)
            canvas.setFillColor(BLACK)
            canvas.drawRightString(doc.width + doc.leftMargin, doc.height + doc.topMargin - 5,
                                  self.branch_name.upper())
        
        # Footer
        canvas.setFont('Helvetica', 8)
        canvas.setFillColor(MID_GRAY)
        canvas.drawString(doc.leftMargin, 0.5*inch, 
                         f"PARTS INCORPORATED (AZ) | Generated {datetime.now().strftime('%Y-%m-%d')}")
        canvas.drawRightString(doc.width + doc.leftMargin, 0.5*inch, f"Page {doc.page}")
        
        canvas.restoreState()

def create_pgroup_table(parts_data, pgroup_name, pgroup_desc):
    """Create clean table for a product group"""
    sorted_parts = sorted(parts_data, key=lambda x: float(x.get('cost_value', 0) or 0), reverse=True)
    
    table_data = [['Part No', 'Description', 'Qty', 'Unit Cost', 'Total Cost', 'Order Ref', 'Last Sold']]
    
    for row in sorted_parts:
        partno = row.get('partno', '')[:16]
        desc = row.get('sdesc', '')[:20]
        qoh = row.get('qoh', '0')
        lcost = float(row.get('lcost', 0) or 0)
        cost_val = float(row.get('cost_value', 0) or 0)
        order_ref = (row.get('last_order_ref') or '—')[:10]
        lsold = (row.get('lsold') or '—')[:10]
        
        table_data.append([
            partno,
            desc,
            qoh,
            f"R{lcost:.2f}",
            f"R{cost_val:,.2f}",
            order_ref,
            lsold
        ])
    
    # Calculate PGROUP total
    pg_total = sum(float(r.get('cost_value', 0) or 0) for r in parts_data)
    pg_qty = sum(int(r.get('qoh', 0) or 0) for r in parts_data)
    
    # Add subtotal row
    table_data.append(['', f'{pgroup_desc} SUBTOTAL', f'{pg_qty:,}', '', f'R{pg_total:,.2f}', '', ''])
    
    table = Table(table_data, colWidths=[70, 90, 35, 55, 65, 55, 55])
    table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Body
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -2), 7),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('ALIGN', (3, 0), (4, -1), 'RIGHT'),
        ('ALIGN', (5, 0), (6, -1), 'CENTER'),
        
        # Subtotal row
        ('BACKGROUND', (0, -1), (-1, -1), LIGHT_GRAY),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 7),
        ('LINEABOVE', (0, -1), (-1, -1), 1, BLACK),
        
        # Grid
        ('GRID', (0, 0), (-1, -1), 0.25, MID_GRAY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, colors.Color(0.97, 0.97, 0.97)]),
        
        # Padding
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
    ]))
    
    return table

def create_pdf_report(base_path, output_path):
    """Create McKinsey-style PDF report"""
    
    all_data = load_all_data(base_path)
    
    total_parts = sum(len(data) for data in all_data.values())
    total_cost = sum(calculate_branch_totals(data)[0] for data in all_data.values())
    total_qty = sum(calculate_branch_totals(data)[1] for data in all_data.values())
    
    # Calculate pgroup totals
    pgroup_totals = {}
    branch_pgroup_data = {}
    
    for branch, data in all_data.items():
        if branch not in branch_pgroup_data:
            branch_pgroup_data[branch] = {}
        
        for row in data:
            pg = row.get('pgroup', 'UNKNOWN')
            
            if pg not in pgroup_totals:
                pgroup_totals[pg] = {'parts': 0, 'qty': 0, 'cost': 0}
            pgroup_totals[pg]['parts'] += 1
            pgroup_totals[pg]['qty'] += int(row.get('qoh', 0) or 0)
            pgroup_totals[pg]['cost'] += float(row.get('cost_value', 0) or 0)
            
            if pg not in branch_pgroup_data[branch]:
                branch_pgroup_data[branch][pg] = []
            branch_pgroup_data[branch][pg].append(row)
    
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=0.6*inch,
        leftMargin=0.6*inch,
        topMargin=0.8*inch,
        bottomMargin=0.7*inch
    )
    
    styles = getSampleStyleSheet()
    
    # McKinsey-style: Serif headings, sans-serif body
    main_title = ParagraphStyle(
        'MainTitle',
        fontSize=24,
        textColor=BLACK,
        spaceAfter=6,
        alignment=TA_CENTER,
        fontName='Times-Bold',
        leading=28
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        fontSize=12,
        textColor=DARK_GRAY,
        spaceAfter=20,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )
    
    section_header = ParagraphStyle(
        'SectionHeader',
        fontSize=16,
        textColor=BLACK,
        spaceBefore=20,
        spaceAfter=10,
        fontName='Times-Bold'
    )
    
    branch_header = ParagraphStyle(
        'BranchHeader',
        fontSize=18,
        textColor=BLACK,
        spaceBefore=0,
        spaceAfter=4,
        fontName='Times-Bold'
    )
    
    pg_header = ParagraphStyle(
        'PGHeader',
        fontSize=11,
        textColor=BLACK,
        spaceBefore=12,
        spaceAfter=4,
        fontName='Times-Bold'
    )
    
    body_style = ParagraphStyle(
        'Body',
        fontSize=9,
        textColor=DARK_GRAY,
        spaceAfter=6,
        fontName='Helvetica',
        leading=12
    )
    
    story = []
    
    # ========== COVER PAGE ==========
    story.append(Spacer(1, 2*inch))
    story.append(Paragraph("DEAD STOCK ANALYSIS", main_title))
    story.append(Paragraph("PARTS INCORPORATED  |  SUPPLIER CODE: AZ", subtitle_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Key metrics table
    metrics = [
        ['Total Parts', 'Total Quantity', 'Cost Value', 'Branches'],
        [f'{total_parts:,}', f'{total_qty:,}', f'R{total_cost:,.0f}', f'{len(all_data)}']
    ]
    
    metrics_table = Table(metrics, colWidths=[100, 100, 100, 100])
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), LIGHT_GRAY),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, -1), 11),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('BOX', (0, 0), (-1, -1), 1, BLACK),
    ]))
    story.append(metrics_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Definition
    story.append(Paragraph("<b>Definition:</b> Stock on hand (QOH > 0) with no sale recorded since before 2024-01-01. Values at cost price.", 
                          ParagraphStyle('Def', fontSize=9, textColor=DARK_GRAY, fontName='Helvetica')))
    story.append(Spacer(1, 0.4*inch))
    
    # ========== BRANCH SUMMARY ==========
    story.append(Paragraph("Branch Summary", section_header))
    
    branch_summary = [['Branch', 'Parts', 'Quantity', 'Cost Value', '% of Total']]
    for branch in ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']:
        if branch in all_data:
            cost, qty = calculate_branch_totals(all_data[branch])
            pct = (cost / total_cost * 100) if total_cost > 0 else 0
            branch_summary.append([
                branch.title(),
                f"{len(all_data[branch]):,}",
                f"{qty:,}",
                f"R{cost:,.2f}",
                f"{pct:.1f}%"
            ])
    
    # Add total row
    branch_summary.append(['TOTAL', f"{total_parts:,}", f"{total_qty:,}", f"R{total_cost:,.2f}", "100.0%"])
    
    branch_table = Table(branch_summary, colWidths=[90, 60, 70, 90, 60])
    branch_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('BACKGROUND', (0, -1), (-1, -1), LIGHT_GRAY),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('LINEABOVE', (0, -1), (-1, -1), 1, BLACK),
        ('GRID', (0, 0), (-1, -1), 0.25, MID_GRAY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, colors.Color(0.98, 0.98, 0.98)]),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(branch_table)
    
    # Bar chart
    story.append(Spacer(1, 0.3*inch))
    branch_costs = {b[:3].upper(): calculate_branch_totals(all_data[b])[0]/1000 
                   for b in ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']
                   if b in all_data}
    story.append(create_simple_bar_chart(branch_costs, "Cost Value by Branch (R thousands)"))
    
    story.append(PageBreak())
    
    # ========== PRODUCT GROUP OVERVIEW ==========
    story.append(Paragraph("Product Group Analysis", section_header))
    
    pg_summary = [['Rank', 'PGroup', 'Description', 'Parts', 'Qty', 'Cost Value', '% Total']]
    sorted_pg = sorted(pgroup_totals.items(), key=lambda x: x[1]['cost'], reverse=True)[:20]
    
    for i, (pg, data) in enumerate(sorted_pg, 1):
        pct = (data['cost'] / total_cost * 100) if total_cost > 0 else 0
        desc = PGROUP_MAP.get(pg, '—')[:25]
        pg_summary.append([
            str(i),
            pg,
            desc,
            f"{data['parts']:,}",
            f"{data['qty']:,}",
            f"R{data['cost']:,.2f}",
            f"{pct:.1f}%"
        ])
    
    pg_table = Table(pg_summary, colWidths=[30, 45, 100, 45, 45, 75, 45])
    pg_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 0.25, MID_GRAY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, colors.Color(0.98, 0.98, 0.98)]),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(pg_table)
    
    story.append(PageBreak())
    
    # ========== BRANCH DETAILS ==========
    for branch in ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']:
        if branch not in all_data:
            continue
        
        branch_cost, branch_qty = calculate_branch_totals(all_data[branch])
        branch_pg_data = branch_pgroup_data.get(branch, {})
        
        # Branch header - McKinsey style
        story.append(Paragraph(branch.title(), branch_header))
        story.append(Paragraph(f"{len(all_data[branch]):,} parts  •  {branch_qty:,} units  •  R{branch_cost:,.0f} cost value", 
                              ParagraphStyle('BranchSum', fontSize=9, textColor=DARK_GRAY, fontName='Helvetica')))
        story.append(Spacer(1, 0.15*inch))
        
        # Sort product groups by cost
        sorted_pg = sorted(branch_pg_data.items(), 
                         key=lambda x: sum(float(r.get('cost_value', 0) or 0) for r in x[1]), 
                         reverse=True)
        
        # Create tables for each product group
        for pg, parts in sorted_pg:
            pg_desc = PGROUP_MAP.get(pg, '—')
            
            # PGROUP header
            story.append(Paragraph(f"<b>{pg}</b> — {pg_desc}", pg_header))
            story.append(Spacer(1, 0.08*inch))
            
            # Parts table
            pg_table = create_pgroup_table(parts, pg, pg_desc)
            story.append(pg_table)
            story.append(Spacer(1, 0.15*inch))
        
        story.append(PageBreak())
    
    doc.build(story)
    return output_path

if __name__ == '__main__':
    base_path = '/root/.openclaw/workspace/dead_stock_reports'
    output_path = os.path.join(base_path, 'dead_stock_report_mckinsey.pdf')
    create_pdf_report(base_path, output_path)
    print(f"PDF created: {output_path}")
