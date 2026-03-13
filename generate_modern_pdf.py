#!/usr/bin/env python3
"""Generate ultra-modern pitch-deck style Dead Stock PDF Report"""

import csv
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, KeepTogether, Flowable
)
from reportlab.graphics.shapes import Drawing, Rect, String, Circle
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.pdfgen import canvas

# Modern startup pitch deck colors
MODERN_BLACK = colors.Color(0.05, 0.05, 0.08)
MODERN_DARK = colors.Color(0.1, 0.1, 0.15)
MODERN_GRAY = colors.Color(0.4, 0.4, 0.45)
MODERN_LIGHT = colors.Color(0.95, 0.95, 0.97)
MODERN_ACCENT = colors.Color(0.2, 0.6, 1.0)  # Bright blue
MODERN_SUCCESS = colors.Color(0.1, 0.8, 0.5)  # Green
MODERN_WARNING = colors.Color(1.0, 0.6, 0.2)  # Orange
MODERN_PURPLE = colors.Color(0.6, 0.3, 0.9)
MODERN_PINK = colors.Color(1.0, 0.3, 0.5)

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

class ModernCard(Flowable):
    """Modern card-style flowable"""
    def __init__(self, width, height, title, value, subtitle="", accent_color=MODERN_ACCENT):
        Flowable.__init__(self)
        self.width = width
        self.height = height
        self.title = title
        self.value = value
        self.subtitle = subtitle
        self.accent_color = accent_color
    
    def draw(self):
        # Background
        self.canv.setFillColor(MODERN_DARK)
        self.canv.roundRect(0, 0, self.width, self.height, 8, fill=1, stroke=0)
        
        # Accent line
        self.canv.setFillColor(self.accent_color)
        self.canv.rect(0, self.height - 4, self.width, 4, fill=1, stroke=0)
        
        # Title
        self.canv.setFillColor(MODERN_GRAY)
        self.canv.setFont("Helvetica", 10)
        self.canv.drawString(12, self.height - 25, self.title)
        
        # Value
        self.canv.setFillColor(MODERN_LIGHT)
        self.canv.setFont("Helvetica-Bold", 20)
        self.canv.drawString(12, self.height - 55, self.value)
        
        # Subtitle
        if self.subtitle:
            self.canv.setFillColor(MODERN_GRAY)
            self.canv.setFont("Helvetica", 8)
            self.canv.drawString(12, self.height - 72, self.subtitle)

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

def create_modern_bar_chart(branch_data):
    """Create modern bar chart"""
    drawing = Drawing(480, 160)
    
    branches = list(branch_data.keys())
    values = [calculate_branch_totals(branch_data[b])[0] / 1000 for b in branches]
    
    # Background
    drawing.add(Rect(0, 0, 480, 160, fillColor=MODERN_DARK, strokeColor=None))
    
    bc = VerticalBarChart()
    bc.x = 40
    bc.y = 20
    bc.height = 110
    bc.width = 400
    bc.data = [values]
    bc.strokeColor = None
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = max(values) * 1.2 if values else 100
    bc.valueAxis.strokeColor = MODERN_GRAY
    bc.valueAxis.labels.fontName = 'Helvetica'
    bc.valueAxis.labels.fontSize = 8
    bc.valueAxis.labels.fillColor = MODERN_LIGHT
    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.dx = 8
    bc.categoryAxis.labels.dy = -2
    bc.categoryAxis.labels.angle = 45
    bc.categoryAxis.labels.fontName = 'Helvetica'
    bc.categoryAxis.labels.fontSize = 8
    bc.categoryAxis.labels.fillColor = MODERN_LIGHT
    bc.categoryAxis.categoryNames = [b[:3].upper() for b in branches]
    bc.bars[0].fillColor = MODERN_ACCENT
    bc.bars[0].strokeColor = None
    
    drawing.add(bc)
    return drawing

def create_modern_pie_chart(pgroup_totals):
    """Create modern pie chart"""
    drawing = Drawing(350, 160)
    
    drawing.add(Rect(0, 0, 350, 160, fillColor=MODERN_DARK, strokeColor=None))
    
    sorted_pg = sorted(pgroup_totals.items(), key=lambda x: x[1]['cost'], reverse=True)[:8]
    
    pie = Pie()
    pie.x = 40
    pie.y = 20
    pie.width = 100
    pie.height = 100
    pie.data = [pg[1]['cost'] / 1000 for pg in sorted_pg]
    pie.labels = None  # We'll add a legend instead
    
    pie_colors = [MODERN_ACCENT, MODERN_SUCCESS, MODERN_WARNING, MODERN_PURPLE,
                  MODERN_PINK, colors.Color(0.3, 0.7, 0.8), colors.Color(0.8, 0.5, 0.2),
                  colors.Color(0.5, 0.5, 0.6)]
    
    for i, color in enumerate(pie_colors[:len(sorted_pg)]):
        pie.slices[i].fillColor = color
        pie.slices[i].strokeColor = None
    
    drawing.add(pie)
    
    # Legend
    y_pos = 140
    for i, (pg, data) in enumerate(sorted_pg[:8]):
        drawing.add(Rect(160, y_pos - 10, 10, 10, fillColor=pie_colors[i], strokeColor=None))
        drawing.add(String(175, y_pos - 8, f"{pg} - R{data['cost']/1000:.0f}K", 
                          fontName='Helvetica', fontSize=8, fillColor=MODERN_LIGHT))
        y_pos -= 15
    
    return drawing

def create_pgroup_part_table(pgroup_name, pgroup_desc, parts_data):
    """Create a modern table for parts in a product group"""
    # Sort by cost value descending
    sorted_parts = sorted(parts_data, key=lambda x: float(x.get('cost_value', 0) or 0), reverse=True)
    
    table_data = [['Part No', 'Description', 'Qty', 'Cost', 'Cost Value', 'Last Order', 'Last Sold']]
    
    for row in sorted_parts:
        partno = row.get('partno', '')[:18]
        desc = row.get('sdesc', '')[:22]
        qoh = row.get('qoh', '0')
        lcost = f"R{float(row.get('lcost', 0) or 0):.2f}"
        cost_val = f"R{float(row.get('cost_value', 0) or 0):.2f}"
        order_ref = row.get('last_order_ref', 'N/A')[:12] if row.get('last_order_ref') else 'N/A'
        lsold = row.get('lsold', 'N/A')[:10] if row.get('lsold') else 'N/A'
        
        table_data.append([partno, desc, qoh, lcost, cost_val, order_ref, lsold])
    
    table = Table(table_data, colWidths=[80, 100, 35, 50, 60, 70, 65])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), MODERN_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), MODERN_LIGHT),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (1, -1), 'LEFT'),
        ('BACKGROUND', (0, 1), (-1, -1), MODERN_DARK),
        ('TEXTCOLOR', (0, 1), (-1, -1), MODERN_LIGHT),
        ('GRID', (0, 0), (-1, -1), 0.5, MODERN_GRAY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [MODERN_DARK, colors.Color(0.12, 0.12, 0.18)]),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    
    return table

def create_pdf_report(base_path, output_path):
    """Create the ultra-modern PDF report"""
    
    all_data = load_all_data(base_path)
    
    total_parts = sum(len(data) for data in all_data.values())
    total_cost = sum(calculate_branch_totals(data)[0] for data in all_data.values())
    total_qty = sum(calculate_branch_totals(data)[1] for data in all_data.values())
    
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
        pagesize=landscape(A4),
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    
    styles = getSampleStyleSheet()
    
    # Modern styles
    title_style = ParagraphStyle(
        'ModernTitle',
        fontSize=42,
        textColor=MODERN_LIGHT,
        spaceAfter=10,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        leading=48
    )
    
    subtitle_style = ParagraphStyle(
        'ModernSubtitle',
        fontSize=16,
        textColor=MODERN_GRAY,
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )
    
    section_style = ParagraphStyle(
        'Section',
        fontSize=24,
        textColor=MODERN_LIGHT,
        spaceAfter=15,
        spaceBefore=20,
        fontName='Helvetica-Bold'
    )
    
    pg_header_style = ParagraphStyle(
        'PGHeader',
        fontSize=16,
        textColor=MODERN_ACCENT,
        spaceAfter=8,
        spaceBefore=15,
        fontName='Helvetica-Bold'
    )
    
    story = []
    
    # ========== COVER PAGE ==========
    story.append(Spacer(1, 1.5*inch))
    story.append(Paragraph("DEAD STOCK<br/>ANALYSIS", title_style))
    story.append(Paragraph("PARTS INCORPORATED  •  SUPPLIER CODE: AZ", subtitle_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Modern metric cards
    card_table = Table([
        [ModernCard(2.2*inch, 0.9*inch, "TOTAL PARTS", f"{total_parts:,}", "Unique SKUs", MODERN_ACCENT),
         ModernCard(2.2*inch, 0.9*inch, "TOTAL QUANTITY", f"{total_qty:,}", "Units on hand", MODERN_SUCCESS),
         ModernCard(2.2*inch, 0.9*inch, "COST VALUE", f"R{total_cost:,.0f}", "At cost", MODERN_WARNING),
         ModernCard(2.2*inch, 0.9*inch, "BRANCHES", f"{len(all_data)}", "Locations", MODERN_PURPLE)]
    ], colWidths=[2.4*inch]*4)
    card_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), MODERN_BLACK),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(card_table)
    
    story.append(Spacer(1, 0.5*inch))
    story.append(Paragraph(f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}  •  No sale since before 2024-01-01", 
                          ParagraphStyle('Date', fontSize=10, textColor=MODERN_GRAY, alignment=TA_CENTER)))
    
    story.append(PageBreak())
    
    # ========== EXECUTIVE SUMMARY ==========
    story.append(Paragraph("Executive Summary", section_style))
    
    # Branch summary table - modern style
    branch_summary = [['BRANCH', 'PARTS', 'QUANTITY', 'COST VALUE', '% TOTAL']]
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
    
    branch_table = Table(branch_summary, colWidths=[1.5*inch, 1*inch, 1*inch, 1.5*inch, 1*inch])
    branch_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), MODERN_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), MODERN_LIGHT),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('BACKGROUND', (0, 1), (-1, -1), MODERN_DARK),
        ('TEXTCOLOR', (0, 1), (-1, -1), MODERN_LIGHT),
        ('GRID', (0, 0), (-1, -1), 0.5, MODERN_GRAY),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(branch_table)
    
    story.append(Spacer(1, 0.3*inch))
    story.append(create_modern_bar_chart(all_data))
    
    story.append(PageBreak())
    
    # ========== PRODUCT GROUP OVERVIEW ==========
    story.append(Paragraph("Product Group Overview", section_style))
    
    pg_summary = [['RANK', 'PGROUP', 'DESCRIPTION', 'PARTS', 'QTY', 'COST VALUE', '% TOTAL']]
    sorted_pg = sorted(pgroup_totals.items(), key=lambda x: x[1]['cost'], reverse=True)[:15]
    
    for i, (pg, data) in enumerate(sorted_pg, 1):
        pct = (data['cost'] / total_cost * 100) if total_cost > 0 else 0
        desc = PGROUP_MAP.get(pg, 'N/A')[:25]
        pg_summary.append([
            str(i),
            pg,
            desc,
            f"{data['parts']:,}",
            f"{data['qty']:,}",
            f"R{data['cost']:,.2f}",
            f"{pct:.1f}%"
        ])
    
    pg_table = Table(pg_summary, colWidths=[0.4*inch, 0.7*inch, 1.8*inch, 0.6*inch, 0.6*inch, 1.1*inch, 0.7*inch])
    pg_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), MODERN_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), MODERN_LIGHT),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('BACKGROUND', (0, 1), (-1, -1), MODERN_DARK),
        ('TEXTCOLOR', (0, 1), (-1, -1), MODERN_LIGHT),
        ('GRID', (0, 0), (-1, -1), 0.5, MODERN_GRAY),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(pg_table)
    
    story.append(Spacer(1, 0.2*inch))
    story.append(create_modern_pie_chart(pgroup_totals))
    
    story.append(PageBreak())
    
    # ========== BRANCH DETAILS WITH PGROUP BREAKDOWN ==========
    for branch in ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']:
        if branch not in all_data:
            continue
        
        branch_cost, branch_qty = calculate_branch_totals(all_data[branch])
        branch_pg_data = branch_pgroup_data.get(branch, {})
        
        # Branch header
        story.append(Paragraph(f"{branch.title()}", section_style))
        
        # Branch summary metrics
        summary_text = f"<font color='#666666'>{len(all_data[branch]):,} parts  •  {branch_qty:,} units  •  R{branch_cost:,.0f} cost value</font>"
        story.append(Paragraph(summary_text, ParagraphStyle('Summary', fontSize=11, textColor=MODERN_GRAY)))
        story.append(Spacer(1, 0.2*inch))
        
        # Sort product groups by cost value
        sorted_pg = sorted(branch_pg_data.items(), 
                         key=lambda x: sum(float(r.get('cost_value', 0) or 0) for r in x[1]), 
                         reverse=True)
        
        # Create tables for each product group (top 10 for space)
        for pg, parts in sorted_pg[:10]:
            pg_desc = PGROUP_MAP.get(pg, 'N/A')
            pg_cost = sum(float(r.get('cost_value', 0) or 0) for r in parts)
            pg_qty = sum(int(r.get('qoh', 0) or 0) for r in parts)
            
            # PGROUP header
            story.append(Paragraph(f"{pg} — {pg_desc}", pg_header_style))
            story.append(Paragraph(f"<font color='#666666'>{len(parts)} parts  •  {pg_qty} units  •  R{pg_cost:,.0f}</font>", 
                                 ParagraphStyle('PGSub', fontSize=9, textColor=MODERN_GRAY)))
            story.append(Spacer(1, 0.1*inch))
            
            # Parts table
            part_table = create_pgroup_part_table(pg, pg_desc, parts)
            story.append(part_table)
            story.append(Spacer(1, 0.2*inch))
        
        # If more than 10 product groups, show summary
        if len(sorted_pg) > 10:
            remaining = len(sorted_pg) - 10
            story.append(Paragraph(f"<font color='#666666'>+ {remaining} more product groups</font>", 
                                 ParagraphStyle('More', fontSize=9, textColor=MODERN_GRAY)))
        
        story.append(PageBreak())
    
    doc.build(story)
    return output_path

if __name__ == '__main__':
    base_path = '/root/.openclaw/workspace/dead_stock_reports'
    output_path = os.path.join(base_path, 'dead_stock_report_modern.pdf')
    create_pdf_report(base_path, output_path)
    print(f"PDF created: {output_path}")
