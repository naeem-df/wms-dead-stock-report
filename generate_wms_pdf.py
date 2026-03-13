#!/usr/bin/env python3
"""Generate WMS GROUP branded Dead Stock PDF Report"""

import csv
import os
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm, mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, 
    PageBreak, KeepTogether
)
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Use built-in fonts: Times (serif) and Helvetica (sans-serif)
# Playfair Display → Times-Bold for headings
# DM Sans → Helvetica for body

# WMS GROUP Brand Colors
WMS_BLACK = colors.HexColor('#111111')
WMS_YELLOW = colors.HexColor('#F5C800')
WMS_WHITE = colors.HexColor('#FFFFFF')
WMS_OFFWHITE = colors.HexColor('#F7F5F2')
WMS_MUTED = colors.HexColor('#888888')
WMS_BORDER = colors.HexColor('#E0E0E0')

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
        csv_path = os.path.join(base_path, f'dead_stock_{branch}_v2.csv')
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

def create_bar_chart(data_dict, width=450, height=150):
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
    bc.strokeColor = None
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = max(values) * 1.15 if values else 100
    bc.valueAxis.strokeColor = WMS_BORDER
    bc.valueAxis.labels.fontName = 'Helvetica'
    bc.valueAxis.labels.fontSize = 9
    bc.valueAxis.labels.fillColor = WMS_MUTED
    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.dx = 5
    bc.categoryAxis.labels.dy = -2
    bc.categoryAxis.labels.angle = 45
    bc.categoryAxis.labels.fontName = 'Helvetica'
    bc.categoryAxis.labels.fontSize = 9
    bc.categoryAxis.labels.fillColor = WMS_BLACK
    bc.categoryAxis.categoryNames = [l[:3].upper() for l in labels]
    bc.bars[0].fillColor = WMS_BLACK
    
    drawing.add(bc)
    return drawing

def create_pdf_report(base_path, output_path):
    """Create WMS GROUP branded PDF report"""
    
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
        rightMargin=1*inch,
        leftMargin=1*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch
    )
    
    # WMS GROUP Styles
    display_style = ParagraphStyle(
        'Display',
        fontSize=44,
        textColor=WMS_WHITE,
        spaceAfter=0,
        alignment=TA_CENTER,
        fontName='Times-Bold',
        leading=48
    )
    
    h1_style = ParagraphStyle(
        'H1',
        fontSize=28,
        textColor=WMS_BLACK,
        spaceBefore=0,
        spaceAfter=8,
        fontName='Times-Bold'
    )
    
    h2_style = ParagraphStyle(
        'H2',
        fontSize=20,
        textColor=WMS_BLACK,
        spaceBefore=12,
        spaceAfter=6,
        fontName='Times-Bold'
    )
    
    body_style = ParagraphStyle(
        'Body',
        fontSize=11,
        textColor=WMS_BLACK,
        spaceAfter=8,
        fontName='Helvetica',
        leading=14
    )
    
    label_style = ParagraphStyle(
        'Label',
        fontSize=10,
        textColor=WMS_YELLOW,
        fontName='Helvetica',
        alignment=TA_CENTER,
        spaceAfter=4
    )
    
    caption_style = ParagraphStyle(
        'Caption',
        fontSize=10,
        textColor=WMS_MUTED,
        fontName='Helvetica',
        spaceAfter=4
    )
    
    story = []
    
    # Metrics cards
    metrics_data = [
        [f"{total_parts:,}", f"{total_qty:,}", f"R{total_cost:,.0f}", f"{len(all_data)}"],
        ['PARTS', 'UNITS', 'COST VALUE', 'BRANCHES']
    ]
    
    metrics_table = Table(metrics_data, colWidths=[90, 90, 90, 90])
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), WMS_YELLOW),
        ('TEXTCOLOR', (0, 0), (-1, 0), WMS_BLACK),
        ('FONTNAME', (0, 0), (-1, 0), 'Times-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 22),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('BACKGROUND', (0, 1), (-1, 1), WMS_BLACK),
        ('TEXTCOLOR', (0, 1), (-1, 1), WMS_WHITE),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, 1), 9),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
    ]))
    
    # ========== TITLE SLIDE (Black background) ==========
    # Cover page as table with black background
    cover_table_data = [
        [Spacer(1, 100)],
        [Paragraph("DEAD STOCK ANALYSIS", display_style)],
        [Paragraph("WMS GROUP", label_style)],
        [Spacer(1, 32)],
        [Paragraph("PARTS INCORPORATED  •  SUPPLIER CODE AZ", 
                  ParagraphStyle('Sub', fontSize=14, textColor=WMS_MUTED, fontName='Helvetica', alignment=TA_CENTER))],
        [Spacer(1, 48)],
        [metrics_table],
        [Spacer(1, 32)],
        [Paragraph(f"Generated {datetime.now().strftime('%Y-%m-%d')}", 
                  ParagraphStyle('Date', fontSize=10, textColor=WMS_MUTED, fontName='Helvetica', alignment=TA_CENTER))],
        [Paragraph("Stock on hand with no sale since before 2024-01-01", 
                  ParagraphStyle('Def', fontSize=10, textColor=WMS_MUTED, fontName='Helvetica', alignment=TA_CENTER))],
    ]
    
    cover_table = Table(cover_table_data, colWidths=[6.5*inch])
    cover_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), WMS_BLACK),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(cover_table)
    story.append(PageBreak())
    
    # ========== EXECUTIVE SUMMARY (White background) ==========
    story.append(Paragraph("Executive Summary", h1_style))
    
    # Yellow accent bar
    bar_table = Table([['']], colWidths=[32], rowHeights=[3])
    bar_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), WMS_YELLOW)]))
    story.append(bar_table)
    story.append(Spacer(1, 12))
    
    # Branch summary table
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
    
    # Total row
    branch_summary.append(['TOTAL', f"{total_parts:,}", f"{total_qty:,}", f"R{total_cost:,.2f}", "100.0%"])
    
    branch_table = Table(branch_summary, colWidths=[85, 55, 70, 85, 55])
    branch_table.setStyle(TableStyle([
        # Header
        ('BACKGROUND', (0, 0), (-1, 0), WMS_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WMS_WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        
        # Body
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('TEXTCOLOR', (0, 1), (-1, -1), WMS_BLACK),
        
        # Alternating rows
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WMS_WHITE, WMS_OFFWHITE]),
        
        # Total row (yellow)
        ('BACKGROUND', (0, -1), (-1, -1), WMS_YELLOW),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica'),
        ('TEXTCOLOR', (0, -1), (-1, -1), WMS_BLACK),
        ('LINEABOVE', (0, -1), (-1, -1), 1, WMS_BLACK),
        
        # Borders - horizontal only
        ('LINEBELOW', (0, 0), (-1, 0), 1, WMS_BLACK),
        ('LINEBELOW', (0, 1), (-1, -2), 0.5, WMS_BORDER),
        
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(branch_table)
    
    # Bar chart
    story.append(Spacer(1, 16))
    branch_costs = {b[:3].upper(): calculate_branch_totals(all_data[b])[0]/1000 
                   for b in ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']
                   if b in all_data}
    story.append(create_bar_chart(branch_costs))
    
    story.append(PageBreak())
    
    # ========== PRODUCT GROUP ANALYSIS ==========
    story.append(Paragraph("Product Group Analysis", h1_style))
    bar_table2 = Table([['']], colWidths=[32], rowHeights=[3])
    bar_table2.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), WMS_YELLOW)]))
    story.append(bar_table2)
    story.append(Spacer(1, 12))
    
    pg_summary = [['RANK', 'PGROUP', 'DESCRIPTION', 'PARTS', 'QTY', 'COST VALUE', '% TOTAL']]
    sorted_pg = sorted(pgroup_totals.items(), key=lambda x: x[1]['cost'], reverse=True)[:15]
    
    for i, (pg, data) in enumerate(sorted_pg, 1):
        pct = (data['cost'] / total_cost * 100) if total_cost > 0 else 0
        desc = PGROUP_MAP.get(pg, '—')[:22]
        pg_summary.append([
            str(i),
            pg,
            desc,
            f"{data['parts']:,}",
            f"{data['qty']:,}",
            f"R{data['cost']:,.2f}",
            f"{pct:.1f}%"
        ])
    
    pg_table = Table(pg_summary, colWidths=[30, 45, 95, 40, 40, 70, 40])
    pg_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), WMS_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WMS_WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
        ('TEXTCOLOR', (0, 1), (-1, -1), WMS_BLACK),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WMS_WHITE, WMS_OFFWHITE]),
        ('LINEBELOW', (0, 0), (-1, 0), 1, WMS_BLACK),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, WMS_BORDER),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(pg_table)
    
    story.append(PageBreak())
    
    # ========== BRANCH DETAILS ==========
    for branch in ['OLIFANTS', 'WADEVILLE', 'CHLOORKOP', 'DAXINA', 'TEMBISA', 'EVATON', 'DIEPSLOOT']:
        if branch not in all_data:
            continue
        
        branch_cost, branch_qty = calculate_branch_totals(all_data[branch])
        branch_pg_data = branch_pgroup_data.get(branch, {})
        
        # Branch header
        story.append(Paragraph(branch.title(), h1_style))
        bar_table3 = Table([['']], colWidths=[32], rowHeights=[3])
        bar_table3.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), WMS_YELLOW)]))
        story.append(bar_table3)
        story.append(Spacer(1, 8))
        
        story.append(Paragraph(f"{len(all_data[branch]):,} parts  •  {branch_qty:,} units  •  R{branch_cost:,.0f} cost value", caption_style))
        story.append(Spacer(1, 12))
        
        # Sort all parts by PGROUP, then by cost
        all_branch_parts = []
        for pg, parts in branch_pg_data.items():
            for part in parts:
                part['_pgroup'] = pg
                all_branch_parts.append(part)
        
        all_branch_parts.sort(key=lambda x: (x.get('_pgroup', ''), -float(x.get('cost_value', 0) or 0)))
        
        # Build table
        table_data = [['PART NO', 'DESCRIPTION', 'QTY', 'UNIT COST', 'TOTAL COST', 'ORDER REF', 'LAST SOLD']]
        
        current_pg = None
        pg_subtotal = 0
        pg_qty_subtotal = 0
        
        for row in all_branch_parts:
            pg = row.get('_pgroup', 'UNKNOWN')
            
            if pg != current_pg:
                if current_pg is not None:
                    pg_desc = PGROUP_MAP.get(current_pg, '—')
                    table_data.append(['', f'{current_pg} — {pg_desc} SUBTOTAL', f'{pg_qty_subtotal:,}', '', f'R{pg_subtotal:,.2f}', '', ''])
                    table_data.append(['', '', '', '', '', '', ''])
                
                pg_subtotal = 0
                pg_qty_subtotal = 0
                current_pg = pg
                
                pg_desc = PGROUP_MAP.get(pg, '—')
                table_data.append([f'▶ {pg}', pg_desc, '', '', '', '', ''])
            
            partno = row.get('partno', '')[:14]
            desc = row.get('sdesc', '')[:18]
            qoh = row.get('qoh', '0')
            lcost = float(row.get('lcost', 0) or 0)
            cost_val = float(row.get('cost_value', 0) or 0)
            order_ref = (row.get('last_order_ref') or '—')[:10]
            lsold = (row.get('lsold') or '—')[:10]
            
            table_data.append([partno, desc, qoh, f"R{lcost:.2f}", f"R{cost_val:,.2f}", order_ref, lsold])
            
            pg_subtotal += cost_val
            pg_qty_subtotal += int(row.get('qoh', 0) or 0)
        
        if current_pg is not None:
            pg_desc = PGROUP_MAP.get(current_pg, '—')
            table_data.append(['', f'{current_pg} — {pg_desc} SUBTOTAL', f'{pg_qty_subtotal:,}', '', f'R{pg_subtotal:,.2f}', '', ''])
        
        table_data.append(['', 'BRANCH TOTAL', f'{branch_qty:,}', '', f'R{branch_cost:,.2f}', '', ''])
        
        branch_table = Table(table_data, colWidths=[65, 85, 30, 52, 60, 52, 52])
        
        style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), WMS_BLACK),
            ('TEXTCOLOR', (0, 0), (-1, 0), WMS_WHITE),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
            ('ALIGN', (3, 0), (4, -1), 'RIGHT'),
            ('ALIGN', (5, 0), (6, -1), 'CENTER'),
            ('TEXTCOLOR', (0, 1), (-1, -1), WMS_BLACK),
            ('LINEBELOW', (0, 0), (-1, 0), 1, WMS_BLACK),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        
        row_idx = 0
        for i, row in enumerate(table_data[1:], 1):
            first_cell = row[0] if row else ''
            
            if first_cell.startswith('▶'):
                style_commands.append(('BACKGROUND', (0, i), (-1, i), WMS_OFFWHITE))
                style_commands.append(('FONTNAME', (0, i), (-1, i), 'Helvetica'))
                style_commands.append(('FONTSIZE', (0, i), (-1, i), 8))
                row_idx = 0
            elif 'SUBTOTAL' in str(row[1]):
                style_commands.append(('BACKGROUND', (0, i), (-1, i), WMS_OFFWHITE))
                style_commands.append(('FONTNAME', (0, i), (-1, i), 'Helvetica'))
                style_commands.append(('LINEABOVE', (0, i), (-1, i), 0.5, WMS_BORDER))
            elif all(cell == '' for cell in row):
                pass
            else:
                if row_idx % 2 == 0:
                    style_commands.append(('BACKGROUND', (0, i), (-1, i), WMS_OFFWHITE))
                row_idx += 1
        
        for i, row in enumerate(table_data):
            if 'BRANCH TOTAL' in str(row[1]):
                style_commands.append(('BACKGROUND', (0, i), (-1, i), WMS_YELLOW))
                style_commands.append(('FONTNAME', (0, i), (-1, i), 'Helvetica'))
                style_commands.append(('TEXTCOLOR', (0, i), (-1, i), WMS_BLACK))
                style_commands.append(('LINEABOVE', (0, i), (-1, i), 1, WMS_BLACK))
        
        branch_table.setStyle(TableStyle(style_commands))
        story.append(branch_table)
        
        story.append(PageBreak())
    
    doc.build(story)
    return output_path

if __name__ == '__main__':
    base_path = '/root/.openclaw/workspace/dead_stock_reports'
    output_path = os.path.join(base_path, 'dead_stock_report_wms.pdf')
    create_pdf_report(base_path, output_path)
    print(f"PDF created: {output_path}")
