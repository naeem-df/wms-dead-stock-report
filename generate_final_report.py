#!/usr/bin/env python3
"""Generate WMS GROUP Branded Dead Stock Report — Combined Styleguide + Formatting"""

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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# WMS GROUP Brand Colors
WMS_BLACK = colors.HexColor('#111111')
WMS_YELLOW = colors.HexColor('#F5C800')
WMS_WHITE = colors.HexColor('#FFFFFF')
WMS_OFFWHITE = colors.HexColor('#F7F5F2')
WMS_MUTED = colors.HexColor('#888888')
WMS_BORDER = colors.HexColor('#E0E0E0')

# PGROUP descriptions (Title Case)
PGROUP_MAP = {
    'A03': 'Balljoints', 'B01': 'Exhaust', 'B02': 'Ignition Leads',
    'B06': 'Brake Pads', 'B09': 'Brake Shoes', 'B19': 'Brake Parts',
    'BO': 'B/O Buy-Out', 'BP': 'Body Parts', 'BP2': 'BP Taillamps',
    'C05': 'Neadle & Seats', 'C06': 'Radiator Caps', 'C07': 'Thermostats',
    'C08': 'F/Belts (V Nos.)', 'C13': 'Clutch Kits', 'C14': 'Cables Speedo',
    'C18': 'CV Joints', 'D02': 'U Joint & Yokes', 'D03': 'Drums',
    'D06': 'Hub,Disc', 'D09': 'Distributor Caps',
    'E02': 'Echlin Points & Cond.', 'E04': 'Globes', 'E06': 'Flashers & Relays',
    'E08': 'Switches Temp.', 'E09': 'Oil Sender Units', 'E10': 'Switch Stop Light',
    'E11': 'Switches Rev Light', 'E12': 'Arm,Rect,& Regulators',
    'E14': 'Sta,& Alt Repair Kits', 'E19': 'Bendix Drives', 'E20': 'Solenoids',
    'E21': 'Startor & Alternator', 'E23': 'Fan Switch', 'E25': 'Coils', 'E28': 'Glow Plugs',
    'F01': 'Fuel Pumps', 'F02': 'Air Filters', 'F03': 'Front Counter',
    'F05': 'Carb Kits', 'F06': 'FS (Slow Movers)', 'F07': 'Fuel Filters',
    'F10': 'Oil & Fuel (Cartridge)', 'F11': 'Oil Filters',
    'FAI': 'FAI', 'FSD': 'Deans', 'FSK': 'Kwh',
    'FSP': 'Front Stock Pia', 'FSS': 'Shield', 'FSTI': 'Stickers', 'G01': 'Shocks',
    'H05': 'Hydraulic Complete Cylinders', 'M01': 'Mounting Eng & G/Box',
    'M04': 'F/Belt (Micro)', 'O03': 'Oil Seals', 'O06': 'Oil Caps',
    'P20': 'Hydraulic Kits', 'PROM': 'Promo', 'Q1': 'Quantum', 'R10': 'Radiator Caps',
    'RAL1': 'Rally', 'S06': 'Spark Plug (Champion)', 'S07': 'Spark Plugs (NGK)',
    'S08': 'Spark Plugs (Bosch)', 'S11': 'Tierods,Draglink',
    'T01': 'Spanners & Sockets', 'T02': 'Thrust Bearings', 'T05': 'Timing Chains',
    'T13': 'Timing Belts', 'T15': 'Timing Belt Tensioners', 'V02': 'V W Special Parts List',
    'W05': 'Wheel Bearing Kits', 'W07': 'Water Pumps', 'W09': 'Wiper Blades',
    'W16': 'Window Tint & Stripe'
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

def format_currency(value):
    """Format as R X,XXX.XX"""
    try:
        return f"R {float(value):,.2f}"
    except:
        return "R 0.00"

def format_date(date_str):
    """Format date or return NO HISTORY"""
    if not date_str or date_str == '' or date_str == 'None':
        return "NO HISTORY"
    try:
        return datetime.strptime(date_str[:10], '%Y-%m-%d').strftime('%Y-%m-%d')
    except:
        return "NO HISTORY"

def format_order_ref(ref):
    """Format order ref or return em dash"""
    if not ref or ref == '' or ref == 'NO HISTORY' or ref == 'None':
        return "—"
    return ref

def create_pdf_report(base_path, output_path):
    """Create WMS GROUP branded PDF with formatted content"""
    
    all_data = load_all_data(base_path)
    
    total_parts = sum(len(data) for data in all_data.values())
    total_cost = sum(sum(float(row.get('cost_value', 0) or 0) for row in data) for data in all_data.values())
    total_qty = sum(sum(int(row.get('qoh', 0) or 0) for row in data) for data in all_data.values())
    
    # Branch totals for sorting
    branch_totals = {}
    for branch, data in all_data.items():
        cost = sum(float(row.get('cost_value', 0) or 0) for row in data)
        qty = sum(int(row.get('qoh', 0) or 0) for row in data)
        branch_totals[branch] = {'cost': cost, 'qty': qty, 'parts': len(data)}
    
    # Product group totals
    pgroup_totals = {}
    for branch, data in all_data.items():
        for row in data:
            pg = row.get('pgroup', 'UNKNOWN').upper()
            if pg not in pgroup_totals:
                pgroup_totals[pg] = {'parts': 0, 'qty': 0, 'cost': 0}
            pgroup_totals[pg]['parts'] += 1
            pgroup_totals[pg]['qty'] += int(row.get('qoh', 0) or 0)
            pgroup_totals[pg]['cost'] += float(row.get('cost_value', 0) or 0)
    
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=1*inch,
        leftMargin=1*inch,
        topMargin=0.75*inch,
        bottomMargin=0.75*inch
    )
    
    # Styles (ALL CAPS headings, Title Case PGROUPs)
    h1_caps = ParagraphStyle(
        'H1Caps',
        fontSize=28,
        textColor=WMS_WHITE,
        alignment=TA_CENTER,
        fontName='Times-Bold',
        spaceAfter=0
    )
    
    h2_caps = ParagraphStyle(
        'H2Caps',
        fontSize=18,
        textColor=WMS_BLACK,
        alignment=TA_CENTER,
        fontName='Times-Bold',
        spaceAfter=12,
        spaceBefore=20
    )
    
    body = ParagraphStyle(
        'Body',
        fontSize=10,
        textColor=WMS_BLACK,
        fontName='Helvetica',
        leading=14
    )
    
    label_caps = ParagraphStyle(
        'LabelCaps',
        fontSize=9,
        textColor=WMS_YELLOW,
        fontName='Helvetica-Bold',
        alignment=TA_CENTER
    )
    
    caption = ParagraphStyle(
        'Caption',
        fontSize=9,
        textColor=WMS_MUTED,
        fontName='Helvetica',
        alignment=TA_CENTER
    )
    
    story = []
    
    # ========== COVER PAGE (Black background) ==========
    cover_data = [
        [Spacer(1, 80)],
        [Paragraph("DEAD STOCK ANALYSIS", h1_caps)],
        [Paragraph("WADEVILLE MOTOR SPARES GROUP", label_caps)],
        [Spacer(1, 20)],
        [Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d')}", caption)],
        [Paragraph("Stock on hand with no sale since before 2024-01-01", caption)],
        [Spacer(1, 40)],
    ]
    
    # Metrics in yellow
    metrics = [
        [f"{total_parts:,}", f"{total_qty:,}", f"R {total_cost:,.0f}", f"{len(all_data)}"],
        ['PARTS', 'UNITS', 'COST VALUE', 'BRANCHES']
    ]
    
    metrics_table = Table(metrics, colWidths=[90, 90, 90, 90])
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), WMS_YELLOW),
        ('TEXTCOLOR', (0, 0), (-1, 0), WMS_BLACK),
        ('FONTNAME', (0, 0), (-1, 0), 'Times-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 20),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('BACKGROUND', (0, 1), (-1, 1), WMS_BLACK),
        ('TEXTCOLOR', (0, 1), (-1, 1), WMS_WHITE),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 8),
        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
    ]))
    cover_data.append([metrics_table])
    cover_data.append([Spacer(1, 60)])
    cover_data.append([Paragraph("CONFIDENTIAL — INTERNAL USE ONLY", caption)])
    
    cover_table = Table(cover_data, colWidths=[6.5*inch])
    cover_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), WMS_BLACK),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    story.append(cover_table)
    story.append(PageBreak())
    
    # ========== EXECUTIVE SUMMARY (ALL CAPS heading) ==========
    story.append(Paragraph("EXECUTIVE SUMMARY", h2_caps))
    
    # Yellow bar
    bar = Table([['']], colWidths=[32], rowHeights=[3])
    bar.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), WMS_YELLOW)]))
    story.append(bar)
    story.append(Spacer(1, 12))
    
    # Branch table (sorted by cost descending)
    sorted_branches = sorted(branch_totals.items(), key=lambda x: x[1]['cost'], reverse=True)
    
    branch_data = [['BRANCH', 'PARTS', 'UNITS', 'COST VALUE', '% TOTAL']]
    for branch, vals in sorted_branches:
        pct = (vals['cost'] / total_cost * 100) if total_cost > 0 else 0
        branch_data.append([
            branch.upper(),
            f"{vals['parts']:,}",
            f"{vals['qty']:,}",
            format_currency(vals['cost']),
            f"{pct:.1f}%"
        ])
    branch_data.append(['TOTAL', f"{total_parts:,}", f"{total_qty:,}", format_currency(total_cost), "100.0%"])
    
    branch_table = Table(branch_data, colWidths=[90, 60, 60, 100, 60])
    branch_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), WMS_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WMS_WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('TEXTCOLOR', (0, 1), (-1, -1), WMS_BLACK),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WMS_WHITE, WMS_OFFWHITE]),
        ('BACKGROUND', (0, -1), (-1, -1), WMS_YELLOW),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('LINEBELOW', (0, 0), (-1, 0), 1, WMS_BLACK),
        ('LINEBELOW', (0, 1), (-1, -2), 0.5, WMS_BORDER),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(branch_table)
    story.append(PageBreak())
    
    # ========== PRODUCT GROUP ANALYSIS ==========
    story.append(Paragraph("PRODUCT GROUP ANALYSIS", h2_caps))
    bar2 = Table([['']], colWidths=[32], rowHeights=[3])
    bar2.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), WMS_YELLOW)]))
    story.append(bar2)
    story.append(Spacer(1, 12))
    
    sorted_pg = sorted(pgroup_totals.items(), key=lambda x: x[1]['cost'], reverse=True)
    
    pg_data = [['RANK', 'PGROUP', 'DESCRIPTION', 'PARTS', 'UNITS', 'COST VALUE', '% TOTAL']]
    for i, (pg, vals) in enumerate(sorted_pg[:20], 1):
        pct = (vals['cost'] / total_cost * 100) if total_cost > 0 else 0
        desc = PGROUP_MAP.get(pg, '—')[:22]
        pg_data.append([
            str(i),
            pg,
            desc,
            f"{vals['parts']:,}",
            f"{vals['qty']:,}",
            format_currency(vals['cost']),
            f"{pct:.1f}%"
        ])
    
    pg_table = Table(pg_data, colWidths=[30, 45, 100, 40, 40, 85, 50])
    pg_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), WMS_BLACK),
        ('TEXTCOLOR', (0, 0), (-1, 0), WMS_WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 7),
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
    
    # ========== BRANCH DETAILS (sorted by cost descending) ==========
    for branch, vals in sorted_branches:
        branch_data = all_data[branch]
        
        # Branch heading (ALL CAPS)
        story.append(Paragraph(branch.upper(), h2_caps))
        bar3 = Table([['']], colWidths=[32], rowHeights=[3])
        bar3.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, -1), WMS_YELLOW)]))
        story.append(bar3)
        story.append(Spacer(1, 8))
        
        story.append(Paragraph(f"{vals['parts']:,} parts  |  {vals['qty']:,} units  |  {format_currency(vals['cost'])} cost value", caption))
        story.append(Spacer(1, 12))
        
        # Group by PGROUP
        pg_data_branch = {}
        for row in branch_data:
            pg = row.get('pgroup', 'UNKNOWN').upper()
            if pg not in pg_data_branch:
                pg_data_branch[pg] = []
            pg_data_branch[pg].append(row)
        
        # Sort PGROUPs by cost
        sorted_branch_pg = sorted(pg_data_branch.items(), 
                                  key=lambda x: sum(float(r.get('cost_value', 0) or 0) for r in x[1]), 
                                  reverse=True)
        
        # Build single table
        table_data = [['PART NO', 'DESCRIPTION', 'QTY', 'UNIT COST', 'TOTAL COST', 'ORDER REF', 'LAST SOLD']]
        
        current_pg = None
        pg_subtotal = 0
        pg_qty = 0
        
        for pg, parts in sorted_branch_pg:
            pg_desc = PGROUP_MAP.get(pg, '—')
            pg_total = sum(float(r.get('cost_value', 0) or 0) for r in parts)
            pg_qty_total = sum(int(r.get('qoh', 0) or 0) for r in parts)
            
            # Sort parts by cost
            sorted_parts = sorted(parts, key=lambda x: float(x.get('cost_value', 0) or 0), reverse=True)
            
            for row in sorted_parts:
                partno = (row.get('partno') or '')[:14]
                desc = (row.get('sdesc') or '')[:18]
                qty = int(row.get('qoh') or 0)
                unit_cost = float(row.get('lcost') or 0)
                total_cost_part = float(row.get('cost_value') or 0)
                order_ref = format_order_ref(row.get('last_order_ref'))
                last_sold = format_date(row.get('lsold'))
                
                table_data.append([
                    partno,
                    desc,
                    f"{qty:,}",
                    format_currency(unit_cost),
                    format_currency(total_cost_part),
                    order_ref,
                    last_sold
                ])
                
                pg_subtotal += total_cost_part
                pg_qty += qty
            
            # Subtotal row
            table_data.append(['', f"{pg} — {pg_desc} Subtotal", f"{pg_qty_total:,}", '', format_currency(pg_total), '', ''])
            
            pg_subtotal = 0
            pg_qty = 0
        
        # Branch total
        table_data.append(['', 'BRANCH TOTAL', f"{vals['qty']:,}", '', format_currency(vals['cost']), '', ''])
        
        branch_table = Table(table_data, colWidths=[65, 90, 35, 65, 70, 55, 55])
        
        style_cmds = [
            ('BACKGROUND', (0, 0), (-1, 0), WMS_BLACK),
            ('TEXTCOLOR', (0, 0), (-1, 0), WMS_WHITE),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
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
        
        # Style subtotal and total rows
        for i, row in enumerate(table_data):
            if 'Subtotal' in str(row[1]):
                style_cmds.append(('BACKGROUND', (0, i), (-1, i), WMS_OFFWHITE))
                style_cmds.append(('FONTNAME', (0, i), (-1, i), 'Helvetica-Bold'))
                style_cmds.append(('LINEABOVE', (0, i), (-1, i), 0.5, WMS_BORDER))
            elif 'BRANCH TOTAL' in str(row[1]):
                style_cmds.append(('BACKGROUND', (0, i), (-1, i), WMS_YELLOW))
                style_cmds.append(('FONTNAME', (0, i), (-1, i), 'Helvetica-Bold'))
                style_cmds.append(('LINEABOVE', (0, i), (-1, i), 1, WMS_BLACK))
        
        branch_table.setStyle(TableStyle(style_cmds))
        story.append(branch_table)
        story.append(PageBreak())
    
    # Footer
    story.append(Paragraph("CONFIDENTIAL — WADEVILLE MOTOR SPARES GROUP — INTERNAL USE ONLY", caption))
    
    doc.build(story)
    print(f"PDF created: {output_path}")
    return output_path

if __name__ == '__main__':
    base_path = '/root/.openclaw/workspace/dead_stock_reports'
    output_path = os.path.join(base_path, 'dead_stock_report_final.pdf')
    create_pdf_report(base_path, output_path)
