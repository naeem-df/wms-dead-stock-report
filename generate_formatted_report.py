#!/usr/bin/env python3
"""Generate formatted dead stock report"""

import csv
import os
from datetime import datetime
import psycopg2

# PGROUP descriptions
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

def generate_report(base_path, output_path):
    """Generate formatted report"""
    
    all_data = load_all_data(base_path)
    
    # Calculate totals
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
    
    # Build output
    lines = []
    
    # Document title block
    lines.append("DEAD STOCK ANALYSIS".center(72))
    lines.append("WADEVILLE MOTOR SPARES GROUP".center(72))
    lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d')}".center(72))
    lines.append("Stock on hand with no sale since before 2024-01-01".center(72))
    lines.append("")
    lines.append("")
    
    # Confidentiality notice
    lines.append("CONFIDENTIAL — WADEVILLE MOTOR SPARES GROUP — INTERNAL USE ONLY".center(72))
    lines.append("")
    lines.append("")
    
    # Executive summary table
    lines.append("EXECUTIVE SUMMARY".center(72))
    lines.append("")
    lines.append("BRANCH".ljust(15) + "|" + "PARTS".rjust(10) + "|" + "UNITS".rjust(10) + "|" + "COST VALUE".rjust(15) + "|" + "% TOTAL".rjust(10))
    lines.append("-" * 15 + "+" + "-" * 10 + "+" + "-" * 10 + "+" + "-" * 15 + "+" + "-" * 10)
    
    sorted_branches = sorted(branch_totals.items(), key=lambda x: x[1]['cost'], reverse=True)
    for branch, vals in sorted_branches:
        pct = (vals['cost'] / total_cost * 100) if total_cost > 0 else 0
        lines.append(
            branch.title().ljust(15) + "|" +
            f"{vals['parts']:,}".rjust(10) + "|" +
            f"{vals['qty']:,}".rjust(10) + "|" +
            format_currency(vals['cost']).rjust(15) + "|" +
            f"{pct:.1f}%".rjust(10)
        )
    
    lines.append("-" * 15 + "+" + "-" * 10 + "+" + "-" * 10 + "+" + "-" * 15 + "+" + "-" * 10)
    lines.append(
        "TOTAL".ljust(15) + "|" +
        f"{total_parts:,}".rjust(10) + "|" +
        f"{total_qty:,}".rjust(10) + "|" +
        format_currency(total_cost).rjust(15) + "|" +
        "100.0%".rjust(10)
    )
    lines.append("")
    lines.append("")
    
    # Product group analysis
    lines.append("PRODUCT GROUP ANALYSIS".center(72))
    lines.append("")
    lines.append("RANK".ljust(6) + "|" + "PGROUP".ljust(8) + "|" + "DESCRIPTION".ljust(25) + "|" + "PARTS".rjust(8) + "|" + "UNITS".rjust(8) + "|" + "COST VALUE".rjust(14) + "|" + "% TOTAL".rjust(9))
    lines.append("-" * 6 + "+" + "-" * 8 + "+" + "-" * 25 + "+" + "-" * 8 + "+" + "-" * 8 + "+" + "-" * 14 + "+" + "-" * 9)
    
    sorted_pg = sorted(pgroup_totals.items(), key=lambda x: x[1]['cost'], reverse=True)
    for i, (pg, vals) in enumerate(sorted_pg[:20], 1):
        pct = (vals['cost'] / total_cost * 100) if total_cost > 0 else 0
        desc = PGROUP_MAP.get(pg, '—')[:23]
        lines.append(
            str(i).ljust(6) + "|" +
            pg.ljust(8) + "|" +
            desc.ljust(25) + "|" +
            f"{vals['parts']:,}".rjust(8) + "|" +
            f"{vals['qty']:,}".rjust(8) + "|" +
            format_currency(vals['cost']).rjust(14) + "|" +
            f"{pct:.1f}%".rjust(9)
        )
    
    lines.append("")
    lines.append("")
    
    # Branch details
    for branch, vals in sorted_branches:
        branch_data = all_data[branch]
        
        lines.append("=" * 72)
        lines.append(branch.upper().center(72))
        lines.append(f"{vals['parts']:,} parts  |  {vals['qty']:,} units  |  {format_currency(vals['cost'])} cost value".center(72))
        lines.append("=" * 72)
        lines.append("")
        
        # Group by PGROUP
        pg_data = {}
        for row in branch_data:
            pg = row.get('pgroup', 'UNKNOWN').upper()
            if pg not in pg_data:
                pg_data[pg] = []
            pg_data[pg].append(row)
        
        # Sort PGROUPs by cost
        sorted_branch_pg = sorted(pg_data.items(), 
                                  key=lambda x: sum(float(r.get('cost_value', 0) or 0) for r in x[1]), 
                                  reverse=True)
        
        for pg, parts in sorted_branch_pg:
            pg_desc = PGROUP_MAP.get(pg, '—')
            pg_total = sum(float(r.get('cost_value', 0) or 0) for r in parts)
            pg_qty = sum(int(r.get('qoh', 0) or 0) for r in parts)
            
            lines.append(f"{pg} — {pg_desc}")
            lines.append("-" * 72)
            
            # Table header
            lines.append(
                "PART NO".ljust(18) + "|" +
                "DESCRIPTION".ljust(22) + "|" +
                "QTY".rjust(5) + "|" +
                "UNIT COST".rjust(12) + "|" +
                "TOTAL COST".rjust(13) + "|" +
                "ORDER REF".rjust(12) + "|" +
                "LAST SOLD".rjust(12)
            )
            lines.append("-" * 18 + "+" + "-" * 22 + "+" + "-" * 5 + "+" + "-" * 12 + "+" + "-" * 13 + "+" + "-" * 12 + "+" + "-" * 12)
            
            # Sort parts by cost
            sorted_parts = sorted(parts, key=lambda x: float(x.get('cost_value', 0) or 0), reverse=True)
            
            for row in sorted_parts:
                partno = (row.get('partno') or '')[:16]
                desc = (row.get('sdesc') or '')[:20]
                qty = int(row.get('qoh') or 0)
                unit_cost = float(row.get('lcost') or 0)
                total_cost_part = float(row.get('cost_value') or 0)
                order_ref = format_order_ref(row.get('last_order_ref'))
                last_sold = format_date(row.get('lsold'))
                
                lines.append(
                    partno.ljust(18) + "|" +
                    desc.ljust(22) + "|" +
                    f"{qty:,}".rjust(5) + "|" +
                    format_currency(unit_cost).rjust(12) + "|" +
                    format_currency(total_cost_part).rjust(13) + "|" +
                    order_ref.rjust(12) + "|" +
                    last_sold.rjust(12)
                )
            
            # Subtotal
            lines.append("-" * 72)
            lines.append(
                "  " + f"{pg} SUBTOTAL".ljust(41) + "|" +
                f"{pg_qty:,}".rjust(5) + "|" +
                "".rjust(12) + "|" +
                format_currency(pg_total).rjust(13) + "|" +
                "".rjust(12) + "|" +
                "".rjust(12)
            )
            lines.append("")
        
        # Branch total
        lines.append("-" * 72)
        lines.append(
            "BRANCH TOTAL".ljust(41) + "|" +
            f"{vals['qty']:,}".rjust(5) + "|" +
            "".rjust(12) + "|" +
            format_currency(vals['cost']).rjust(13) + "|" +
            "".rjust(12) + "|" +
            "".rjust(12)
        )
        lines.append("")
        lines.append("")
    
    # Footer
    lines.append("")
    lines.append("CONFIDENTIAL — WADEVILLE MOTOR SPARES GROUP — INTERNAL USE ONLY".center(72))
    
    # Write to file
    with open(output_path, 'w') as f:
        f.write('\n'.join(lines))
    
    print(f"Report generated: {output_path}")
    return output_path

if __name__ == '__main__':
    base_path = '/root/.openclaw/workspace/dead_stock_reports'
    output_path = os.path.join(base_path, 'dead_stock_report.txt')
    generate_report(base_path, output_path)
