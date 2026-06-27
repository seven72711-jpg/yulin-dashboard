#!/usr/bin/env python3
"""
一键更新：从新报表提取数据 → 更新 MD 文件
用法：python update_from_report.py "01_财务报表/原始数据/XXXX年X月月度报表玉林万达店.xlsx"
"""

import json
import sys
import os
from datetime import datetime
from extract_report import extract_all

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(PROJECT_DIR, "02_经营分析")


def load_md_data(filename, default=None):
    """Load JSON from MD frontmatter or return default"""
    path = os.path.join(DATA_DIR, filename)
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return default or []


def save_md_data(filename, data):
    """Save JSON to MD file"""
    path = os.path.join(DATA_DIR, filename)
    with open(path, 'w') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  ✓ {filename} ({len(data) if isinstance(data, list) else 'updated'})")


def update_all(report_path):
    """Extract new report and update all MD data files"""
    data = extract_all(report_path)
    new_month = data['dashboard']['period']
    
    print(f"\n📊 处理 {os.path.basename(report_path)} → {new_month}")
    print("-" * 50)
    
    # 1. 月度经营数据 — append or update the month
    monthly_db = load_md_data("月度经营数据.json", [])
    existing_periods = {m['period'] for m in monthly_db}
    
    for m in data['monthly']:
        if m['period'] in existing_periods:
            # Update existing
            for i, existing in enumerate(monthly_db):
                if existing['period'] == m['period']:
                    monthly_db[i] = m
                    break
        else:
            monthly_db.append(m)
    
    # Sort by period
    monthly_db.sort(key=lambda x: x['period'])
    save_md_data("月度经营数据.json", monthly_db)
    
    # 2. 储值金追踪 — append new row
    sv_tracker = load_md_data("储值金追踪.json", [])
    sv = data.get('water_bill', {}).get('stored_value', {})
    sv_entry = {
        'period': new_month,
        '充值': sv.get('inflow_充值', 0),
        '开卡': sv.get('inflow_开卡', 0),
        '套餐': sv.get('inflow_套餐', 0),
        '流入合计': sv.get('inflow_total', 0),
        '卡金消费': sv.get('outflow_消费', 0),
        '净增': sv.get('net_change', 0),
    }
    
    # Update or append
    updated = False
    for i, row in enumerate(sv_tracker):
        if row['period'] == new_month:
            sv_tracker[i] = sv_entry
            updated = True
            break
    if not updated:
        sv_tracker.append(sv_entry)
    sv_tracker.sort(key=lambda x: x['period'])
    save_md_data("储值金追踪.json", sv_tracker)
    
    # 3. 渠道分析 — per month
    ch_tracker = load_md_data("渠道分析.json", [])
    channels = data.get('water_bill', {}).get('channels', {})
    ch_entry = {'period': new_month}
    for ch, info in channels.items():
        ch_entry[ch] = {
            'orders': info['orders'],
            'revenue': info['received'],
            'discount': info['discount_rate'],
            'share': info['share'],
        }
    
    updated = False
    for i, row in enumerate(ch_tracker):
        if row['period'] == new_month:
            ch_tracker[i] = ch_entry
            updated = True
            break
    if not updated:
        ch_tracker.append(ch_entry)
    ch_tracker.sort(key=lambda x: x['period'])
    save_md_data("渠道分析.json", ch_tracker)
    
    # 4. 核心指标看板 — latest snapshot
    dashboard = data['dashboard']
    dashboard['updated_at'] = datetime.now().isoformat()
    dashboard['total_months'] = len(monthly_db)
    
    # Add trend comparison
    if len(monthly_db) >= 3:
        last3 = monthly_db[-3:]
        dashboard['avg_revenue_3m'] = round(sum(m['revenue'] for m in last3) / 3, 2)
        dashboard['avg_profit_3m'] = round(sum(m['net_profit'] for m in last3) / 3, 2)
        
        # YoY comparison if available
        year, month_str = new_month.split('-')
        prev_year = str(int(year) - 1)
        prev_period = f"{prev_year}-{month_str}"
        for m in monthly_db:
            if m['period'] == prev_period:
                dashboard['yoy_revenue'] = m['revenue']
                dashboard['yoy_profit'] = m['net_profit']
                dashboard['yoy_revenue_change'] = round((dashboard['revenue'] - m['revenue']) / m['revenue'] * 100, 1)
                break
    
    save_md_data("核心指标看板.json", dashboard)
    
    # 5. 总部结算追踪
    settlement = data.get('settlement')
    if settlement:
        settle_tracker = load_md_data("总部结算追踪.json", [])
        settle_entry = {
            'period': new_month,
            'deductions': settlement['deductions'],
            'total_deductions': settlement['total_deductions'],
            'net_settlement': settlement['net_settlement'],
        }
        updated = False
        for i, row in enumerate(settle_tracker):
            if row['period'] == new_month:
                settle_tracker[i] = settle_entry
                updated = True
                break
        if not updated:
            settle_tracker.append(settle_entry)
        settle_tracker.sort(key=lambda x: x['period'])
        save_md_data("总部结算追踪.json", settle_tracker)
    
    print()
    print(f"✅ 完成。{len(monthly_db)}个月数据已同步。")
    print(f"   核心指标看板: revenue={dashboard['revenue']:,.0f} profit={dashboard['net_profit']:,.0f} margin={dashboard['net_margin']}%")
    
    return dashboard


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python update_from_report.py <report.xlsx>")
        sys.exit(1)
    
    update_all(sys.argv[1])
