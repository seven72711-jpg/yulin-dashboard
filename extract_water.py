#!/usr/bin/env python3
"""提取水单明细数据 — 时段/渠道/价格带/技师人效/翻台率

列定位改为「表头名映射」，报表新增列时不再错位。
"""
import openpyxl, json, sys, os, calendar
from collections import Counter, defaultdict
from datetime import datetime, date

# 需要的列（按表头名定位）
COLS = {
    'order_type': '订单类型',
    'time': '消费时间',
    'rev_orig': '原价业绩',
    'rev_recv': '实收金额',
    'pay': '支付明细',
    'tech': '服务员工',
    'designated': '指定',
    'project': '消费项目',
    'new_customer': '是否新客',
}


def find_water_sheet(wb):
    for s in wb.sheetnames:
        if '水单' in s:
            return wb[s]
    return None


def header_map(ws):
    """返回 {逻辑名: 列索引(0-based)}，缺列直接报错。"""
    hdr = None
    for r in ws.iter_rows(min_row=1, max_row=3, values_only=True):
        if r and any(v and '水单号' in str(v) for v in r):
            hdr = list(r)
            break
    if hdr is None:
        raise RuntimeError("未找到表头行（含水单号）")
    name2idx = {}
    for i, v in enumerate(hdr):
        if v is not None:
            name2idx[str(v).strip()] = i
    missing = [v for v in COLS.values() if v not in name2idx]
    if missing:
        raise RuntimeError(f"水单表缺列: {missing}；实际列={list(name2idx)}")
    return {k: name2idx[v] for k, v in COLS.items()}, hdr


def num(v):
    """安全转数字：空/文本/异常都返回 0.0"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '').replace('¥', '')
    try:
        return float(s)
    except ValueError:
        return 0.0


def extract_water_bill(report_path):
    wb = openpyxl.load_workbook(report_path, read_only=True, data_only=True)
    ws = find_water_sheet(wb)
    if not ws:
        print("No 水单 sheet found")
        return None
    ix, hdr = header_map(ws)
    print(f"水单表列数: {len(hdr)} | 关键列位置: " +
          ", ".join(f"{v}={ix[k]}" for k, v in COLS.items()))

    hourly, dow, bands = Counter(), Counter(), Counter()
    channels = Counter()
    channel_rev, channel_gap = defaultdict(float), defaultdict(float)
    tech_orders, tech_rev = Counter(), defaultdict(float)
    tech_desig, tech_private = Counter(), Counter()
    total_orders = total_received = total_discount = 0
    new_cust = 0
    private_orders = 0
    skipped = 0

    def cell(r, key):
        return r[ix[key]] if ix[key] < len(r) else None

    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r:
            continue
        ot = str(cell(r, 'order_type') or '')
        if '服务' not in ot:
            continue
        t = cell(r, 'time')
        if isinstance(t, datetime):
            h, d = t.hour, t.weekday()
        elif isinstance(t, str):
            try:
                dt = datetime.strptime(t[:19], '%Y-%m-%d %H:%M:%S')
                h, d = dt.hour, dt.weekday()
            except Exception:
                skipped += 1
                continue
        else:
            skipped += 1
            continue

        rev_orig = num(cell(r, 'rev_orig'))
        rev_recv = num(cell(r, 'rev_recv'))
        pay = str(cell(r, 'pay') or '')
        tech = str(cell(r, 'tech') or '').strip()
        proj = str(cell(r, 'project') or '')
        desig = str(cell(r, 'designated') or '').strip() == '是'
        is_new = str(cell(r, 'new_customer') or '').strip() == '是'

        hourly[h] += 1; dow[d] += 1; total_orders += 1
        total_received += rev_recv; total_discount += (rev_orig - rev_recv)
        tech_orders[tech] += 1; tech_rev[tech] += rev_orig
        if desig: tech_desig[tech] += 1
        if '私人定制' in proj:
            tech_private[tech] += 1; private_orders += 1
        if is_new: new_cust += 1

        if rev_recv < 80: bands['<¥80'] += 1
        elif rev_recv < 120: bands['¥80-120'] += 1
        elif rev_recv < 160: bands['¥120-160'] += 1
        elif rev_recv < 200: bands['¥160-200'] += 1
        elif rev_recv < 260: bands['¥200-260'] += 1
        else: bands['¥260+'] += 1

        if '美团' in pay: k = '美团'
        elif '抖音' in pay: k = '抖音'
        elif '会员卡' in pay: k = '储值卡'
        elif any(x in pay for x in ['现金', '微信', '支付宝', '扫', '转账']): k = '直付'
        else: k = '其他'
        channels[k] += 1; channel_rev[k] += rev_recv; channel_gap[k] += (rev_orig - rev_recv)
    wb.close()

    if skipped:
        print(f"⚠ 跳过 {skipped} 行（时间字段无法解析）")

    try:
        parts = os.path.basename(report_path).replace('.xlsx', '').split('年')
        yr = int(parts[0][-4:])
        mo = int(parts[1].split('月')[0])
    except Exception:
        raise RuntimeError(f"文件名无法解析年月: {report_path}")
    days_in_month = calendar.monthrange(yr, mo)[1]
    beds = 17
    daily = total_orders / days_in_month
    wd = sum(dow[d] for d in range(5)); we = sum(dow[d] for d in range(5, 7))
    wd_days = len([d for d in range(1, days_in_month + 1) if date(yr, mo, d).weekday() < 5])
    we_days = days_in_month - wd_days

    return {
        'period': f"{yr}-{mo:02d}",
        'total_orders': total_orders,
        'total_received': round(total_received, 2),
        'total_discount': round(total_discount, 2),
        'daily_avg': round(daily, 1),
        'per_bed': round(daily / beds, 2),
        'wkday_avg': round(wd / wd_days) if wd_days else 0,
        'wkend_avg': round(we / we_days) if we_days else 0,
        'hourly_pct': [round(hourly.get(h, 0) / total_orders * 100, 1) if total_orders else 0 for h in range(24)],
        'day_orders': [dow[d] for d in range(7)],
        'price_bands': {b: {'orders': bands[b], 'pct': round(bands[b] / total_orders * 100, 1)} for b in ['<¥80', '¥80-120', '¥120-160', '¥160-200', '¥200-260', '¥260+']},
        'channels': {ch: {'orders': channels[ch], 'pct': round(channels[ch] / total_orders * 100, 1),
                          'avg_gap': round(channel_gap[ch] / channels[ch]), 'revenue': round(channel_rev[ch])}
                     for ch in ['美团', '储值卡', '抖音', '直付', '其他'] if channels[ch]},
        'technicians': [{'name': t, 'orders': tech_orders[t], 'daily': round(tech_orders[t] / days_in_month, 1),
                         'revenue': round(tech_rev[t]), 'designated': tech_desig[t], 'private': tech_private[t]}
                        for t in sorted(tech_orders, key=lambda x: tech_orders[x], reverse=True)],
        'new_customer_orders': new_cust,
        'private_orders': private_orders,
        'days_in_month': days_in_month,
    }


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 extract_water.py <report.xlsx>")
        sys.exit(1)
    data = extract_water_bill(sys.argv[1])
    if data:
        print(json.dumps({k: v for k, v in data.items() if k != 'technicians'}, indent=2, ensure_ascii=False))
        print("technicians:", json.dumps(data['technicians'], ensure_ascii=False))
        base = os.path.dirname(os.path.abspath(__file__))
        out_path = os.path.join(base, '02_经营分析', '水单明细数据.json')
        with open(out_path, 'w') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"\nSaved to {out_path}")
