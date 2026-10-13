#!/usr/bin/env python3
"""
十三绝门店月度报表数据提取器
输入：月度报表 Excel 文件路径
输出：结构化 JSON，包含所有经营分析需要的核心数据
用法：python extract_report.py /path/to/report.xlsx
"""

import openpyxl
import json
import sys
import os
import datetime as dt_lib
from collections import defaultdict


def safe_float(v):
    try:
        return float(v) if v is not None else 0
    except (ValueError, TypeError):
        return 0


def extract_all(filepath):
    """从月报中提取所有核心数据"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    result = {
        "source": os.path.basename(filepath),
        "extracted_at": dt_lib.datetime.now().isoformat(),
    }

    # ========== 1. 店面经营报表 ==========
    ws = wb['店面经营报表']
    
    # Detect year ranges from merged headers
    # Columns: 2024=C(3)-P(16), 2025=Q(17)-AD(30), 2026=AF(32)-AS(45)
    layout = {
        '2024': {'start': 3, 'end': 16, 'months': ['01','02','03','04','05','06','07','08','09','10','11','12'], 'cum_col': 15, 'pct_col': 16},
        '2025': {'start': 17, 'end': 30, 'months': ['01','02','03','04','05','06','07','08','09','10','11','12'], 'cum_col': 29, 'pct_col': 30},
        '2026': {'start': 32, 'end': 45, 'months': ['01','02','03','04','05','06','07','08','09','10','11','12'], 'cum_col': 44, 'pct_col': 45},
    }

    monthly = []
    
    for year, cfg in layout.items():
        for i, month in enumerate(cfg['months']):
            col = cfg['start'] + i
            if col > cfg['end']:
                break
            
            revenue = safe_float(ws.cell(row=3, column=col).value)  # 主营收入
            if revenue == 0:
                continue
            
            other_income = safe_float(ws.cell(row=4, column=col).value)
            total_revenue = safe_float(ws.cell(row=5, column=col).value)
            
            entry = {
                "period": f"{year}-{month}",
                "revenue": round(total_revenue, 2),
                "main_revenue": round(revenue, 2),
                "other_income": round(other_income, 2),
                
                # 人力成本
                "labor_salary": round(safe_float(ws.cell(row=6, column=col).value), 2),
                "labor_social": round(safe_float(ws.cell(row=7, column=col).value), 2),
                "labor_meals": round(safe_float(ws.cell(row=8, column=col).value), 2),
                "labor_dorm_util": round(safe_float(ws.cell(row=9, column=col).value), 2),
                "labor_dorm_rent": round(safe_float(ws.cell(row=10, column=col).value), 2),
                "labor_benefits": round(safe_float(ws.cell(row=11, column=col).value), 2),
                "labor_training": round(safe_float(ws.cell(row=12, column=col).value), 2),
                "labor_uniform": round(safe_float(ws.cell(row=13, column=col).value), 2),
                "labor_total": round(safe_float(ws.cell(row=14, column=col).value), 2),
                
                # 店铺基本成本
                "mgmt_fee": round(safe_float(ws.cell(row=15, column=col).value), 2),
                "rent": round(safe_float(ws.cell(row=16, column=col).value), 2),
                "utilities": round(safe_float(ws.cell(row=17, column=col).value), 2),
                "property": round(safe_float(ws.cell(row=18, column=col).value), 2),
                "office": round(safe_float(ws.cell(row=19, column=col).value), 2),
                "bank_fees": round(safe_float(ws.cell(row=20, column=col).value), 2),
                "consumables": round(safe_float(ws.cell(row=21, column=col).value), 2),
                "non_consumables": round(safe_float(ws.cell(row=22, column=col).value), 2),
                "repairs": round(safe_float(ws.cell(row=23, column=col).value), 2),
                "travel": round(safe_float(ws.cell(row=24, column=col).value), 2),
                "phone": round(safe_float(ws.cell(row=25, column=col).value), 2),
                "refunds": round(safe_float(ws.cell(row=26, column=col).value), 2),
                "accidents": round(safe_float(ws.cell(row=27, column=col).value), 2),
                "interest": round(safe_float(ws.cell(row=28, column=col).value), 2),
                "tax_fees": round(safe_float(ws.cell(row=29, column=col).value), 2),
                "base_cost_total": round(safe_float(ws.cell(row=30, column=col).value), 2),
                
                # 营销成本
                "promotion": round(safe_float(ws.cell(row=31, column=col).value), 2),
                "advertising": round(safe_float(ws.cell(row=32, column=col).value), 2),
                "free_service_cost": round(safe_float(ws.cell(row=33, column=col).value), 2),
                "koc_visits": round(safe_float(ws.cell(row=34, column=col).value), 2),
                "marketing_total": round(safe_float(ws.cell(row=35, column=col).value), 2),
                
                # 汇总
                "total_cost": round(safe_float(ws.cell(row=36, column=col).value), 2),
                "gross_profit": round(safe_float(ws.cell(row=37, column=col).value), 2),
                "manager_bonus": round(safe_float(ws.cell(row=38, column=col).value), 2),
                "operating_profit": round(safe_float(ws.cell(row=39, column=col).value), 2),
                
                # 税务行（行42-43可能有不同年份的差异）
                "income_tax": round(safe_float(ws.cell(row=42, column=col).value), 2),
                "net_profit": round(safe_float(ws.cell(row=43, column=col).value), 2),
            }
            
            # 计算比率
            if entry["revenue"] > 0:
                entry["labor_ratio"] = round(entry["labor_total"] / entry["revenue"] * 100, 1)
                entry["net_margin"] = round(entry["net_profit"] / entry["revenue"] * 100, 1)
                entry["base_cost_ratio"] = round(entry["base_cost_total"] / entry["revenue"] * 100, 1)
                entry["marketing_ratio"] = round(entry["marketing_total"] / entry["revenue"] * 100, 1)
            
            monthly.append(entry)
    
    result["monthly"] = monthly

    # ========== 2. 水单明细 → 渠道 & 储值分析 ==========
    # Sheet name varies: "十三绝系统" vs "伊智系统"
    detail_name = None
    for s in wb.sheetnames:
        if '水单明细' in s:
            detail_name = s
            break
    if not detail_name:
        raise KeyError("找不到水单明细 sheet")
    ws_detail = wb[detail_name]
    
    channels = defaultdict(lambda: {"orders": 0, "raw": 0, "received": 0})
    order_types = defaultdict(lambda: {"count": 0, "amount": 0})
    stored_value = {"充值": 0, "开卡": 0, "疗程套餐": 0, "卡金消费": 0}
    total_service_perf = 0
    
    # === 表头映射（兼容伊智25列 / 十三绝33列 / 37列等结构变化）===
    _hdr = None
    for _r in ws_detail.iter_rows(min_row=1, max_row=3, values_only=True):
        if _r and any(v is not None and '订单类型' in str(v) for v in _r):
            _hdr = list(_r); break
    _IX = {}
    if _hdr:
        _IX = {str(v).strip(): k for k, v in enumerate(_hdr) if v is not None}

    def _col(*names):
        for n in names:
            if n in _IX:
                return _IX[n]
        for n in names:
            for k, v in _IX.items():
                if n in k:
                    return v
        return None

    _C = {
        'type': _col('订单类型'),
        'raw': _col('原价业绩', '门店业绩'),
        'recv': _col('实收金额', '实收业绩'),
        'perf': _col('实收业绩', '员工业绩'),
        'plat': _col('平台类型'),
        'card': _col('卡金业绩'),
        'cash': _col('现金业绩'),
        'other': _col('其他业绩'),
        'pay': _col('支付明细'),
    }

    def _g(row, key):
        idx = _C.get(key)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    # 伊智旧系统无表头（或表头缺列）时，回退到固定列号
    if _C['raw'] is None:
        is_new_system = ws_detail.max_column > 30
    else:
        is_new_system = True
    print(f"  水单列映射: " + ", ".join(f"{k}={v}" for k, v in _C.items() if v is not None)
          + ("  [回退固定列号]" if _C['raw'] is None else ""))

    for row in ws_detail.iter_rows(min_row=2, max_row=ws_detail.max_row, max_col=ws_detail.max_column, values_only=True):
        try:
            if _C['type'] is not None:
                order_type = str(_g(row, 'type') or '').strip()
                raw_price = safe_float(_g(row, 'raw'))
                received = safe_float(_g(row, 'recv'))
                perf = safe_float(_g(row, 'perf'))
                platform = str(_g(row, 'plat') or '').strip()
                card_perf = safe_float(_g(row, 'card'))
                cash_perf = safe_float(_g(row, 'cash'))
                other_perf = safe_float(_g(row, 'other'))
                pay = str(_g(row, 'pay') or '')
            else:
                # === 伊智系统: 固定列回退 ===
                order_type = str(row[1] or '').strip()
                raw_price = safe_float(row[8])
                received = safe_float(row[8])
                perf = safe_float(row[8])
                platform = ''
                card_perf = cash_perf = other_perf = 0
                pay = ''

            if '开单' in order_type:
                order_type = '服务项目'
            if raw_price == 0:
                continue

            channels[platform]['orders'] += 1
            channels[platform]['raw'] += raw_price
            channels[platform]['received'] += received

            order_types[order_type]['count'] += 1
            order_types[order_type]['amount'] += received

            if '充值' in order_type:
                stored_value['充值'] += received
            elif '开卡' in order_type:
                stored_value['开卡'] += received
            elif '疗程' in order_type:
                stored_value['疗程套餐'] += received
            else:
                total_service_perf += raw_price
                stored_value['卡金消费'] += card_perf
        except Exception:
            continue

    # Dedup: 卡金消费 already counted
    channel_summary = {}
    for ch, data in channels.items():
        ch_name = ch if ch else '直付/现金'
        channel_summary[ch_name] = {
            "orders": data['orders'],
            "raw": round(data['raw'], 2),
            "received": round(data['received'], 2),
            "discount_rate": round((1 - data['received']/data['raw'])*100, 1) if data['raw'] > 0 else 0,
            "share": round(data['received']/sum(c['received'] for c in channels.values())*100, 1) if sum(c['received'] for c in channels.values()) > 0 else 0,
        }
    
    result["water_bill"] = {
        "channels": channel_summary,
        "order_types": {k: {"count": v['count'], "amount": round(v['amount'], 2)} for k, v in sorted(order_types.items())},
        "stored_value": {
            "inflow_充值": round(stored_value['充值'], 2),
            "inflow_开卡": round(stored_value['开卡'], 2),
            "inflow_套餐": round(stored_value['疗程套餐'], 2),
            "inflow_total": round(stored_value['充值'] + stored_value['开卡'] + stored_value['疗程套餐'], 2),
            "outflow_消费": round(stored_value['卡金消费'], 2),
            "net_change": round(stored_value['充值'] + stored_value['开卡'] + stored_value['疗程套餐'] - stored_value['卡金消费'], 2),
        },
        "service_revenue": round(total_service_perf, 2),
    }

    # ========== 3. 店长提成 ==========
    try:
        ws_bonus = wb['店长提成奖励']
        result["manager_bonus"] = {
            "actual_performance": round(safe_float(ws_bonus.cell(row=1, column=2).value), 2),
            "target_performance": round(safe_float(ws_bonus.cell(row=1, column=4).value), 2),
            "salary_cost": round(safe_float(ws_bonus.cell(row=2, column=2).value), 2),
            "utilities_cost": round(safe_float(ws_bonus.cell(row=3, column=2).value), 2),
            "consumables_cost": round(safe_float(ws_bonus.cell(row=5, column=2).value), 2),
            "total_manager_cost": round(safe_float(ws_bonus.cell(row=6, column=2).value), 2),
            "manager_margin": round(safe_float(ws_bonus.cell(row=7, column=2).value), 2),
            "bonus_amount": round(safe_float(ws_bonus.cell(row=14, column=2).value), 2),
        }
    except:
        result["manager_bonus"] = None

    # ========== 4. 总部结算 ==========
    try:
        ws_settle = wb[[s for s in wb.sheetnames if '总部结算' in s][0]]
        deductions = []
        settle_amount = 0
        for row in ws_settle.iter_rows(min_row=3, max_row=ws_settle.max_row, max_col=7, values_only=True):
            desc = str(row[1] or '').strip()
            amount = safe_float(row[2])
            if not desc:
                continue
            if '结算金额' in desc or '应结算' in desc:
                settle_amount = amount
            else:
                deductions.append({"item": desc, "amount": round(amount, 2)})
        
        result["settlement"] = {
            "deductions": deductions,
            "total_deductions": round(sum(d['amount'] for d in deductions), 2),
            "net_settlement": round(settle_amount, 2),
        }
    except:
        result["settlement"] = None

    # ========== 5. 银行流水摘要 ==========
    try:
        bank_name = None
        for s in wb.sheetnames:
            if '银行交易流水' in s:
                bank_name = s
                break
        if bank_name:
            ws_bank = wb[bank_name]
        cash_flows = {"inflow": [], "outflow": []}
        for row in ws_bank.iter_rows(min_row=4, max_row=ws_bank.max_row, max_col=11, values_only=True):
            date_val = str(row[0] or '')
            purpose = str(row[1] or '').strip()
            debit = safe_float(row[2])  # 支出
            credit = safe_float(row[3])  # 收入
            counterparty = str(row[6] or '').strip()
            
            if debit > 0:
                cash_flows['outflow'].append({
                    "date": date_val[:10] if date_val else '',
                    "purpose": purpose,
                    "amount": round(debit, 2),
                    "counterparty": counterparty,
                })
            elif credit > 0:
                cash_flows['inflow'].append({
                    "date": date_val[:10] if date_val else '',
                    "purpose": purpose,
                    "amount": round(credit, 2),
                    "counterparty": counterparty,
                })
        
        result["bank_flows"] = cash_flows
    except:
        result["bank_flows"] = None

    # ========== 6. 核心指标一键速览 ==========
    if monthly:
        latest = monthly[-1]
        result["dashboard"] = {
            "period": latest["period"],
            "revenue": latest["revenue"],
            "net_profit": latest["net_profit"],
            "net_margin": latest["net_margin"],
            "labor_ratio": latest["labor_ratio"],
            "marketing_ratio": latest["marketing_ratio"],
            "payroll": latest["labor_salary"],
            "rent": latest["rent"],
            "promotion_fee": latest["promotion"],
            "mgmt_fee": latest["mgmt_fee"],
        }
        
        # 卡金数据
        if "water_bill" in result:
            sv = result["water_bill"]["stored_value"]
            result["dashboard"]["stored_value_in"] = sv["inflow_total"]
            result["dashboard"]["stored_value_out"] = sv["outflow_消费"]
            result["dashboard"]["stored_value_net"] = sv["net_change"]

    # ========== 7. 工资表提取（表头映射，兼容报表结构变化）==========
    wage_sheet_names = [s for s in wb.sheetnames if '工资表' in s]
    if wage_sheet_names:
        ws_wage = wb[wage_sheet_names[0]]
        maxc = ws_wage.max_column
        h2 = [str(ws_wage.cell(2, c).value or '').strip() for c in range(1, maxc + 1)]
        h3 = [str(ws_wage.cell(3, c).value or '').strip() for c in range(1, maxc + 1)]
        colnames = {}
        for _i in range(maxc):
            _n = h3[_i] or h2[_i]
            if _n:
                colnames.setdefault(_n, _i + 1)

        def find_col(*cands):
            for cand in cands:
                if cand in colnames:
                    return colnames[cand]
            for cand in cands:
                for n, idx in colnames.items():
                    if cand in n:
                        return idx
            return None

        FIELD_MAP = {
            '序号': ('序号',), '姓名': ('姓名',),
            '入职日期': ('正式上钟日期', '入职日期'),
            '店铺': ('所属店铺', '店铺'), '岗位': ('岗位',),
            '应出勤': ('应出勤天数', '应出勤'), '实际出勤': ('实际出勤天数', '实际出勤'),
            '应发底薪': ('应发底薪', '岗位工资'), '实发底薪': ('实发底薪', '实发岗位工资'),
            '绩效工资': ('绩效工资',), '实发绩效工资': ('实发绩效工资',),
            '全勤补贴': ('全勤补贴', '全勤'), '妆服补贴': ('妆服补贴',), '底薪合计': ('底薪合计',),
            '劳动业绩': ('总劳动业绩', '劳动业绩'), '业绩提成': ('业绩提成',),
            '指定业绩': ('指定业绩',), '指定提成点': ('指定提成点',), '指定提成': ('指定提成',),
            '非指定业绩': ('非指定业绩',), '非指定提成点': ('非指定提成点',), '非指定提成': ('非指定提成',),
            '开卡金额': ('开卡总金额', '开卡金额'), '开卡提成点': ('开卡提成点',), '开卡提成': ('开卡提成',),
            '基础提成': ('基础提成',), '工资小计': ('工资小计', '应发小计'),
            '奖金': ('奖金', '加店长利润提成'),
            '证书补贴': ('证书补贴',), '工龄补贴': ('加工龄补贴', '工龄补贴'),
            '岗位补贴': ('岗位补贴',), '餐补': ('餐补',),
            '师徒带教补贴': ('师徒带教补贴',),
            '个人客流提成': ('加：个人客流提成', '个人客流提成'),
            '门店充值奖励': ('门店充值奖励',), '达标奖励': ('达标奖励',),
            '好评奖励': ('好评奖励', '加：美团/抖音优质好评奖励'),
            '扣点客未达标': ('扣点客未达标',), '扣业绩未达标': ('扣业绩未达标',),
            '扣充值未达标': ('扣充值未达标',), '扣好评未达标': ('扣好评未达标',),
            '扣差评客诉': ('扣差评客诉', '减差评客诉'),
            '扣请假迟到': ('扣请假迟到', '减请假/迟到'),
            '其他扣款': ('其他扣款', '其他扣补贴明细/未达标'),
            '扣绩效': ('扣绩效',),
            '应发合计': ('应发工资合计', '应发合计'),
            '扣餐费': ('扣餐费', '减餐费'), '扣水电物业': ('扣水电物业', '减水电物业'),
            '扣房费': ('扣房费', '减房费'), '扣社保': ('扣社保', '减社保'),
            '扣个税': ('扣个税', '减个税'), '扣工服': ('扣工服', '减工服'),
            '减预支': ('减预支',), '预留款': ('预留款',),
            '实发工资': ('实发工资', '实际发放'), '备注': ('备注',),
        }
        CMAP = {k: find_col(*v) for k, v in FIELD_MAP.items()}
        _mapped = ", ".join(f"{k}={v}" for k, v in CMAP.items() if v)
        print(f"  工资表列映射({len(colnames)}列): {_mapped}")

        wage_records = []
        for row in ws_wage.iter_rows(min_row=4, max_row=ws_wage.max_row, values_only=True):
            def g(key, _row=row):
                idx = CMAP.get(key)
                if not idx or idx > len(_row):
                    return None
                return _row[idx - 1]
            name = str(g('姓名') or '').strip()
            if not name:
                continue
            wage_records.append({
                "序号": safe_float(g('序号')),
                "姓名": name,
                "入职日期": str(g('入职日期') or '')[:10],
                "店铺": str(g('店铺') or '').strip(),
                "岗位": str(g('岗位') or '').strip(),
                "应出勤": safe_float(g('应出勤')),
                "实际出勤": safe_float(g('实际出勤')),
                "应发底薪": round(safe_float(g('应发底薪')), 2),
                "实发底薪": round(safe_float(g('实发底薪')), 2),
                "绩效工资": round(safe_float(g('绩效工资')), 2),
                "实发绩效工资": round(safe_float(g('实发绩效工资')), 2),
                "全勤补贴": round(safe_float(g('全勤补贴')), 2),
                "妆服补贴": round(safe_float(g('妆服补贴')), 2),
                "底薪合计": round(safe_float(g('底薪合计')), 2),
                "劳动业绩": round(safe_float(g('劳动业绩')), 2),
                "业绩提成": round(safe_float(g('业绩提成')), 2),
                "指定业绩": round(safe_float(g('指定业绩')), 2),
                "指定提成点": safe_float(g('指定提成点')),
                "指定提成": round(safe_float(g('指定提成')), 2),
                "非指定业绩": round(safe_float(g('非指定业绩')), 2),
                "非指定提成点": safe_float(g('非指定提成点')),
                "非指定提成": round(safe_float(g('非指定提成')), 2),
                "开卡金额": round(safe_float(g('开卡金额')), 2),
                "开卡提成点": safe_float(g('开卡提成点')),
                "开卡提成": round(safe_float(g('开卡提成')), 2),
                "基础提成": round(safe_float(g('基础提成')), 2),
                "工资小计": round(safe_float(g('工资小计')), 2),
                "奖金": round(safe_float(g('奖金')), 2),
                "证书补贴": round(safe_float(g('证书补贴')), 2),
                "工龄补贴": round(safe_float(g('工龄补贴')), 2),
                "岗位补贴": round(safe_float(g('岗位补贴')), 2),
                "餐补": round(safe_float(g('餐补')), 2),
                "师徒带教补贴": round(safe_float(g('师徒带教补贴')), 2),
                "个人客流提成": round(safe_float(g('个人客流提成')), 2),
                "门店充值奖励": round(safe_float(g('门店充值奖励')), 2),
                "达标奖励": round(safe_float(g('达标奖励')), 2),
                "好评奖励": round(safe_float(g('好评奖励')), 2),
                "扣点客未达标": round(safe_float(g('扣点客未达标')), 2),
                "扣业绩未达标": round(safe_float(g('扣业绩未达标')), 2),
                "扣充值未达标": round(safe_float(g('扣充值未达标')), 2),
                "扣好评未达标": round(safe_float(g('扣好评未达标')), 2),
                "扣差评客诉": round(safe_float(g('扣差评客诉')), 2),
                "扣请假迟到": round(safe_float(g('扣请假迟到')), 2),
                "其他扣款": round(safe_float(g('其他扣款')), 2),
                "扣绩效": round(safe_float(g('扣绩效')), 2),
                "应发合计": round(safe_float(g('应发合计')), 2),
                "扣餐费": round(safe_float(g('扣餐费')), 2),
                "扣水电物业": round(safe_float(g('扣水电物业')), 2),
                "扣房费": round(safe_float(g('扣房费')), 2),
                "扣社保": round(safe_float(g('扣社保')), 2),
                "扣个税": round(safe_float(g('扣个税')), 2),
                "扣工服": round(safe_float(g('扣工服')), 2),
                "减预支": round(safe_float(g('减预支')), 2),
                "预留款": round(safe_float(g('预留款')), 2),
                "实发工资": round(safe_float(g('实发工资')), 2),
                "备注": str(g('备注') or '').strip(),
            })

        # Categorize
        non_tech = [r for r in wage_records if r['岗位'] in ('店长', '前台', '保洁', '财务')]
        techs = [r for r in wage_records if '技师' in r['岗位']]

        result["wage_table"] = {
            "period": latest["period"] if monthly else "",
            "records": wage_records,
            "summary": {
                "非技师": {
                    "人次": len(non_tech),
                    "底薪合计": round(sum(r['应发底薪'] for r in non_tech), 2),
                    "实发合计": round(sum(r['实发工资'] for r in non_tech), 2),
                    "明细": [{"姓名": r['姓名'], "岗位": r['岗位'], "底薪": r['应发底薪'], "实发": r['实发工资']} for r in non_tech],
                },
                "技师": {
                    "人次": len(techs),
                    "提成合计": round(sum(r['工资小计'] for r in techs), 2),
                    "实发合计": round(sum(r['实发工资'] for r in techs), 2),
                    "指定业绩总计": round(sum(r['指定业绩'] for r in techs), 2),
                    "非指定业绩总计": round(sum(r['非指定业绩'] for r in techs), 2),
                    "劳动业绩总计": round(sum(r['劳动业绩'] for r in techs), 2),
                },
                "全部实发总计": round(sum(r['实发工资'] for r in wage_records), 2),
            }
        }

    # ========== 8. 费用明细表提取 ==========
    if '费用明细表' in [s for s in wb.sheetnames]:
        ws_fee = wb['费用明细表']
        fee_records = []
        fee_by_category = defaultdict(lambda: {"count": 0, "amount": 0})
        
        for row in ws_fee.iter_rows(min_row=3, max_row=ws_fee.max_row, max_col=7, values_only=True):
            date_val = str(row[1] or '') if len(row) > 1 else ''
            desc = str(row[2] or '').strip() if len(row) > 2 else ''
            category = str(row[3] or '').strip() if len(row) > 3 else ''
            amount = safe_float(row[4]) if len(row) > 4 else 0
            
            if not desc or amount == 0:
                continue
            
            # Parse date (could be Excel serial number or YYYY-MM-DD)
            date_str = ''
            try:
                if date_val and date_val != 'None':
                    if isinstance(row[1], (int, float)):
                        from datetime import datetime as dt_parser, timedelta
                        date_str = (dt_parser(1899, 12, 30) + timedelta(days=int(date_val))).strftime('%Y-%m-%d')
                    else:
                        date_str = date_val[:10]
            except:
                date_str = str(date_val)[:10]
            
            rec = {
                "日期": date_str,
                "摘要": desc,
                "类别": category,
                "金额": round(amount, 2),
                "备注": str(row[5] or '').strip() if len(row) > 5 else '',
            }
            fee_records.append(rec)
            fee_by_category[category]["count"] += 1
            fee_by_category[category]["amount"] += amount
        
        result["expense_detail"] = {
            "records": fee_records,
            "total": round(sum(r['金额'] for r in fee_records), 2),
            "by_category": {k: {"count": v["count"], "amount": round(v["amount"], 2)} for k, v in sorted(fee_by_category.items())},
        }

    # ========== 9. 银行存款日记账提取 ==========
    if '银行存款日记账' in [s for s in wb.sheetnames]:
        ws_bank_journal = wb['银行存款日记账']
        journal_records = []
        
        for row in ws_bank_journal.iter_rows(min_row=4, max_row=min(ws_bank_journal.max_row, 500), max_col=12, values_only=True):
            if not row[1]:
                continue
            date_val = str(row[1] or '') if len(row) > 1 else ''
            desc = str(row[2] or '').strip() if len(row) > 2 else ''
            category = str(row[3] or '').strip() if len(row) > 3 else ''
            expense_invoice = safe_float(row[4]) if len(row) > 4 else 0
            expense_no_invoice = safe_float(row[5]) if len(row) > 5 else 0
            balance = safe_float(row[7]) if len(row) > 7 else 0
            
            if not desc:
                continue
            
            date_str = ''
            try:
                if date_val and date_val != 'None':
                    if isinstance(row[1], (int, float)):
                        from datetime import datetime, timedelta
                        date_str = (datetime(1899, 12, 30) + timedelta(days=int(date_val))).strftime('%Y-%m-%d')
                    else:
                        date_str = date_val[:10]
            except:
                date_str = str(date_val)[:10]
            
            rec = {
                "日期": date_str,
                "摘要": desc,
                "类别": category,
                "支出_有发票": round(expense_invoice, 2),
                "支出_无发票": round(expense_no_invoice, 2),
                "支出合计": round(expense_invoice + expense_no_invoice, 2),
                "余额": round(balance, 2),
            }
            journal_records.append(rec)
        
        month_records = [r for r in journal_records if r['日期'] and (latest['period'][:7] if monthly else '') in r['日期']]
        
        result["bank_journal"] = {
            "all_records": journal_records,
            "current_month": month_records,
            "monthly_expense": round(sum(r['支出合计'] for r in month_records), 2) if month_records else 0,
            "ending_balance": round(month_records[-1]['余额'], 2) if month_records else 0,
        }

    # ========== 10. 更新核心指标看板（补全工资/费用数据） ==========
    if "wage_table" in result:
        result["dashboard"]["wage_non_tech_total"] = result["wage_table"]["summary"]["非技师"]["实发合计"]
        result["dashboard"]["wage_tech_total"] = result["wage_table"]["summary"]["技师"]["实发合计"]
        result["dashboard"]["wage_all_total"] = result["wage_table"]["summary"]["全部实发总计"]
    if "expense_detail" in result:
        result["dashboard"]["expense_detail_total"] = result["expense_detail"]["total"]
    if "bank_journal" in result:
        result["dashboard"]["bank_ending_balance"] = result["bank_journal"]["ending_balance"]
    
    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python extract_report.py <report.xlsx>")
        sys.exit(1)
    
    data = extract_all(sys.argv[1])
    print(json.dumps(data, ensure_ascii=False, indent=2))
