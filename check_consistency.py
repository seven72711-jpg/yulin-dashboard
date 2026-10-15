"""跨文件一致性校验——提交前必跑"""
import json, re, sys

errors = []

# 1. 技师人数、姓名、等级一致性
with open("00_项目总览/技师名单.md") as f:
    md = f.read()
with open("00_项目总览/数字孪生.md") as f:
    twin = f.read()
with open("dashboard.html") as f:
    html = f.read()

# 提取技师名单中的12人
import re
names_md = set(re.findall(r'\|\s*(林燕|唐婉清|禤妮|郑家珍|梁雪华|陈宇华|谭庆彬|吕茜茜|姚梨|丘洁荣|陈燕军|蔡饶萱)\s*\|', md))
names_twin = set(re.findall(r'\|\s*(林燕|唐婉清|禤妮|郑家珍|梁雪华|陈宇华|谭庆彬|吕茜茜|姚梨|丘洁荣|陈燕军|蔡饶萱)\s*\|', twin))
names_html = set(re.findall(r"n:'(林燕|唐婉清|禤妮|郑家珍|梁雪华|陈宇华|谭庆彬|吕茜茜|姚梨|丘洁荣|陈燕军|蔡饶萱)'", html))

if names_md != names_twin:
    errors.append(f"技师名单 vs 数字孪生: {names_md ^ names_twin}")
if names_md != names_html:
    errors.append(f"技师名单 vs dashboard: {names_md ^ names_html}")
if names_twin != names_html:
    errors.append(f"数字孪生 vs dashboard: {names_twin ^ names_html}")

print(f"技师人数: 技师名单={len(names_md)}, 数字孪生={len(names_twin)}, dashboard={len(names_html)}")

# 2. 每个技师的等级一致性
# Extract grade from 技师名单
grades_md = {}
for name in names_md:
    m = re.search(rf'\|\s*{name}\s*\|\s*(特级|高级|中级|初级)\s*\|', md)
    if m: grades_md[name] = m.group(1)

grades_twin = {}
for name in names_twin:
    m = re.search(rf'\|\s*{name}\s*\|\s*(特级|高级|中级|初级)\s*\|', twin)
    if m: grades_twin[name] = m.group(1)

for name in names_md & names_twin:
    g1 = grades_md.get(name)
    g2 = grades_twin.get(name)
    if g1 and g2 and g1 != g2:
        errors.append(f"{name}等级不一致: 技师名单={g1}, 数字孪生={g2}")

# 3. 用语合规检查——报告/文案中不得出现的AI式表达
BANNED = [
    # 标签式小标题（用户点名反感，非第一次）
    r'\*\*\s*一句话\s*\*\*', r'一句话总结', r'一句话结论',
    r'说白了', r'说到底', r'说穿了',
    # AI 句式
    r'不是[^，。]{1,20}而是', r'不仅[^，。]{1,20}更', r'既是[^，。]{1,20}也是',
    r'本质上', r'赋能', r'重塑', r'旨在', r'底层逻辑', r'黄金法则', r'里程碑',
    r'值得注意的是', r'综上所述',
]
SCAN_DIRS = ['02_经营分析/月度报告', '00_项目总览', '03_策略建议']
SCAN_ROOT = ['AI_INDEX.md']
found = []
import os
for d in SCAN_DIRS:
    if not os.path.isdir(d):
        continue
    for root, _, files in os.walk(d):
        for fn in files:
            if not fn.endswith(('.md', '.html')):
                continue
            p = os.path.join(root, fn)
            if '_旧版归档' in p or '_历史备份' in p:
                continue
            with open(p, encoding='utf-8') as f:
                for ln, line in enumerate(f, 1):
                    for pat in BANNED:
                        if re.search(pat, line):
                            found.append(f"{p}:{ln} [{pat}] {line.strip()[:60]}")
for p in SCAN_ROOT:
    if os.path.exists(p):
        with open(p, encoding='utf-8') as f:
            for ln, line in enumerate(f, 1):
                for pat in BANNED:
                    if re.search(pat, line):
                        found.append(f"{p}:{ln} [{pat}] {line.strip()[:60]}")

if found:
    print(f"\n❌ 用语违规 {len(found)} 处:")
    for x in found:
        print(f"  {x}")
    errors.extend(found)
else:
    print("用语合规: ✅ 无AI式标签")

if errors:
    print("\n❌ 不一致:")
    for e in errors:
        print(f"  {e}")
    sys.exit(1)
else:
    print("✅ 全部一致")
