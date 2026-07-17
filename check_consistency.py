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

if errors:
    print("\n❌ 不一致:")
    for e in errors:
        print(f"  {e}")
    sys.exit(1)
else:
    print("✅ 全部一致")
