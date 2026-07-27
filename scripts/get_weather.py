#!/usr/bin/env python3
"""获取玉林历史天气 (Open-Meteo API)
用法: python3 get_weather.py 2026-07-19
输出: 上午状况~下午状况 高/低温° 降水Zmm
"""
import sys, json, urllib.request
from datetime import date

target = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else date.today()

url = (f"https://api.open-meteo.com/v1/forecast?"
       f"latitude=22.63&longitude=110.15"
       f"&hourly=temperature_2m,precipitation,weather_code"
       f"&timezone=Asia/Shanghai"
       f"&start_date={target}&end_date={target}")

with urllib.request.urlopen(url, timeout=8) as resp:
    d = json.load(resp)['hourly']

WC = {0:'晴',1:'晴',2:'多云',3:'阴',45:'雾',48:'雾',
      51:'小雨',53:'小雨',55:'小雨',61:'中雨',63:'大雨',65:'暴雨',
      80:'阵雨',81:'暴雨',82:'大暴雨',95:'雷阵雨',96:'雷暴',99:'大雷暴'}

codes, total_rain, temps = set(), 0, []
for i in range(len(d['time'])):
    codes.add(d['weather_code'][i])
    total_rain += d['precipitation'][i]
    temps.append(d['temperature_2m'][i])

descs = sorted(set(WC.get(c, f'?{c}') for c in codes))
rng = f'{descs[0]}~{descs[-1]}' if len(descs) > 1 else descs[0]
print(f'{rng} {max(temps):.0f}/{min(temps):.0f}° 降水{total_rain:.1f}mm')
