# Design

## Theme

深色统一调色板。背景与卡片亮度差控制在 3:1 以内，长时间观看不疲劳。紫作为唯一强调色。

## Color Palette

| Token | Value | Usage |
|-------|-------|-------|
| `--bg` | #1c1e26 | 页面底色 |
| `--surface` | #252831 | 卡片/容器底色 |
| `--surface-accent` | #282636 | 强调卡片底色（如YTD行） |
| `--surface-raised` | #2c2f3a | 浮层/模拟器结果区 |
| `--border` | #323540 | 卡片/表格边框 |
| `--border-active` | #444858 | hover边框 |
| `--text` | #b4b8c4 | 正文 |
| `--text-dim` | #787c88 | 次级文字 |
| `--text-bright` | #eceef4 | 高亮文字/标题 |
| `--accent` | #8b7ec8 | 主强调色（紫） |
| `--green` | #68c090 | 正向指标 |
| `--red` | #d4686e | 负向指标 |
| `--amber` | #c89040 | 警告/中性 |
| `--cyan` | #68b8d0 | 辅助信息色 |
| `--gold` | #c89840 | 金色保留用于特殊标记 |

## Typography

| Role | Family | Weight | Size |
|------|--------|--------|------|
| Display heading | Noto Serif SC | 900 | 48px |
| Section title | Noto Sans SC | 700 | 11px |
| Card title | Noto Sans SC | 700 | 11px |
| Body | Noto Sans SC | 400 | — |
| Data / numbers | DM Mono | 400 | 40px (KPI) |
| Small data | DM Mono | 400 | 12px (table) |

## Spacing Scale

基于 4px 单位：4, 8, 10, 12, 14, 16, 20, 24, 28, 30, 32, 40, 52, 56, 64

## Components

### Card
- background: var(--surface)
- border: 1px var(--border)
- border-radius: 10px
- padding: 30px
- animation: revealUp 0.5s stagger

### KPI Card
- 同上，hover时 border-color → var(--accent), translateY(-2px)
- YTD行用 var(--surface-accent) + 紫色边框

### Data Table
- 斑马纹 nth-child(even) rgba(255,255,255,0.015)
- 行线 rgba(42,45,56,0.4)
- 首列 body font，末列右对齐
- hover: background var(--surface-raised)

### Bar Chart
- track: 18px height, border-radius 9px
- fill: border-radius 9px, 1.2s ease-out-quart

### Section Title
- ◆ 菱形标记（6px, accent色）
- font: 11px/700, letter-spacing 3px
- margin: 52px 0 22px

### Insight Box
- 上边框分隔
- 颜色文字（非色块背景）
- 三种状态：green / amber / red

### Badge
- 2px 8px padding, 9px font
- 三种等级：t1(purple) / t2(cyan) / t3(amber)

## Motion

- Card reveal: 0.5s ease-out-quart, 6-card stagger 60ms
- KPI hover: translateY(-2px)
- Bar fill: 1.2s ease-out-quart
- prefers-reduced-motion: 全部关闭

## Interaction

- focus-visible: 2px accent outline, 2px offset
- hover (desktop only): @media (hover: hover)
- touch: @media (hover: none) 禁用hover
