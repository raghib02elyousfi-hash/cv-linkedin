#!/usr/bin/env python3
"""Generate charts (corrected prices + lean scenario comparison)"""

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

BLUE  = '#1F3864'
MED   = '#2E75B6'
LIGHT = '#D6E4F0'
GOLD  = '#F4B942'
GREEN = '#70AD47'
RED   = '#C00000'
GREY  = '#F2F2F2'
ORANGE= '#ED7D31'

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'font.size': 9,
    'axes.titlesize': 11,
    'axes.titleweight': 'bold',
    'axes.titlecolor': BLUE,
    'axes.edgecolor': '#CCCCCC',
    'axes.spines.top': False,
    'axes.spines.right': False,
    'figure.facecolor': 'white',
})

# ── CORRECT FINANCIAL CONSTANTS ───────────────────────────────────────────────
SS_TOTAL    = 0.3178
PAGAS       = 14
HOURS_YEAR  = 1687.5
HOURS_MO    = HOURS_YEAR / 12     # 140.625
STRUCTURE   = 1430.0
STRUCT_LEAN = 780.0

SAL_IV      = 1107.22 * 1.03 * PAGAS        # annual salary Cat IV +3%
SS_EMP_Y    = (SAL_IV / 12) * 12 * SS_TOTAL  # employer SS
LABOR_MO    = (SAL_IV + SS_EMP_Y) / 12       # 1 756 €/month per employee

LABOR_H     = (SAL_IV + SS_EMP_Y) / HOURS_YEAR   # 12.47 €/h
ABSENT_H    = LABOR_H * 0.035                      # 0.44 €/h
DIRECT_B2B  = LABOR_H + ABSENT_H + 0.40 + 0.25    # 13.56 €/h
DIRECT_B2C  = LABOR_H + ABSENT_H + 0.55 + 0.45    # 13.71 €/h
STRUCT_3H   = STRUCTURE / (3 * HOURS_MO)            # 3.39 €/h

PRICE_B2B   = round((DIRECT_B2B + STRUCT_3H) * 1.30, 2)   # 22.03 €/h ✓
PRICE_B2C   = round((DIRECT_B2C + STRUCT_3H) * 1.30, 2)   # 22.22 €/h
PRICE_AVG   = PRICE_B2B * 0.70 + PRICE_B2C * 0.30          # 22.09 €/h
VAR_H       = (0.40 + 0.25) * 0.70 + (0.55 + 0.45) * 0.30 # 0.755 €/h
CONTRIB_H   = PRICE_AVG - VAR_H                             # 21.33 €/h

OWNER_SS    = [80]*6 + [160]*6 + [320]*6   # tarifa plana evolution

# ── CHART 1: Break-even (corrected) ──────────────────────────────────────────

fig, axes = plt.subplots(1, 3, figsize=(16, 6))
fig.suptitle(
    f'BRILLOPAL – Punto de Equilibrio por Escenario\n'
    f'(Precio B2B correcto: {PRICE_B2B:.2f} €/h | Precio B2C: {PRICE_B2C:.2f} €/h)',
    fontsize=13, fontweight='bold', color=BLUE, y=1.02)

scenarios_bep = [(1, GREEN, 'Estructura completa'),
                 (3, MED,   'Estructura completa'),
                 (5, BLUE,  'Estructura completa')]

for ax, (n_emp, color, _) in zip(axes, scenarios_bep):
    fixed_mo = n_emp * LABOR_MO + STRUCTURE + 80   # +80 owner SS ref.
    bep_h    = fixed_mo / CONTRIB_H
    avail    = n_emp * HOURS_MO

    h_range  = np.linspace(0, avail * 1.1, 400)
    revenue  = h_range * PRICE_AVG
    tot_cost = fixed_mo + h_range * VAR_H
    profit   = revenue - tot_cost

    ax.plot(h_range, revenue,   color=GREEN, lw=2.2, label='Ingresos')
    ax.plot(h_range, tot_cost,  color=RED,   lw=2.2, label='Costes totales')
    ax.fill_between(h_range, revenue, tot_cost,
                    where=(revenue >= tot_cost), alpha=0.15, color=GREEN, label='Beneficio')
    ax.fill_between(h_range, revenue, tot_cost,
                    where=(revenue <  tot_cost), alpha=0.15, color=RED,   label='Pérdida')

    ax.axvline(bep_h, color=GOLD, lw=2, ls='--')
    bep_rev = bep_h * PRICE_AVG
    ax.annotate(
        f'BEP\n{bep_h:.0f} h/mes\n({bep_h/avail*100:.0f}% ocup.)',
        xy=(bep_h, bep_rev),
        xytext=(bep_h + avail*0.08, bep_rev + fixed_mo*0.08),
        arrowprops=dict(arrowstyle='->', color=GOLD, lw=1.5),
        fontsize=8.5, color=GOLD, fontweight='bold')

    ax.text(avail*0.02, fixed_mo + fixed_mo*0.04,
            f'Costes fijos\n{fixed_mo:.0f} €/mes', fontsize=7.5, color=MED)
    ax.axhline(fixed_mo, color=MED, lw=1, ls=':', alpha=0.6)

    ax.set_title(f'{n_emp} empleado{"s" if n_emp>1 else ""}', color=BLUE, pad=8)
    ax.set_xlabel('Horas facturadas / mes', color=BLUE)
    ax.set_ylabel('€ / mes', color=BLUE)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:,.0f}€'))
    ax.set_xlim(0, avail * 1.1)
    ax.set_ylim(0, avail * PRICE_AVG * 1.1)
    ax.legend(fontsize=7.5, framealpha=0.9)
    ax.set_facecolor(GREY)

plt.tight_layout()
plt.savefig('/home/user/cv-linkedin/brillopal/BrillOPal_Grafico_BreakEven.png',
            dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print(f"Saved BreakEven (B2B={PRICE_B2B}€/h)")


# ── CHART 2: Cost waterfall + 18M EBITDA margin comparison ───────────────────

fig, axes = plt.subplots(1, 2, figsize=(15, 6))
fig.suptitle('BRILLOPAL – Estructura de costes y margen EBITDA por ocupación\n'
             f'Referencia: 3 empleados · Precio B2B {PRICE_B2B:.2f}€/h | B2C {PRICE_B2C:.2f}€/h',
             fontsize=12, fontweight='bold', color=BLUE)

# --- 2a: Cost waterfall per hour (B2B vs B2C) --------------------------------
ax = axes[0]
ax.set_title('Desglose de costes y precio de venta / hora (3 empleados)', color=BLUE)

cost_items_b2b = {
    f'Labor+SS\n{LABOR_H:.2f}€':    LABOR_H,
    f'Absentismo\n{ABSENT_H:.2f}€': ABSENT_H,
    f'Materiales\n0.40€':           0.40,
    f'Desplaz.\n0.25€':             0.25,
    f'Estructura\n{STRUCT_3H:.2f}€': STRUCT_3H,
}
cost_items_b2c = {
    f'Labor+SS\n{LABOR_H:.2f}€':    LABOR_H,
    f'Absentismo\n{ABSENT_H:.2f}€': ABSENT_H,
    f'Materiales\n0.55€':           0.55,
    f'Desplaz.\n0.45€':             0.45,
    f'Estructura\n{STRUCT_3H:.2f}€': STRUCT_3H,
}

colors_stack = [BLUE, MED, LIGHT, GOLD, '#8EA9DB']
b2b_vals  = list(cost_items_b2b.values())
b2c_vals  = list(cost_items_b2c.values())
b2b_total = sum(b2b_vals)
b2c_total = sum(b2c_vals)
b2b_margin = PRICE_B2B - b2b_total
b2c_margin = PRICE_B2C - b2c_total

bottom_b = np.zeros(1)
bottom_c = np.zeros(1)

for i, (label, vb, vc) in enumerate(zip(
        cost_items_b2b.keys(), b2b_vals, b2c_vals)):
    ax.bar([0],    [vb], bottom=bottom_b, color=colors_stack[i],
           label=label, width=0.38, edgecolor='white', lw=0.5)
    ax.bar([0.48], [vc], bottom=bottom_c, color=colors_stack[i],
           width=0.38, edgecolor='white', lw=0.5)
    bottom_b += vb
    bottom_c += vc

ax.bar([0],    [b2b_margin], bottom=[b2b_total],
       color=GREEN, width=0.38, edgecolor='white', label=f'Beneficio (+30%)', lw=0.5)
ax.bar([0.48], [b2c_margin], bottom=[b2c_total],
       color=GREEN, width=0.38, edgecolor='white', lw=0.5)

ax.text(0,    PRICE_B2B+0.25, f'{PRICE_B2B}€\n(s/IVA)',  ha='center',
        fontsize=9, fontweight='bold', color=GREEN)
ax.text(0.48, PRICE_B2C+0.25, f'{PRICE_B2C}€\n(s/IVA)',  ha='center',
        fontsize=9, fontweight='bold', color=GREEN)
ax.text(0,    PRICE_B2B*1.21+0.4, f'c/IVA:\n{PRICE_B2B*1.21:.2f}€',
        ha='center', fontsize=7.5, color=BLUE, style='italic')
ax.text(0.48, PRICE_B2C*1.21+0.4, f'c/IVA:\n{PRICE_B2C*1.21:.2f}€',
        ha='center', fontsize=7.5, color=BLUE, style='italic')

ax.set_xticks([0, 0.48])
ax.set_xticklabels(['B2B', 'B2C'], fontsize=12, fontweight='bold', color=BLUE)
ax.set_ylabel('€ / hora', color=BLUE)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:.1f}€'))
ax.legend(loc='upper left', fontsize=7.5, framealpha=0.9, ncol=2)
ax.set_facecolor(GREY)
ax.set_xlim(-0.45, 0.95)

# --- 2b: EBITDA vs occupancy (3 employees) -----------------------------------
ax2 = axes[1]
ax2.set_title('EBITDA mensual por % de ocupación (3 empleados)', color=BLUE)

occ_range = np.linspace(0, 1.0, 200)
fixed_3   = 3 * LABOR_MO + STRUCTURE + 160   # mid-year owner SS

ebitda_full = [3*HOURS_MO*o*PRICE_AVG - 3*HOURS_MO*o*VAR_H - fixed_3
               for o in occ_range]
ebitda_lean = [3*HOURS_MO*o*PRICE_AVG - 3*HOURS_MO*o*VAR_H
               - (3*LABOR_MO + STRUCT_LEAN + 160)
               for o in occ_range]

bep_full = (fixed_3) / (3*HOURS_MO*CONTRIB_H / (3*HOURS_MO))
bep_lean_v = (3*LABOR_MO + STRUCT_LEAN + 160)
bep_occ_full = fixed_3   / (3 * HOURS_MO * CONTRIB_H)
bep_occ_lean = bep_lean_v / (3 * HOURS_MO * CONTRIB_H)

ax2.plot(occ_range*100, [e/1000 for e in ebitda_full],
         color=BLUE, lw=2.5, label=f'Estructura completa (1.430€/mes)\nBEP: {bep_occ_full:.0%} ocupación')
ax2.plot(occ_range*100, [e/1000 for e in ebitda_lean],
         color=GREEN, lw=2.5, ls='--', label=f'Lean: home+coche propio (780€/mes)\nBEP: {bep_occ_lean:.0%} ocupación')

ax2.axhline(0, color=RED, lw=1, ls='--', alpha=0.7)
ax2.axvline(bep_occ_full*100, color=BLUE, lw=1, ls=':', alpha=0.6)
ax2.axvline(bep_occ_lean*100, color=GREEN, lw=1, ls=':', alpha=0.6)

ax2.fill_between(occ_range*100, [e/1000 for e in ebitda_full], 0,
                 where=[e>=0 for e in ebitda_full], alpha=0.10, color=BLUE)
ax2.fill_between(occ_range*100, [e/1000 for e in ebitda_lean], 0,
                 where=[e>=0 for e in ebitda_lean], alpha=0.10, color=GREEN)

# Annotate key occupancy points
for occ_pct in [0.75, 0.85, 0.95]:
    idx = int(occ_pct * 199)
    ef = ebitda_full[idx] / 1000
    el = ebitda_lean[idx] / 1000
    ax2.annotate(f'{ef:.1f}k€', xy=(occ_pct*100, ef),
                 fontsize=7, color=BLUE, ha='center', va='bottom')
    ax2.annotate(f'{el:.1f}k€', xy=(occ_pct*100, el),
                 fontsize=7, color=GREEN, ha='center', va='bottom')

ax2.set_xlabel('Ocupación mensual (%)', color=BLUE)
ax2.set_ylabel('EBITDA (miles €/mes)', color=BLUE)
ax2.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:.0f}%'))
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:.1f}k€'))
ax2.legend(fontsize=7.5, framealpha=0.9, loc='upper left')
ax2.set_facecolor(GREY)
ax2.set_xlim(0, 100)

plt.tight_layout()
plt.savefig('/home/user/cv-linkedin/brillopal/BrillOPal_Grafico_Margenes.png',
            dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved Margenes")


# ── CHART 3: Lean vs Full – 18-month progressive cash flow ───────────────────

SCHEDULE = [
    (0, 0.00), (0, 0.00),
    (1, 0.43), (1, 0.64),
    (1, 0.78), (1, 0.89),
    (2, 0.60), (2, 0.75),
    (2, 0.82), (2, 0.88),
    (2, 0.92), (3, 0.68),
    (3, 0.75), (3, 0.80),
    (3, 0.85), (3, 0.88),
    (3, 0.90), (3, 0.90),
]
OWNER_SS_MO = [80]*6 + [160]*6 + [320]*6
MONTHS_LABEL = [f"M{i+1}" for i in range(18)]

def cash_flow(struct, setup):
    acum = -setup
    monthly_ebitda, monthly_acum = [], []
    for m_i, (n_emp, occ) in enumerate(SCHEDULE):
        h_bill = n_emp * HOURS_MO * occ if n_emp > 0 else 0
        rev    = h_bill * PRICE_AVG
        costs  = n_emp * LABOR_MO + struct + OWNER_SS_MO[m_i] + h_bill * VAR_H
        ebitda = rev - costs
        acum  += ebitda
        monthly_ebitda.append(ebitda)
        monthly_acum.append(acum)
    return monthly_ebitda, monthly_acum

ebitda_full, acum_full = cash_flow(STRUCTURE,   2800)
ebitda_lean, acum_lean = cash_flow(STRUCT_LEAN,  800)

fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10), sharex=True)
fig.suptitle('BRILLOPAL – FLUJO DE CAJA PROGRESIVO: 18 meses\n'
             'Contratación escalonada (0→1→2→3 empleados)',
             fontsize=13, fontweight='bold', color=BLUE)

x = np.arange(18)
width = 0.35

# EBITDA mensual
bars_full = ax1.bar(x - width/2, [e/1000 for e in ebitda_full], width,
                    color=[GREEN if v>=0 else RED for v in ebitda_full],
                    alpha=0.75, label='Estructura completa (1.430€/mes)', edgecolor='white')
bars_lean = ax1.bar(x + width/2, [e/1000 for e in ebitda_lean], width,
                    color=[MED if v>=0 else ORANGE for v in ebitda_lean],
                    alpha=0.85, label='Lean: home+coche propio (780€/mes)', edgecolor='white')

ax1.axhline(0, color=BLUE, lw=1, alpha=0.4)
ax1.set_ylabel('EBITDA mensual (miles €)', color=BLUE)
ax1.set_title('EBITDA mensual por escenario de estructura', color=BLUE, pad=6)
ax1.legend(fontsize=8.5, framealpha=0.9)
ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:.1f}k€'))
ax1.set_facecolor(GREY)
ax1.tick_params(axis='y', colors=BLUE)

# Label first positive months
for i, v in enumerate(ebitda_full):
    if v > 0:
        ax1.annotate('✓F', xy=(i-width/2, v/1000+0.05), fontsize=7, color=GREEN, ha='center')
        break
for i, v in enumerate(ebitda_lean):
    if v > 0:
        ax1.annotate('✓L', xy=(i+width/2, v/1000+0.05), fontsize=7, color=MED, ha='center')
        break

# Employee count annotations
emp_counts = [n for n, _ in SCHEDULE]
prev = -1
for i, n in enumerate(emp_counts):
    if n != prev:
        ax1.axvline(i - 0.5, color=GOLD, lw=1.2, ls='--', alpha=0.6)
        if n > 0:
            ax1.text(i, ax1.get_ylim()[1]*0.88 if ax1.get_ylim()[1] > 0 else 0.5,
                     f'+Emp.{n}', fontsize=7.5, color=GOLD, fontweight='bold',
                     ha='center')
        prev = n

# Cumulative cash flow
ax2.plot(x, [a/1000 for a in acum_full], color=BLUE,  lw=2.5, marker='o', ms=5,
         label=f'Acumulado – Completa  (min: {min(acum_full)/1000:.1f}k€)')
ax2.plot(x, [a/1000 for a in acum_lean], color=GREEN, lw=2.5, marker='s', ms=5,
         ls='--', label=f'Acumulado – Lean  (min: {min(acum_lean)/1000:.1f}k€)')

ax2.fill_between(x, [a/1000 for a in acum_full], 0,
                 where=[a>=0 for a in acum_full], alpha=0.10, color=BLUE)
ax2.fill_between(x, [a/1000 for a in acum_lean], 0,
                 where=[a>=0 for a in acum_lean], alpha=0.10, color=GREEN)
ax2.fill_between(x, [a/1000 for a in acum_full], 0,
                 where=[a<0  for a in acum_full], alpha=0.10, color=RED)
ax2.fill_between(x, [a/1000 for a in acum_lean], 0,
                 where=[a<0  for a in acum_lean], alpha=0.10, color=ORANGE)

ax2.axhline(0, color=RED, lw=1.5, ls='--', alpha=0.6, label='Break-even acumulado')

# Annotate BEP
for label_str, acum_data, xoffset, col in [
    ('BEP\nCompleta', acum_full, +0.3, BLUE),
    ('BEP\nLean',     acum_lean, -0.3, GREEN),
]:
    for i, v in enumerate(acum_data):
        if v >= 0:
            ax2.annotate(label_str, xy=(i, 0),
                         xytext=(i+xoffset, -min(acum_full)/2000),
                         arrowprops=dict(arrowstyle='->', color=col, lw=1.5),
                         fontsize=8, color=col, fontweight='bold')
            break

ax2.set_ylabel('Flujo acumulado (miles €)', color=BLUE)
ax2.set_title('Flujo de caja acumulado – Capital mínimo necesario', color=BLUE, pad=6)
ax2.legend(fontsize=8.5, framealpha=0.9)
ax2.set_xticks(x)
ax2.set_xticklabels(MONTHS_LABEL, fontsize=8, color=BLUE)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:.1f}k€'))
ax2.set_facecolor(GREY)
ax2.tick_params(axis='both', colors=BLUE)

# Capital needed annotation
min_full = min(acum_full)
min_lean = min(acum_lean)
ax2.annotate(f'Capital\nnecesario\n{abs(min_full)/1000:.1f}k€',
             xy=(acum_full.index(min_full), min_full/1000),
             xytext=(acum_full.index(min_full)+1.5, min_full/1000-0.5),
             arrowprops=dict(arrowstyle='->', color=BLUE),
             fontsize=8, color=BLUE)
ax2.annotate(f'Capital\nnecesario\n{abs(min_lean)/1000:.1f}k€',
             xy=(acum_lean.index(min_lean), min_lean/1000),
             xytext=(acum_lean.index(min_lean)-2.5, min_lean/1000-0.3),
             arrowprops=dict(arrowstyle='->', color=GREEN),
             fontsize=8, color=GREEN)

plt.tight_layout()
plt.savefig('/home/user/cv-linkedin/brillopal/BrillOPal_Grafico_CashFlow_Lean.png',
            dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved CashFlow_Lean")
