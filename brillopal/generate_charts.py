#!/usr/bin/env python3
"""Generate charts: break-even and B2B vs B2C margin comparison"""

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

plt.rcParams.update({
    'font.family': 'DejaVu Sans',
    'font.size': 9,
    'axes.titlesize': 12,
    'axes.titleweight': 'bold',
    'axes.titlecolor': BLUE,
    'axes.edgecolor': '#CCCCCC',
    'axes.spines.top': False,
    'axes.spines.right': False,
    'figure.facecolor': 'white',
})

# ── Chart 1: Break-Even ────────────────────────────────────────────────────────

fig, axes = plt.subplots(1, 3, figsize=(16, 6), sharey=False)
fig.suptitle('BRILLOPAL – Punto de Equilibrio por Escenario de Empleados',
             fontsize=14, fontweight='bold', color=BLUE, y=1.01)

SS_TOTAL = 0.3178
PAY14    = 14
HOURS_MO = 140.625
PREM     = 1.03
SAL_IV   = 1107.22 * PREM
LABOR_MO = (SAL_IV * PAY14 / 12) * (1 + SS_TOTAL)

STRUCT   = 1430.0
OWNER_SS = 80.0

REV_B2B  = 20.41
REV_B2C  = 22.50
DIR_B2B  = 13.22 + 0.40 + 0.25
DIR_B2C  = 13.22 + 0.55 + 0.45
B2B_PCT  = 0.70
B2C_PCT  = 0.30

# Weighted avg revenue and direct cost per hour
REV_AVG  = REV_B2B * B2B_PCT + REV_B2C * B2C_PCT
DIR_AVG  = DIR_B2B * B2B_PCT + DIR_B2C * B2C_PCT
VAR_MARG = REV_AVG - DIR_AVG  # contribution margin per hour

scenarios = [(1, GREEN), (3, MED), (5, BLUE)]
labels = ["1 empleado", "3 empleados", "5 empleados"]

for ax, (n_emp, color), label in zip(axes, scenarios, labels):
    fixed = STRUCT + n_emp * LABOR_MO + OWNER_SS
    bep   = fixed / VAR_MARG
    avail = n_emp * HOURS_MO
    h_range = np.linspace(0, avail * 1.05, 300)

    revenue = h_range * REV_AVG
    total_cost = fixed + h_range * DIR_AVG
    profit = revenue - total_cost

    ax.plot(h_range, revenue,    color=GREEN, lw=2,   label='Ingresos')
    ax.plot(h_range, total_cost, color=RED,   lw=2,   label='Costes totales')
    ax.fill_between(h_range, revenue, total_cost,
                    where=(revenue >= total_cost), alpha=0.15, color=GREEN,
                    label='Zona beneficio')
    ax.fill_between(h_range, revenue, total_cost,
                    where=(revenue < total_cost),  alpha=0.15, color=RED,
                    label='Zona pérdida')

    ax.axvline(bep, color=GOLD, lw=2, ls='--')
    ax.annotate(f'BEP\n{bep:.0f} h/mes\n({bep/avail*100:.0f}% ocupac.)',
                xy=(bep, fixed + bep * DIR_AVG),
                xytext=(bep + avail*0.07, fixed + bep * DIR_AVG + fixed*0.05),
                arrowprops=dict(arrowstyle='->', color=GOLD),
                fontsize=8, color=GOLD, fontweight='bold')

    ax.axhline(fixed, color=MED, lw=1, ls=':', alpha=0.7)
    ax.text(avail*0.02, fixed + fixed*0.03, f'Costes fijos\n{fixed:.0f} €/mes',
            fontsize=7.5, color=MED)

    ax.set_title(f'Escenario: {label}', color=BLUE, pad=8)
    ax.set_xlabel('Horas facturadas / mes', color=BLUE)
    ax.set_ylabel('€ / mes', color=BLUE)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:,.0f}€'))
    ax.set_xlim(0, avail * 1.05)
    ax.set_ylim(0, avail * REV_AVG * 1.05)
    ax.legend(fontsize=7.5, framealpha=0.8)
    ax.set_facecolor(GREY)
    ax.tick_params(colors=BLUE)

plt.tight_layout()
plt.savefig('/home/user/cv-linkedin/brillopal/BrillOPal_Grafico_BreakEven.png',
            dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved: BrillOPal_Grafico_BreakEven.png")


# ── Chart 2: B2B vs B2C margin comparison ─────────────────────────────────────

fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle('BRILLOPAL – Comparativa de Márgenes B2B vs B2C\n(Escenario 3 empleados)',
             fontsize=13, fontweight='bold', color=BLUE)

# --- 2a: Cost waterfall per hour ---
ax = axes[0]
ax.set_title('Desglose de costes y precio de venta / hora', color=BLUE)

struct_h = STRUCT / (3 * HOURS_MO)

# B2B
b2b_costs = {
    'Labor+SS\n(12,47€)': 12.47,
    'Absentismo\n(0,44€)': 0.44,
    'Materiales\n(0,40€)': 0.40,
    'Desplaz.\n(0,25€)': 0.25,
    f'Estructura\n({struct_h:.2f}€)': struct_h,
}
# B2C
b2c_costs = {
    'Labor+SS\n(12,47€)': 12.47,
    'Absentismo\n(0,44€)': 0.44,
    'Materiales\n(0,55€)': 0.55,
    'Desplaz.\n(0,45€)': 0.45,
    f'Estructura\n({struct_h:.2f}€)': struct_h,
}

x = np.arange(len(b2b_costs))
width = 0.35
colors_stack = [BLUE, MED, LIGHT, GOLD, '#8EA9DB']

b2b_vals = list(b2b_costs.values())
b2c_vals = list(b2c_costs.values())

b2b_total = sum(b2b_vals)
b2c_total = sum(b2c_vals)
b2b_price = b2b_total * 1.30
b2c_price = b2c_total * 1.30
b2b_margin = b2b_price - b2b_total
b2c_margin = b2c_price - b2c_total

bottom_b2b = np.zeros(1)
bottom_b2c = np.zeros(1)

for i, (label, (vb, vc)) in enumerate(zip(b2b_costs.keys(),
                                           zip(b2b_vals, b2c_vals))):
    ax.bar([0], [vb], bottom=bottom_b2b, color=colors_stack[i],
           label=label, width=0.35, edgecolor='white', linewidth=0.5)
    ax.bar([0.45], [vc], bottom=bottom_b2c, color=colors_stack[i],
           width=0.35, edgecolor='white', linewidth=0.5)
    bottom_b2b += vb
    bottom_b2c += vc

# Add margin on top
ax.bar([0],    [b2b_margin], bottom=[b2b_total], color=GREEN,
       width=0.35, edgecolor='white', label=f'Beneficio 30%', linewidth=0.5)
ax.bar([0.45], [b2c_margin], bottom=[b2c_total], color=GREEN,
       width=0.35, edgecolor='white', linewidth=0.5)

# Annotations
ax.text(0,    b2b_price + 0.3, f'{b2b_price:.2f}€\n(s/IVA)', ha='center',
        fontsize=9, fontweight='bold', color=GREEN)
ax.text(0.45, b2c_price + 0.3, f'{b2c_price:.2f}€\n(s/IVA)', ha='center',
        fontsize=9, fontweight='bold', color=GREEN)

ax.set_xticks([0, 0.45])
ax.set_xticklabels(['B2B', 'B2C'], fontsize=11, fontweight='bold', color=BLUE)
ax.set_ylabel('€ / hora', color=BLUE)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: f'{x:.1f}€'))
ax.legend(loc='upper left', fontsize=7.5, framealpha=0.85, ncol=2)
ax.set_facecolor(GREY)
ax.set_xlim(-0.4, 0.9)


# --- 2b: Monthly margin evolution over 18 months ---
ax2 = axes[1]
ax2.set_title('Evolución del margen mensual – 3 empleados (18 meses)', color=BLUE)

RAMP = [0.35,0.45,0.55,0.65,0.70,0.72,
        0.75,0.75,0.78,0.80,0.82,0.85,
        0.85,0.82,0.87,0.88,0.90,0.90]

months = [f"M{i+1}" for i in range(18)]
n_emp = 3
rev_b2b_mo = [HOURS_MO*n_emp*r*B2B_PCT*REV_B2B for r in RAMP]
rev_b2c_mo = [HOURS_MO*n_emp*r*B2C_PCT*REV_B2C for r in RAMP]
owner_ss = [80]*6 + [160]*6 + [320]*6
labor_mo = [n_emp * LABOR_MO]*18
mat_mo   = [HOURS_MO*n_emp*r*(DIR_B2B-12.47-0.44) for r in RAMP]
total_rev = [b+c for b,c in zip(rev_b2b_mo, rev_b2c_mo)]
total_cost = [l+o+m+STRUCT for l,o,m in zip(labor_mo, owner_ss, mat_mo)]
ebitda = [r-c for r,c in zip(total_rev, total_cost)]

x = np.arange(18)
ax2.bar(x, [r/1000 for r in total_rev], color=LIGHT, label='Ingresos (miles €)', edgecolor=MED, linewidth=0.5)
ax2.plot(x, [e/1000 for e in ebitda], color=GREEN, lw=2.5, marker='o', ms=5,
         label='EBITDA (miles €)', zorder=5)
ax2.axhline(0, color=RED, lw=1, ls='--', alpha=0.6)
ax2.fill_between(x, [e/1000 for e in ebitda], 0,
                 where=[e>=0 for e in ebitda], alpha=0.2, color=GREEN)
ax2.fill_between(x, [e/1000 for e in ebitda], 0,
                 where=[e<0 for e in ebitda], alpha=0.2, color=RED)

# Mark BEP month
for i, e in enumerate(ebitda):
    if e >= 0:
        ax2.annotate(f'Rentable\nM{i+1}',
                     xy=(i, ebitda[i]/1000), xytext=(i+0.8, ebitda[i]/1000+0.5),
                     arrowprops=dict(arrowstyle='->', color=GOLD),
                     fontsize=7.5, color=GOLD, fontweight='bold')
        break

ax2.set_xticks(x)
ax2.set_xticklabels(months, fontsize=7, rotation=45, color=BLUE)
ax2.set_ylabel('Miles de € / mes', color=BLUE)
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v,_: f'{v:.1f}k€'))
ax2.legend(fontsize=8, framealpha=0.85)
ax2.set_facecolor(GREY)

plt.tight_layout()
plt.savefig('/home/user/cv-linkedin/brillopal/BrillOPal_Grafico_Margenes.png',
            dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("Saved: BrillOPal_Grafico_Margenes.png")
