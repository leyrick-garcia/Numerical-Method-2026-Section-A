"""
Laboratory Exercise: Numerical explorations of the number e
--------------------------------------------------------------
Dark, "telemetry / HUD" visual system — deliberately different from the
light default-Matplotlib look in the source PDF. Every color used here
(BG, PANEL, GRID, TEXT, MUTED, CYAN, MAGENTA, AMBER) is also defined as a
CSS custom property in index.html, so the report and the figures share
one design system. Change a hex value below and re-run to keep both in
sync.

Outputs: exercise1.png, exercise2.png, exercise3.png, results.json
Requires: numpy, matplotlib
"""

import json
import math
from decimal import Decimal, getcontext

import numpy as np
import matplotlib.pyplot as plt

getcontext().prec = 60
E = math.e

# ----------------------------------------------------------------------
# Design tokens — mirrored 1:1 as CSS variables in index.html
# ----------------------------------------------------------------------
BG      = "#080B12"   # canvas
PANEL   = "#0E1420"   # axes background
GRID    = "#1C2536"   # gridlines / hairlines
TEXT    = "#E8EDF7"   # primary text
MUTED   = "#6E7A90"   # secondary text / ticks
CYAN    = "#29E5C6"   # Exercise 1 accent
MAGENTA = "#C74BFF"   # Exercise 2 accent
AMBER   = "#FFB020"   # Exercise 3 accent

MONO = "monospace"

plt.rcParams.update({
    "figure.facecolor": BG,
    "axes.facecolor": PANEL,
    "savefig.facecolor": BG,
    "text.color": TEXT,
    "axes.edgecolor": GRID,
    "axes.labelcolor": MUTED,
    "xtick.color": MUTED,
    "ytick.color": MUTED,
    "font.family": MONO,
    "font.size": 9.5,
    "grid.color": GRID,
    "grid.linewidth": 0.7,
})


def style_panel(ax, title, accent):
    """Give an axes the shared HUD look: dark panel, hairline grid,
    left-aligned monospace title, muted spines."""
    ax.set_facecolor(PANEL)
    ax.grid(True, alpha=0.55, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_color(GRID)
        spine.set_linewidth(0.8)
    ax.tick_params(colors=MUTED, labelsize=8.3)
    ax.set_title(title, color=TEXT, fontsize=10.5, fontweight="bold",
                 loc="left", family=MONO, pad=10)
    ax.title.set_bbox(dict(facecolor="none", edgecolor="none"))
    ax.plot(0.0, 1.09, marker="o", markersize=4.5, color=accent,
             transform=ax.transAxes, clip_on=False)


def hud_corners(ax, color, size=0.035, lw=1.3, alpha=0.9):
    """Four L-shaped corner brackets in axes-fraction coordinates —
    the instrument-panel framing device used on every chart."""
    for x, y in [(0, 0), (0, 1), (1, 0), (1, 1)]:
        dx = size if x == 0 else -size
        dy = size if y == 0 else -size
        ax.plot([x, x + dx], [y, y], transform=ax.transAxes, color=color,
                 lw=lw, alpha=alpha, clip_on=False, solid_capstyle="butt")
        ax.plot([x, x], [y, y + dy], transform=ax.transAxes, color=color,
                 lw=lw, alpha=alpha, clip_on=False, solid_capstyle="butt")


def glow_bars(ax, x, heights, color, width=0.62, z=3):
    """Bars with a soft halo behind them (stacked low-alpha copies) to
    fake a glow without leaving Matplotlib."""
    for w, a in [(width * 2.2, 0.06), (width * 1.6, 0.10), (width, 0.95)]:
        ax.bar(x, heights, width=w, color=color, alpha=a, zorder=z,
               edgecolor="none")
    ax.bar(x, heights, width=width, color="none", edgecolor=color,
           linewidth=0.9, alpha=0.9, zorder=z + 1)


def glow_line(ax, x, y, color, lw=2.0, z=3):
    for w, a in [(lw * 5, 0.08), (lw * 3, 0.14), (lw, 1.0)]:
        ax.plot(x, y, color=color, linewidth=w, alpha=a, zorder=z)


# ----------------------------------------------------------------------
# Exercise 1:  (1 + 1/n)^n  ->  e     as the compounding frequency n grows
# ----------------------------------------------------------------------
labels_1 = [
    "yearly", "1/2 yr", "quarter", "month", "week", "day",
    "hour", "minute", "second", "millisec", "microsec", "nanosec",
]
n_values = [1, 2, 4, 12, 52, 365, 8760, 525600, 31536000,
            31536000 * 1000, 31536000 * 1_000_000, 31536000 * 1_000_000_000]

# Decimal precision: at n ~ 1e16, float64's 1/n underflows machine epsilon
# and (1 + 1/n) would silently round to 1.0, flattening the tail wrongly.
vals_1_dec = [(Decimal(1) + Decimal(1) / Decimal(n)) ** n for n in n_values]
vals_1 = [float(v) for v in vals_1_dec]
err_1 = [abs(v - E) for v in vals_1]

fig, axes = plt.subplots(1, 2, figsize=(13, 5.6))
fig.suptitle("EX.01 — COMPOUNDING LIMIT   ·   (1 + 1/n)ⁿ → e",
              fontsize=12.5, fontweight="bold", color=TEXT, family=MONO, x=0.02, ha="left")

ax = axes[0]
x = np.arange(len(labels_1))
glow_bars(ax, x, vals_1, CYAN)
ax.axhline(E, color=MAGENTA, linestyle=(0, (5, 3)), linewidth=1.3, zorder=5,
           label=f"e = {E:.6f}")
ax.set_xticks(x)
ax.set_xticklabels(labels_1, rotation=55, ha="right")
ax.set_ylim(1.8, 3.0)
ax.set_ylabel("(1 + 1/n)ⁿ")
style_panel(ax, "VALUE PER COMPOUNDING PERIOD", CYAN)
ax.legend(loc="lower right", frameon=False, fontsize=8.5, labelcolor=TEXT)
hud_corners(ax, CYAN)

ax = axes[1]
glow_line(ax, x, err_1, AMBER)
ax.scatter(x, err_1, color=AMBER, s=22, zorder=5, edgecolor=BG, linewidth=0.6)
ax.set_yscale("log")
ax.set_xticks(x)
ax.set_xticklabels(labels_1, rotation=55, ha="right")
ax.set_ylabel("|error|  (log)")
style_panel(ax, "ERROR DECAY", AMBER)
hud_corners(ax, AMBER)

plt.tight_layout(rect=[0, 0, 1, 0.93])
plt.savefig("exercise1.png", dpi=150)
plt.close()

# ----------------------------------------------------------------------
# Exercise 2:  (a^h - 1) / h  ->  ln(a)     as h shrinks toward 0
# ----------------------------------------------------------------------
h_values = [0.1, 0.01, 0.001, 0.0001, 1e-5, 1e-6, 1e-7]
bases = {"a = 2": 2.0, "a = e": E, "a = 3": 3.0}
colors_2 = {"a = 2": CYAN, "a = e": AMBER, "a = 3": MAGENTA}

quotients = {name: [(a ** h - 1) / h for h in h_values] for name, a in bases.items()}
limits = {name: math.log(a) for name, a in bases.items()}

fig, ax = plt.subplots(figsize=(12.5, 6.6))
xg = np.arange(len(h_values))
width = 0.24
for i, (name, vals) in enumerate(quotients.items()):
    offset = (i - 1) * width
    c = colors_2[name]
    ax.bar(xg + offset, vals, width, color=c, alpha=0.95, zorder=3,
           label=f"{name}  ·  limit {limits[name]:.4f}",
           edgecolor=BG, linewidth=0.4)
for name, lim in limits.items():
    ax.axhline(lim, color=colors_2[name], linestyle=(0, (5, 3)), linewidth=1.0,
               alpha=0.85, zorder=2)

ax.set_xticks(xg)
ax.set_xticklabels([f"h={h:g}" for h in h_values])
ax.set_ylabel("(aʰ − 1) / h")
ax.set_ylim(0, 1.32)
style_panel(ax, "EX.02 — DIFFERENCE QUOTIENT   ·   (aʰ − 1)/h → ln(a)", MAGENTA)
ax.legend(loc="upper center", bbox_to_anchor=(0.5, 1.16), ncol=3, frameon=False,
          fontsize=9.2, labelcolor=TEXT, columnspacing=1.6, handlelength=1.4)
hud_corners(ax, MAGENTA)

plt.tight_layout(rect=[0, 0, 1, 0.90])
plt.savefig("exercise2.png", dpi=150)
plt.close()

# ----------------------------------------------------------------------
# Exercise 3:  e^x = sum_{n=0}^{inf} x^n / n!     (x = 1), up to N = 10000
# ----------------------------------------------------------------------
N_max = 10000
checkpoints = [1, 2, 3, 5, 10, 20, 50, 100, 1000, N_max]
checkpoint_sums = {}
digits_correct = {}

s, t = 1.0, 1.0
for n in range(1, N_max + 1):
    t *= 1.0 / n
    s += t
    if n in checkpoints:
        checkpoint_sums[n] = s
        err = abs(s - E)
        digits_correct[n] = -math.log10(err) if err > 0 else 16.0
checkpoint_sums[1] = 1.0
digits_correct.setdefault(1, -math.log10(abs(1.0 - E)))

fig, axes = plt.subplots(1, 2, figsize=(13, 5.6))
fig.suptitle("EX.03 — MACLAURIN SERIES   ·   eˣ = Σ xⁿ/n!  (x = 1, N ≤ 10,000)",
              fontsize=12.5, fontweight="bold", color=TEXT, family=MONO, x=0.02, ha="left")

ax = axes[0]
xN = np.arange(len(checkpoints))
cp_vals = [checkpoint_sums[k] for k in checkpoints]
glow_bars(ax, xN, cp_vals, MAGENTA)
ax.axhline(E, color=CYAN, linestyle=(0, (5, 3)), linewidth=1.3, zorder=5,
           label=f"e = {E:.6f}")
ax.set_xticks(xN)
ax.set_xticklabels([str(k) for k in checkpoints], rotation=45, ha="right")
ax.set_ylim(0.9, 3.0)
ax.set_ylabel("Sₙ = Σ 1/n!")
style_panel(ax, "PARTIAL SUMS", MAGENTA)
ax.legend(loc="lower right", frameon=False, fontsize=8.5, labelcolor=TEXT)
hud_corners(ax, MAGENTA)

ax = axes[1]
cp_digits = [digits_correct[k] for k in checkpoints]
glow_line(ax, checkpoints, cp_digits, AMBER)
ax.scatter(checkpoints, cp_digits, color=AMBER, s=24, zorder=5, edgecolor=BG, linewidth=0.6)
ax.set_xscale("log")
ax.set_ylabel("correct decimal digits")
style_panel(ax, "PRECISION GAINED vs N", AMBER)
for k, d in zip(checkpoints, cp_digits):
    ax.annotate(f"{d:.1f}", (k, d), textcoords="offset points", xytext=(0, 7),
                ha="center", fontsize=7.3, color=MUTED, family=MONO)
hud_corners(ax, AMBER)

plt.tight_layout(rect=[0, 0, 1, 0.93])
plt.savefig("exercise3.png", dpi=150)
plt.close()

# ----------------------------------------------------------------------
# Dump the numeric tables used above so the HTML can render exact values
# ----------------------------------------------------------------------
results = {
    "e": E,
    "palette": {"bg": BG, "panel": PANEL, "grid": GRID, "text": TEXT,
                "muted": MUTED, "cyan": CYAN, "magenta": MAGENTA, "amber": AMBER},
    "exercise1": {"labels": labels_1, "n": n_values, "values": vals_1},
    "exercise2": {"h": h_values, "bases": {name: quotients[name] for name in bases},
                  "limits": limits},
    "exercise3": {"checkpoints": checkpoints,
                  "sums": [checkpoint_sums[k] for k in checkpoints],
                  "digits": [digits_correct[k] for k in checkpoints]},
}

with open("results.json", "w") as f:
    json.dump(results, f, indent=2)

print("Done. Wrote exercise1.png, exercise2.png, exercise3.png, results.json")
