"""
Full Dissolved Gas Analysis (DGA) — Transformer Oil
=====================================================
Methods covered:
  1. Duval's Triangle (IEC 60599)
  2. Rogers Ratio Method (IEC 60599)
  3. Doernenburg Ratio Method
  4. Key Gas Method (IEEE C57.104)
  5. IEC 60599 Ratio Method (Revised)
  6. TDCG / IEEE C57.104 Condition Assessment
  7. CO2/CO Ratio (Cellulose Degradation)
  8. Individual Gas Limit Check (IS 9434)

All results are written to a multi-sheet Excel report and
all charts are saved as PNG files in the same folder.
"""

import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Polygon



# ==============================
# INPUT GAS VALUES
# ==============================

H2   = float(input("Enter H2 value: "))
O2   = float(input("Enter O2 value: "))
N2   = float(input("Enter N2 value: "))
CO   = float(input("Enter CO value: "))
CH4  = float(input("Enter CH4 value: "))
CO2  = float(input("Enter CO2 value: "))
C2H4 = float(input("Enter C2H4 value: "))
C2H6 = float(input("Enter C2H6 value: "))
C2H2 = float(input("Enter C2H2 value: "))
C3H6 = float(input("Enter C3H6 value: "))
C3H8 = float(input("Enter C3H8 value: "))

# Derived totals
TDCG = H2 + CO + CH4 + C2H4 + C2H6 + C2H2
# ─────────────────────────────────────────────────────────────────────────────
# ② INDIVIDUAL GAS LIMITS  (IS 9434 / IEEE C57.104)
# ─────────────────────────────────────────────────────────────────────────────
GAS_LIMITS = {
    # (value, limit_val, limit_str, unit)
    "H₂  (Hydrogen)":         (H2,   50,   "50 Max",   "ppm"),
    "O₂  (Oxygen)":           (O2,   None, "NS",       "ppm"),
    "N₂  (Nitrogen)":         (N2,   None, "NS",       "ppm"),
    "CO  (Carbon Monoxide)":  (CO,   400,  "400 Max",  "ppm"),
    "CH₄ (Methane)":          (CH4,  30,   "30 Max",   "ppm"),
    "CO₂ (Carbon Dioxide)":   (CO2,  3800, "3800 Max", "ppm"),
    "C₂H₄ (Ethylene)":        (C2H4, 60,   "60 Max",   "ppm"),
    "C₂H₆ (Ethane)":          (C2H6, 20,   "20 Max",   "ppm"),
    "C₂H₂ (Acetylene)":       (C2H2, 2,    "2 Max",    "ppm"),
    "C₃H₆ (Propylene)":       (C3H6, None, "NS",       "ppm"),
    "C₃H₈ (Propane)":         (C3H8, None, "NS",       "ppm"),
    "TDCG":                   (TDCG, None, "NS",       "ppm"),
}

# ─────────────────────────────────────────────────────────────────────────────
# ③ TDCG / IEEE C57.104 CONDITION LEVELS
# ─────────────────────────────────────────────────────────────────────────────
def tdcg_condition(tdcg_val):
    if tdcg_val <= 720:
        return 1, "Normal Operation", "No action required.", "E2EFDA"
    elif tdcg_val <= 1920:
        return 2, "Caution", "Resample & re-test in 1 month.", "FFEB9C"
    elif tdcg_val <= 4630:
        return 3, "Warning", "Resample in 1 week. Plan inspection.", "FFC7CE"
    else:
        return 4, "Critical", "Immediate action / take offline.", "FF0000"

# ─────────────────────────────────────────────────────────────────────────────
# ④ CO₂/CO RATIO — Cellulose / Paper Degradation
# ─────────────────────────────────────────────────────────────────────────────
def co2_co_ratio(co2_val, co_val):
    if co_val == 0:
        return None, "Indeterminate (CO = 0)"
    r = co2_val / co_val
    if r < 3:
        return r, "Possible cellulose fault (overheating of paper insulation)"
    elif r <= 10:
        return r, "Normal cellulose aging — no active paper fault"
    else:
        return r, "Possible CO₂ from non-fault source OR extensive paper degradation"

# ─────────────────────────────────────────────────────────────────────────────
# ⑤ ROGERS RATIO METHOD (IEC 60599)
# ─────────────────────────────────────────────────────────────────────────────
def rogers_ratio(h2, ch4, c2h2, c2h4, c2h6):
    """Returns (fault_code, fault_description, ratios_dict)"""
    r1 = c2h2 / c2h4 if c2h4 > 0 else 0    # C2H2/C2H4
    r2 = ch4  / h2   if h2   > 0 else 0    # CH4/H2
    r5 = c2h4 / c2h6 if c2h6 > 0 else 0    # C2H4/C2H6

    def code(val, thresholds):
        """Map ratio to Rogers code (0/1/2)."""
        lo, hi = thresholds
        if val < lo:   return 0
        elif val < hi: return 1
        else:          return 2

    c1 = code(r1, (0.1, 3.0))
    c2 = code(r2, (0.1, 1.0))
    c5 = code(r5, (1.0, 3.0))

    cases = {
        (0,0,0): "Normal Aging",
        (0,1,0): "Partial Discharge (low energy density)",
        (1,1,0): "Partial Discharge (high energy density / tracking)",
        (1,0,1): "Low Energy Discharge (arcing, corona)",
        (1,0,2): "High Energy Discharge (arcing)",
        (0,0,1): "Thermal Fault < 150°C (slight overheating)",
        (0,0,2): "Thermal Fault 150–200°C",
        (0,1,1): "Thermal Fault 200–300°C",
        (0,1,2): "Thermal Fault > 300°C (conductor overheating)",
    }
    key = (c1, c2, c5)
    fault = cases.get(key, f"Unclassified (codes: C2H2/C2H4={c1}, CH4/H2={c2}, C2H4/C2H6={c5})")
    return fault, {"C₂H₂/C₂H₄": r1, "CH₄/H₂": r2, "C₂H₄/C₂H₆": r5}

# ─────────────────────────────────────────────────────────────────────────────
# ⑥ DOERNENBURG RATIO METHOD
# ─────────────────────────────────────────────────────────────────────────────
def doernenburg(h2, ch4, c2h2, c2h4, c2h6, co):
    """
    Requires minimum gas concentrations to be valid.
    Returns (fault, ratios, valid)
    """
    min_vals = {"H2": 100, "CH4": 120, "C2H2": 35, "C2H4": 50}
    valid = (h2 >= min_vals["H2"] or ch4 >= min_vals["CH4"] or
             c2h2 >= min_vals["C2H2"] or c2h4 >= min_vals["C2H4"])

    r1 = ch4  / h2   if h2   > 0 else 0   # CH4/H2
    r2 = c2h2 / c2h4 if c2h4 > 0 else 0   # C2H2/C2H4
    r3 = c2h2 / ch4  if ch4  > 0 else 0   # C2H2/CH4
    r4 = c2h6 / c2h2 if c2h2 > 0 else 0   # C2H6/C2H2 (inverse for check)

    if not valid:
        fault = "Insufficient gas levels — Doernenburg method not applicable"
    elif r1 > 1.0 and r2 < 0.75 and r3 < 0.3:
        fault = "Thermal Decomposition"
    elif r1 < 0.1 and r2 < 0.75 and r3 < 0.3:
        fault = "Partial Discharge (corona)"
    elif r2 >= 0.75 and r3 >= 0.3:
        fault = "Arcing (high energy discharge)"
    else:
        fault = "Normal / No significant fault"

    return fault, {"CH₄/H₂": r1, "C₂H₂/C₂H₄": r2, "C₂H₂/CH₄": r3}, valid

# ─────────────────────────────────────────────────────────────────────────────
# ⑦ IEC 60599 REVISED RATIO METHOD
# ─────────────────────────────────────────────────────────────────────────────
def iec_60599(h2, ch4, c2h2, c2h4, c2h6):
    r1 = c2h2 / c2h4 if c2h4 > 0 else 0
    r2 = ch4  / h2   if h2   > 0 else 0

    if r1 < 0.1:
        if r2 < 0.1:
            fault = "Partial Discharge"
        elif r2 < 1.0:
            fault = "Thermal Fault (stray gassing / low temp)"
        else:
            fault = "Thermal Fault (high temperature > 300°C)"
    elif r1 < 3.0:
        if r2 < 0.1:
            fault = "Discharge of Low Energy (sparking)"
        else:
            fault = "Discharge of High Energy (arcing)"
    else:
        fault = "Discharge of High Energy (severe arcing)"

    return fault, {"C₂H₂/C₂H₄": r1, "CH₄/H₂": r2}

# ─────────────────────────────────────────────────────────────────────────────
# ⑧ KEY GAS METHOD (IEEE C57.104 / Transformer Aging Guide)
# ─────────────────────────────────────────────────────────────────────────────
def key_gas_method(h2, ch4, c2h2, c2h4, c2h6, co):
    """
    Identifies dominant fault type from the single highest key gas.
    """
    gases = {
        "H₂":   h2,
        "CH₄":  ch4,
        "C₂H₄": c2h4,
        "C₂H₂": c2h2,
        "C₂H₆": c2h6,
        "CO":   co,
    }
    key = max(gases, key=gases.get)
    val = gases[key]

    interpretations = {
        "H₂":   "Corona / Partial Discharge (low-energy sparking in oil)",
        "CH₄":  "Low-temperature thermal fault (< 300°C in oil)",
        "C₂H₄": "High-temperature thermal fault in oil (> 300°C)",
        "C₂H₂": "Electrical arcing / high-energy discharge",
        "C₂H₆": "Low-to-moderate thermal fault in oil (< 200°C)",
        "CO":   "Thermal fault involving cellulose (paper/pressboard overheating)",
    }
    return key, val, interpretations[key], gases

# ─────────────────────────────────────────────────────────────────────────────
# ⑨ DUVAL TRIANGLE
# ─────────────────────────────────────────────────────────────────────────────
def ternary_to_cart(pCH4, pC2H4, pC2H2):
    a, b, c = pCH4/100, pC2H4/100, pC2H2/100
    x = 0.5 * (2*b + c) / (a+b+c)
    y = (np.sqrt(3)/2) * c / (a+b+c)
    return x, y

DUVAL_ZONES = {
    "PD": {"color":"#B3D9FF","label":"PD\nPartial\nDischarge",
           "points":[(98,0,2),(100,0,0),(98,2,0)]},
    "T1": {"color":"#FFFF99","label":"T1\n<300°C",
           "points":[(98,0,2),(98,2,0),(76,24,0),(77,0,23)]},
    "T2": {"color":"#FFD966","label":"T2\n300–700°C",
           "points":[(77,0,23),(76,24,0),(40,60,0),(46,0,54)]},
    "T3": {"color":"#FF9900","label":"T3\n>700°C",
           "points":[(46,0,54),(40,60,0),(0,100,0),(0,93,7),(0,0,100)]},
    "D1": {"color":"#FF7F7F","label":"D1\nLow Energy\nDischarge",
           "points":[(100,0,0),(98,2,0),(76,24,0),(87,0,13)]},
    "D2": {"color":"#FF3333","label":"D2\nHigh Energy\nDischarge",
           "points":[(87,0,13),(76,24,0),(40,60,0),(23,0,77)]},
    "DT": {"color":"#CC66FF","label":"DT\nMixed\nDisch+Therm",
           "points":[(23,0,77),(40,60,0),(0,93,7),(0,0,100)]},
}

DUVAL_MEANINGS = {
    "PD":"Partial Discharge","T1":"Thermal Fault < 300°C",
    "T2":"Thermal Fault 300–700°C","T3":"Thermal Fault > 700°C",
    "D1":"Low Energy Electrical Discharge","D2":"High Energy Discharge (Arc)",
    "DT":"Mixed Discharge + Thermal",
}

def classify_duval(pCH4, pC2H4, pC2H2):
    if pC2H2 >= 29:
        return "DT" if pCH4 <= 23 else "D2"
    if pC2H2 >= 13:
        return "D2"
    if pC2H2 >= 2:
        if pC2H4 < 24: return "T1"
        return "D1" if pCH4 >= 87 else "D2"
    if pC2H4 >= 60: return "T3"
    if pC2H4 >= 24: return "T2"
    if pCH4  >= 98: return "PD"
    return "T1"

def plot_duval(ch4_v, c2h4_v, c2h2_v):
    total = ch4_v + c2h4_v + c2h2_v
    if total == 0:
        pCH4, pC2H4, pC2H2 = 33.3, 33.3, 33.3
    else:
        pCH4  = ch4_v  / total * 100
        pC2H4 = c2h4_v / total * 100
        pC2H2 = c2h2_v / total * 100

    fault_zone = classify_duval(pCH4, pC2H4, pC2H2)

    fig, ax = plt.subplots(figsize=(10, 9), facecolor="white")
    ax.set_aspect("equal")
    ax.axis("off")

    for zname, zdata in DUVAL_ZONES.items():
        pts = [ternary_to_cart(*p) for p in zdata["points"]]
        poly = Polygon(pts, closed=True, facecolor=zdata["color"],
                       edgecolor="white", linewidth=1.0, alpha=0.85, zorder=1)
        ax.add_patch(poly)
        cx = np.mean([p[0] for p in pts])
        cy = np.mean([p[1] for p in pts])
        ax.text(cx, cy, zdata["label"], ha="center", va="center",
                fontsize=7.5, fontweight="bold", color="#222222",
                zorder=4, linespacing=1.3)

    tri = Polygon([(0,0),(1,0),(0.5,np.sqrt(3)/2)], closed=True,
                  fill=False, edgecolor="black", linewidth=2, zorder=5)
    ax.add_patch(tri)

    for tv in [20,40,60,80]:
        for pts in [
            [ternary_to_cart(tv,0,100-tv),    ternary_to_cart(tv,100-tv,0)],
            [ternary_to_cart(0,tv,100-tv),    ternary_to_cart(100-tv,tv,0)],
            [ternary_to_cart(0,100-tv,tv),    ternary_to_cart(100-tv,0,tv)],
        ]:
            ax.plot([pts[0][0],pts[1][0]], [pts[0][1],pts[1][1]],
                    color="gray", lw=0.4, ls="--", zorder=2)
        x,y = ternary_to_cart(100-tv,tv,0)
        ax.text(x, y-0.04, f"{tv}%", ha="center", va="top", fontsize=7, color="#555")

    off = 0.045
    ax.text(-off, -off,           "100%\nCH₄",  ha="center", va="top",    fontsize=9, fontweight="bold")
    ax.text(1+off, -off,          "100%\nC₂H₄", ha="center", va="top",    fontsize=9, fontweight="bold")
    ax.text(0.5,  np.sqrt(3)/2+0.02,"100%\nC₂H₂",ha="center",va="bottom", fontsize=9, fontweight="bold")

    sx, sy = ternary_to_cart(pCH4, pC2H4, pC2H2)
    ax.scatter([sx],[sy], s=220, color="red", edgecolors="black",
               linewidths=1.5, zorder=10, marker="*")
    ax.annotate(
        f"  Sample\n  CH₄={pCH4:.1f}%\n  C₂H₄={pC2H4:.1f}%\n  C₂H₂={pC2H2:.1f}%",
        (sx,sy), xytext=(sx+0.13, sy+0.07), fontsize=8, color="darkred",
        fontweight="bold",
        arrowprops=dict(arrowstyle="->", color="darkred", lw=1.2),
        bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                  edgecolor="darkred", alpha=0.9), zorder=11)

    info = (f"Gas Values\nCH₄  = {ch4_v:.2f} ppm → {pCH4:.1f}%\n"
            f"C₂H₄ = {c2h4_v:.2f} ppm → {pC2H4:.1f}%\n"
            f"C₂H₂ = {c2h2_v:.2f} ppm → {pC2H2:.1f}%\n"
            f"──────────────────\n"
            f"Zone: {fault_zone}\n→ {DUVAL_MEANINGS[fault_zone]}")
    ax.text(1.02, 0.75, info, transform=ax.transAxes, fontsize=8.5,
            va="top", ha="left", fontfamily="monospace",
            bbox=dict(boxstyle="round,pad=0.5", facecolor="#FFF8DC",
                      edgecolor="#999", alpha=0.95))

    patches = [mpatches.Patch(color=v["color"],
               label=v["label"].replace("\n"," ")) for v in DUVAL_ZONES.values()]
    ax.legend(handles=patches, loc="lower right", bbox_to_anchor=(1.38,0.0),
              fontsize=7.5, title="Fault Zones", framealpha=0.9, title_fontsize=8)

    ax.set_title("Duval's Triangle — IEC 60599\nTransformer DGA Analysis",
                 fontsize=13, fontweight="bold", pad=14)
    plt.tight_layout()
    plt.show()
    return fault_zone, pCH4, pC2H4, pC2H2

# ═════════════════════════════════════════════════════════════════════════════
# RUN ALL ANALYSES
# ═════════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("  FULL DGA ANALYSIS — TRANSFORMER OIL")
print("=" * 60)

# Duval
duval_zone, pCH4, pC2H4, pC2H2 = plot_duval(CH4, C2H4, C2H2)
print(f"\n[Duval Triangle]  Zone: {duval_zone} — {DUVAL_MEANINGS[duval_zone]}")

# Rogers
rogers_fault, rogers_ratios = rogers_ratio(H2, CH4, C2H2, C2H4, C2H6)
print(f"[Rogers Ratio]    {rogers_fault}")

# Doernenburg
doern_fault, doern_ratios, doern_valid = doernenburg(H2, CH4, C2H2, C2H4, C2H6, CO)
print(f"[Doernenburg]     {doern_fault}")

# IEC 60599
iec_fault, iec_ratios = iec_60599(H2, CH4, C2H2, C2H4, C2H6)
print(f"[IEC 60599]       {iec_fault}")

# Key Gas
key_gas, key_val, key_interp, all_key_gases = key_gas_method(H2, CH4, C2H2, C2H4, C2H6, CO)
print(f"[Key Gas]         Dominant: {key_gas} = {key_val:.2f} ppm → {key_interp}")

# TDCG
tdcg_level, tdcg_status, tdcg_action, tdcg_color = tdcg_condition(TDCG)
print(f"[TDCG]            {TDCG:.2f} ppm → Level {tdcg_level}: {tdcg_status}")

# CO2/CO
co2co_ratio, co2co_interp = co2_co_ratio(CO2, CO)
if co2co_ratio is not None:
    print(f"[CO₂/CO Ratio]    {co2co_ratio:.2f} → {co2co_interp}")
else:
    print(f"[CO₂/CO Ratio]    N/A → {co2co_interp}")

# Severity scores for radar (0=normal, 1=critical)
severity_map = {
    "Normal Aging":0.05, "Partial Discharge (low energy density)":0.3,
    "Partial Discharge (high energy density / tracking)":0.5,
    "Low Energy Discharge (arcing, corona)":0.6,
    "High Energy Discharge (arcing)":0.9,
    "Thermal Fault < 150°C (slight overheating)":0.2,
    "Thermal Fault 150–200°C":0.35, "Thermal Fault 200–300°C":0.45,
    "Thermal Fault > 300°C (conductor overheating)":0.65,
}
duval_sev  = {"PD":0.4,"T1":0.2,"T2":0.45,"T3":0.7,"D1":0.55,"D2":0.85,"DT":0.75}
tdcg_sev   = {1:0.05, 2:0.35, 3:0.65, 4:1.0}
radar_scores = [
    duval_sev.get(duval_zone, 0.1),
    severity_map.get(rogers_fault, 0.1),
    0.1 if not doern_valid else (0.7 if "Arcing" in doern_fault else 0.1),
    0.1 if "Normal" in iec_fault or "stray" in iec_fault else 0.5,
    0.1 if key_gas in ("C₂H₆","CH₄") else (0.8 if key_gas=="C₂H₂" else 0.3),
    tdcg_sev.get(tdcg_level, 0.1),
]
radar_labels = ["Duval\nTriangle","Rogers\nRatio","Doernenburg",
                "IEC 60599","Key Gas\nMethod","TDCG\nIEEE C57.104"]


# ─────────────────────────────────────────────────────────────────────────────
# 5. PLOT THE DUVAL TRIANGLE
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 9), facecolor="white")
ax.set_aspect("equal")
ax.axis("off")

# Draw filled zone polygons
for zone_name, zone_data in zones.items():
    pts_cart = [ternary_to_cart(*p) for p in zone_data["points"]]
    poly = Polygon(pts_cart, closed=True,
                   facecolor=zone_data["color"], edgecolor="white",
                   linewidth=1.0, alpha=0.85, zorder=1)
    ax.add_patch(poly)

    # Zone label at centroid
    cx = np.mean([p[0] for p in pts_cart])
    cy = np.mean([p[1] for p in pts_cart])
    ax.text(cx, cy, zone_data["label"], ha="center", va="center",
            fontsize=7.5, fontweight="bold", color="#222222",
            zorder=4, linespacing=1.3)

# Triangle outline
triangle = Polygon([(0, 0), (1, 0), (0.5, np.sqrt(3)/2)],
                   closed=True, fill=False, edgecolor="black", linewidth=2, zorder=5)
ax.add_patch(triangle)

# ─────────────────────────────────────────────────────────────────────────────
# 6. AXIS TICK LINES & LABELS (grid lines at 20% intervals)
# ─────────────────────────────────────────────────────────────────────────────
tick_vals = [20, 40, 60, 80]

for tv in tick_vals:
    # CH4 axis lines (parallel to bottom-right edge)
    p1 = ternary_to_cart(tv, 0, 100-tv)
    p2 = ternary_to_cart(tv, 100-tv, 0)
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color="gray", lw=0.4, ls="--", zorder=2)

    # C2H4 axis lines
    p1 = ternary_to_cart(0, tv, 100-tv)
    p2 = ternary_to_cart(100-tv, tv, 0)
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color="gray", lw=0.4, ls="--", zorder=2)

    # C2H2 axis lines
    p1 = ternary_to_cart(0, 100-tv, tv)
    p2 = ternary_to_cart(100-tv, 0, tv)
    ax.plot([p1[0], p2[0]], [p1[1], p2[1]], color="gray", lw=0.4, ls="--", zorder=2)

# Axis corner labels
offset = 0.04
ax.text(-offset,    -offset,    "100%\nCH₄",  ha="center", va="top",    fontsize=9, fontweight="bold")
ax.text(1+offset,   -offset,    "100%\nC₂H₄", ha="center", va="top",    fontsize=9, fontweight="bold")
ax.text(0.5, np.sqrt(3)/2+0.02, "100%\nC₂H₂", ha="center", va="bottom", fontsize=9, fontweight="bold")

# Axis percentage ticks on bottom
for tv in tick_vals:
    x, y = ternary_to_cart(100-tv, tv, 0)
    ax.text(x, y - 0.04, f"{tv}%", ha="center", va="top", fontsize=7, color="#555555")
    x2, y2 = ternary_to_cart(tv, 0, 100-tv)
    ax.text(x2 - 0.03, y2, f"{tv}%", ha="right",  va="center", fontsize=7, color="#555555")
    x3, y3 = ternary_to_cart(0, tv, 100-tv)
    ax.text(x3 + 0.03, y3, f"{tv}%", ha="left",   va="center", fontsize=7, color="#555555")

# ─────────────────────────────────────────────────────────────────────────────
# 7. PLOT SAMPLE POINT
# ─────────────────────────────────────────────────────────────────────────────
sx, sy = ternary_to_cart(pCH4, pC2H4, pC2H2)

# Highlight zone where point falls

ax.scatter([sx], [sy], s=200, color="red", edgecolors="black",
           linewidths=1.5, zorder=10, marker="*")
ax.annotate(
    f"  Sample Point\n  CH₄={pCH4:.1f}%\n  C₂H₄={pC2H4:.1f}%\n  C₂H₂={pC2H2:.1f}%",
    (sx, sy),
    xytext=(sx + 0.12, sy + 0.06),
    fontsize=8,
    color="darkred",
    fontweight="bold",
    arrowprops=dict(arrowstyle="->", color="darkred", lw=1.2),
    bbox=dict(boxstyle="round,pad=0.3", facecolor="white", edgecolor="darkred", alpha=0.85),
    zorder=11
)

# ─────────────────────────────────────────────────────────────────────────────
# 8. TITLE, LEGEND & INFO BOX
# ─────────────────────────────────────────────────────────────────────────────
ax.set_title("Duval's Triangle (IEC 60599)\nTransformer Oil DGA Analysis",
             fontsize=14, fontweight="bold", pad=14)

# Info box
info_text = (
    f"Gas Values (from DGA Report)\n"
    f"CH₄   = {CH4:.2f} ppm  →  {pCH4:.1f}%\n"
    f"C₂H₄  = {C2H4:.2f} ppm  →  {pC2H4:.1f}%\n"
    f"C₂H₂  = {C2H2:.2f} ppm  →  {pC2H2:.1f}%\n"
    f"─────────────────────\n"
    f"⚠ Fault Zone: {fault_zone}\n"
    f"→ {fault_meanings[fault_zone]}"
)
ax.text(1.02, 0.72, info_text, transform=ax.transAxes,
        fontsize=8.5, va="top", ha="left",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="#FFF8DC", edgecolor="#999", alpha=0.95),
        fontfamily="monospace")

# Legend patches
legend_patches = [mpatches.Patch(color=v["color"], label=v["label"].replace("\n", " "))
                  for v in zones.values()]
ax.legend(handles=legend_patches, loc="lower right",
          bbox_to_anchor=(1.35, 0.0), fontsize=7.5, title="Fault Zones",
          framealpha=0.9, title_fontsize=8)

plt.tight_layout()

plt.show()


print("\n" + "="*60)
print("  OVERALL ASSESSMENT")
print("="*60)
print(f"  Severity: {sev_label(overall_sev).upper()} (index={overall_sev:.2f})")
print(f"  Duval   : {duval_zone} — {DUVAL_MEANINGS[duval_zone]}")
print(f"  Rogers  : {rogers_fault}")
print(f"  IEC     : {iec_fault}")
print(f"  TDCG    : Level {tdcg_level} — {tdcg_status}")
print(f"  CO₂/CO  : {co2co_ratio:.2f} → {co2co_interp}" if co2co_ratio else "  CO₂/CO  : N/A")
print("="*60)
