import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import Polygon
import pdfplumber
import re
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os

# ─────────────────────────────────────────────────────────────────────────────
# 1. SELECT PDF FILE
# ─────────────────────────────────────────────────────────────────────────────

Tk().withdraw()

pdf_path = askopenfilename(
    title="Select DGA PDF Report",
    filetypes=[("PDF Files", "*.pdf")]
)

if not pdf_path:
    print("No PDF selected.")
    exit()

if not os.path.exists(pdf_path):
    print("PDF file not found.")
    exit()

# ─────────────────────────────────────────────────────────────────────────────
# 2. READ PDF SAFELY
# ─────────────────────────────────────────────────────────────────────────────

text = ""

try:

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            page_text = page.extract_text()

            if page_text:
                text += page_text + "\n"

except Exception as e:

    print("Error reading PDF:")
    print(e)
    exit()

# ─────────────────────────────────────────────────────────────────────────────
# 3. SAFE EXTRACTION FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def extract_text(pattern, text):

    try:

        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)

        if match:
            return match.group(1).strip()

        return "NOT FOUND"

    except:
        return "NOT FOUND"

def extract_number(pattern, text):

    try:

        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)

        if match:

            value = match.group(1).strip()

            # ND → 0
            if value.upper() == "ND":
                return 0.0

            return float(value)

        return 0.0

    except:
        return 0.0

# ─────────────────────────────────────────────────────────────────────────────
# 4. EXTRACT BASIC DETAILS
# ─────────────────────────────────────────────────────────────────────────────

equipment = extract_text(
    r"Equipment Designation\s+(.+)",
    text
)

serial_no = extract_text(
    r"Manufacturer's Sl\. No\.\s+(\S+)",
    text
)

manufacturer = extract_text(
    r"Manufacturer\s+([A-Z0-9\- ]+)",
    text
)

bdv = extract_number(
    r"Electric Strength \(BDV\).*?\s(\d+)\s+30 Min",
    text
)

water = extract_number(
    r"Water Content By Karl Fischer.*?\s(\d+)\s+40 Max",
    text
)


# ─────────────────────────────────────────────────────────────────────────────
# 5. EXTRACT ALL GASES
# ─────────────────────────────────────────────────────────────────────────────

H2 = extract_number(
    r"Hydrogen.*?IS 9434\s+([\d.]+|ND)",
    text
)

O2 = extract_number(
    r"Oxygen.*?IS 9434\s+([\d.]+|ND)",
    text
)

N2 = extract_number(
    r"Nitrogen.*?IS 9434\s+([\d.]+|ND)",
    text
)

CO = extract_number(
    r"Carbon Monoxide.*?IS 9434\s+([\d.]+|ND)",
    text
)

CH4 = extract_number(
    r"Methane.*?IS 9434\s+([\d.]+|ND)",
    text
)

CO2 = extract_number(
    r"Carbon Dioxide.*?IS 9434\s+([\d.]+|ND)",
    text
)

C2H4 = extract_number(
    r"Ethylene.*?IS 9434\s+([\d.]+|ND)",
    text
)

C2H6 = extract_number(
    r"Ethane.*?IS 9434\s+([\d.]+|ND)",
    text
)

C2H2 = extract_number(
    r"Acetylene.*?IS 9434\s+([\d.]+|ND)",
    text
)

C3H6 = extract_number(
    r"Propylene.*?IS 9434\s+([\d.]+|ND)",
    text
)

C3H8 = extract_number(
    r"Propane.*?IS 9434\s+([\d.]+|ND)",
    text
)

TDCG = extract_number(
    r"Total Dissolved Combustible.*?\s([\d.]+)",
    text
)

# ─────────────────────────────────────────────────────────────────────────────
# 6. PRINT COMPLETE REPORT
# ─────────────────────────────────────────────────────────────────────────────


# ─────────────────────────────────────────────────────────────────────────────
# 7. DUVAL CALCULATIONS
# ─────────────────────────────────────────────────────────────────────────────

total = CH4 + C2H4 + C2H2

pCH4 = (CH4 / total) * 100 if total > 0 else 0
pC2H4 = (C2H4 / total) * 100 if total > 0 else 0
pC2H2 = (C2H2 / total) * 100 if total > 0 else 0

# Duval percentage coordinates
pCH4  = (CH4  / total * 100) if total > 0 else 0
pC2H4 = (C2H4 / total * 100) if total > 0 else 0
pC2H2 = (C2H2 / total * 100) if total > 0 else 0

print(f"CH₄  = {CH4} ppm  → {pCH4:.2f}%")
print(f"C₂H₄ = {C2H4} ppm  → {pC2H4:.2f}%")
print(f"C₂H₂ = {C2H2} ppm  → {pC2H2:.2f}%")
print(f"Sum  = {total} ppm")

# ─────────────────────────────────────────────────────────────────────────────
# 2. TERNARY → CARTESIAN CONVERSION
# ─────────────────────────────────────────────────────────────────────────────
# In Duval's triangle:
#   Bottom-left  corner = 100% CH₄
#   Bottom-right corner = 100% C₂H₄
#   Top          corner = 100% C₂H₂
# We map ternary (a,b,c) where a+b+c=100 to 2D Cartesian

def ternary_to_cart(pCH4, pC2H4, pC2H2):
    """Convert ternary % to Cartesian (x,y) in equilateral triangle."""
    a = pCH4 / 100    # bottom-left axis
    b = pC2H4 / 100   # bottom-right axis
    c = pC2H2 / 100   # top axis
    x = 0.5 * (2 * b + c) / (a + b + c)
    y = (np.sqrt(3) / 2) * c / (a + b + c)
    return x, y

# ─────────────────────────────────────────────────────────────────────────────
# 3. DUVAL ZONE BOUNDARIES (standard IEC 60599 zones)
# ─────────────────────────────────────────────────────────────────────────────
# Zones: PD, T1, T2, T3, D1, D2, DT
# Each zone is a polygon defined in ternary (CH4%, C2H4%, C2H2%) coordinates

zones = {
    "PD": {
        "color": "#B3D9FF",
        "label": "PD\nPartial Discharge",
        "points": [(98,0,2),(100,0,0),(98,2,0)],  # top-left corner small triangle
    },
    "T1": {
        "color": "#FFFF99",
        "label": "T1\n< 300°C",
        "points": [(98,0,2),(98,2,0),(76,24,0),(77,0,23)],
    },
    "T2": {
        "color": "#FFD966",
        "label": "T2\n300–700°C",
        "points": [(77,0,23),(76,24,0),(40,60,0),(46,0,54)],
    },
    "T3": {
        "color": "#FF9900",
        "label": "T3\n> 700°C",
        "points": [(46,0,54),(40,60,0),(0,100,0),(0,93,7),(0,0,100)],
    },
    "D1": {
        "color": "#FF7F7F",
        "label": "D1\nLow Energy\nDischarge",
        "points": [(100,0,0),(98,2,0),(76,24,0),(87,0,13)],
    },
    "D2": {
        "color": "#FF3333",
        "label": "D2\nHigh Energy\nDischarge",
        "points": [(87,0,13),(76,24,0),(40,60,0),(23,0,77)],
    },
    "DT": {
        "color": "#CC66FF",
        "label": "DT\nMixed Discharge\n+ Thermal",
        "points": [(23,0,77),(40,60,0),(0,93,7),(0,0,100)],
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# 4. DETERMINE FAULT ZONE FOR THIS SAMPLE
# ─────────────────────────────────────────────────────────────────────────────
def classify_duval(pCH4, pC2H4, pC2H2):
    """Standard Duval triangle fault classification."""
    # Using standard boundary rules
    if pC2H2 >= 29:
        if pC2H4 >= 0 and pCH4 <= 23:
            return "DT"
        return "D2"
    if pC2H2 >= 13:
        if pC2H4 <= 60:
            return "D2" if pCH4 <= 23 else "D2"
        return "DT"
    if pC2H2 >= 2:
        if pC2H4 < 24:
            return "T1"  # overlapping with D1 region
        return "D1" if pCH4 >= 87 else "D2"
    # pC2H2 < 2
    if pC2H4 >= 60:
        return "T3"
    if pC2H4 >= 24:
        return "T2"
    if pCH4 >= 98:
        return "PD"
    return "T1"

fault_zone = classify_duval(pCH4, pC2H4, pC2H2)

fault_meanings = {
    "PD": "Partial Discharge",
    "T1": "Thermal Fault < 300°C",
    "T2": "Thermal Fault 300–700°C",
    "T3": "Thermal Fault > 700°C",
    "D1": "Low Energy Electrical Discharge",
    "D2": "High Energy Electrical Discharge (Arc)",
    "DT": "Mixed Discharge + Thermal Fault",
}

# ─────────────────────────────────────────────────────────────────────────────
# 5. PLOT THE DUVAL TRIANGLE
# ─────────────────────────────────────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(17, 10), facecolor="white")
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
zone_color = zones.get(fault_zone, {}).get("color", "yellow")
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

# ─────────────────────────────────────────────────────────────────────────────
# COMPLETE REPORT PANEL
# ─────────────────────────────────────────────────────────────────────────────

info_text = (
    "TRANSFORMER DGA PDF ANALYSIS\n"
    "══════════════════════════════\n\n"

    "BASIC DETAILS\n"
    "────────────────────────────\n"
    f"Equipment : {equipment}\n"
    f"Serial No : {serial_no}\n"
    f"Manufacturer : {manufacturer}\n\n"

    "OIL PARAMETERS\n"
    "────────────────────────────\n"
    f"BDV            : {bdv:.2f} kV\n"
    f"Water Content  : {water:.2f} ppm\n\n"

    "DGA GASES\n"
    "────────────────────────────\n"
    f"H₂     : {H2:.2f} ppm\n"
    f"O₂     : {O2:.2f} ppm\n"
    f"N₂     : {N2:.2f} ppm\n"
    f"CO     : {CO:.2f} ppm\n"
    f"CH₄    : {CH4:.2f} ppm\n"
    f"CO₂    : {CO2:.2f} ppm\n"
    f"C₂H₄   : {C2H4:.2f} ppm\n"
    f"C₂H₆   : {C2H6:.2f} ppm\n"
    f"C₂H₂   : {C2H2:.2f} ppm\n"
    f"C₃H₆   : {C3H6:.2f} ppm\n"
    f"C₃H₈   : {C3H8:.2f} ppm\n\n"

    f"TDCG   : {TDCG:.2f} ppm\n\n"

    "DUVAL PERCENTAGES\n"
    "────────────────────────────\n"
    f"CH₄   → {pCH4:.2f}%\n"
    f"C₂H₄ → {pC2H4:.2f}%\n"
    f"C₂H₂ → {pC2H2:.2f}%\n\n"

    "FAULT ANALYSIS\n"
    "────────────────────────────\n"
    f"Zone : {fault_zone}\n"
    f"{fault_meanings[fault_zone]}"
)

ax.text(
    1.02,
    0.98,
    info_text,
    transform=ax.transAxes,
    fontsize=8,
    va="top",
    ha="left",
    bbox=dict(
        boxstyle="round,pad=0.6",
        facecolor="#FFF8DC",
        edgecolor="#666666",
        alpha=0.98
    ),
    fontfamily="monospace"
)
# Legend patches
legend_patches = [mpatches.Patch(color=v["color"], label=v["label"].replace("\n", " "))
                  for v in zones.values()]
ax.legend(handles=legend_patches, loc="lower right",
          bbox_to_anchor=(1.35, 0.0), fontsize=7.5, title="Fault Zones",
          framealpha=0.9, title_fontsize=8)

plt.tight_layout()

plt.show()