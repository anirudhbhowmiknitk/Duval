"""
Transformer DGA Analysis — Duval Triangle (IEC 60599)
======================================================
Run with:   python dga_duval_from_pdf.py

A file-picker dialog opens. Select either:
  • A PDF  report  (.pdf)
  • An Excel report (.xlsx / .xls)

The script extracts all DGA gas values + metadata, plots the
Duval Triangle in reference-chart style, and saves a PNG image
next to the source file.

Dependencies:
    pip install pypdf openpyxl pandas matplotlib numpy
"""

import os
import re
import sys
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.patches import Polygon, FancyBboxPatch

# ── file-picker ──────────────────────────────────────────────────────────────
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# ── PDF reader ───────────────────────────────────────────────────────────────
try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

# ── Excel reader ─────────────────────────────────────────────────────────────
try:
    import pandas as pd
    _PANDAS = True
except ImportError:
    _PANDAS = False


# ═════════════════════════════════════════════════════════════════════════════
# FILE READING
# ═════════════════════════════════════════════════════════════════════════════

def read_pdf(path):
    """Return raw text from every page of a PDF."""
    if PdfReader is None:
        raise ImportError("pypdf is not installed.  Run: pip install pypdf")
    reader = PdfReader(path)
    pages  = []
    for page in reader.pages:
        t = page.extract_text()
        if t:
            pages.append(t)
    return "\n".join(pages)


def read_excel(path):
    """
    Return (text_blob, dataframe_dict) from an Excel workbook.
    text_blob  – everything stringified for regex scanning
    df_dict    – {sheet_name: DataFrame} for structured lookup
    """
    if not _PANDAS:
        raise ImportError("pandas is not installed.  Run: pip install pandas openpyxl")
    xl    = pd.ExcelFile(path, engine="openpyxl")
    blobs = []
    dfs   = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=None)
        dfs[sheet] = df
        blobs.append(df.to_string(index=False, header=False, na_rep=""))
    return "\n".join(blobs), dfs


# ═════════════════════════════════════════════════════════════════════════════
# VALUE EXTRACTION  (PDF / text-blob)
# ═════════════════════════════════════════════════════════════════════════════


def load_file(path):

    ext = os.path.splitext(path)[1].lower()

    if ext == ".pdf":
        text = read_pdf(path)

    elif ext in [".xlsx", ".xls", ".xlsm"]:
        text, _ = read_excel(path)

    else:
        raise ValueError("Unsupported file type")

    # ─────────────────────────────────────────────
    # SAFE GAS EXTRACTION
    # ─────────────────────────────────────────────

    def extract_gas_value(gas_name):

        for line in text.splitlines():

            if gas_name.lower() in line.lower():

                parts = line.split()

                for i, part in enumerate(parts):

                    if part == "9434":

                        if i + 1 < len(parts):

                            value = parts[i + 1]

                            if value.upper() == "ND":
                                return 0.0

                            try:
                                return float(value)

                            except:
                                return 0.0

        return 0.0

    # ─────────────────────────────────────────────
    # GENERIC EXTRACTION FUNCTIONS
    # ─────────────────────────────────────────────

    def extract_text(pattern, text):

        try:

            match = re.search(
                pattern,
                text,
                re.DOTALL | re.IGNORECASE
            )

            if match:
                return match.group(1).strip()

            return "NOT FOUND"

        except:
            return "NOT FOUND"

    def extract_number(pattern, text):

        try:

            match = re.search(
                pattern,
                text,
                re.DOTALL | re.IGNORECASE
            )

            if match:

                value = match.group(1).strip()

                if value.upper() == "ND":
                    return 0.0

                return float(value)

            return 0.0

        except:
            return 0.0
# ═════════════════════════════════════════════════════════════════════════════
# DUVAL TRIANGLE  —  IEC 60599
# ═════════════════════════════════════════════════════════════════════════════

# Zone boundary points in (CH4%, C2H4%, C2H2%) ternary coords
DUVAL_ZONES = {
    "PD": {
        "color": "#18182A", "tc": "white",
        "points": [(98,0,2),(100,0,0),(98,2,0)],
    },
    "T1": {
        "color": "#D4740A", "tc": "white",
        "points": [(98,0,2),(98,2,0),(76,24,0),(77,0,23)],
    },
    "T2": {
        "color": "#7B3A10", "tc": "white",
        "points": [(77,0,23),(76,24,0),(40,60,0),(46,0,54)],
    },
    "T3": {
        "color": "#4A1E00", "tc": "white",
        "points": [(46,0,54),(40,60,0),(0,100,0),(0,93,7),(0,0,100)],
    },
    "D1": {
        "color": "#A8D8C0", "tc": "#111111",
        "points": [(100,0,0),(98,2,0),(76,24,0),(87,0,13)],
    },
    "D2": {
        "color": "#4AAFC8", "tc": "white",
        "points": [(87,0,13),(76,24,0),(40,60,0),(23,0,77)],
    },
    "DT": {
        "color": "#2A8BA8", "tc": "white",
        "points": [(23,0,77),(40,60,0),(0,93,7),(0,0,100)],
    },
    "NORMAL": {
    "color": "#666666",
    "tc": "white",
    "points": [(98,0,2),(100,0,0),(98,2,0)],
},
}

DUVAL_MEANINGS = {
    "PD": "Partial Discharge",
    "T1": "Thermal Fault < 300°C",
    "T2": "Thermal Fault 300–700°C",
    "T3": "Thermal Fault > 700°C",
    "D1": "Low Energy Electrical Discharge",
    "D2": "High Energy Discharge (Arc)",
    "DT": "Electrical and Thermal",
    "NORMAL": "No Significant Fault / Trace Gas",
}

LEGEND_LABELS = {
    "PD": "Partial Discharge",
    "T1": "Thermal Fault less than 300°C",
    "T2": "Thermal Fault between 300°C and 700°C",
    "T3": "Thermal Fault greater than 700°C",
    "D1": "Low Energy Discharge (Sparking)",
    "D2": "High Energy Discharge (Arcing)",
    "DT": "Mix of Thermal and Electrical Faults",
}


def t2c(ch4, c2h4, c2h2):
    """Ternary (CH4%, C2H4%, C2H2%) → Cartesian (x, y)."""
    tot = ch4 + c2h4 + c2h2
    if tot == 0:
        return 0.5, np.sqrt(3) / 6
    b = c2h4 / tot
    c = c2h2 / tot
    return 0.5 * (2*b + c), (np.sqrt(3)/2) * c


def classify(pCH4, pC2H4, pC2H2):

    # PD
    if pCH4 >= 98:
        return "PD"

    # T1
    if pC2H2 < 2 and pC2H4 < 20:
        return "T1"

    # T2
    if pC2H2 < 5 and 20 <= pC2H4 < 50:
        return "T2"

    # T3
    if pC2H2 < 15 and pC2H4 >= 50:
        return "T3"

    # D1
    if 2 <= pC2H2 < 15:
        return "D1"

    # D2
    if pC2H2 >= 15 and pCH4 >= 20:
        return "D2"

    # DT
    if pC2H2 >= 15 and pC2H4 >= 20:
        return "DT"

    return "T1"


# ═════════════════════════════════════════════════════════════════════════════
# ADDITIONAL ANALYSIS METHODS
# ═════════════════════════════════════════════════════════════════════════════

def tdcg_level(v):
    if v <= 720:  return 1, "Normal Operation"
    if v <= 1920: return 2, "Caution"
    if v <= 4630: return 3, "Warning"
    return 4, "Critical"

def rogers(h2, ch4, c2h2, c2h4, c2h6):
    r1 = c2h2/c2h4 if c2h4 > 0 else 0
    r2 = ch4 /h2   if h2   > 0 else 0
    r5 = c2h4/c2h6 if c2h6 > 0 else 0
    def code(v, lo, hi): return 0 if v < lo else (1 if v < hi else 2)
    key = (code(r1,0.1,3), code(r2,0.1,1), code(r5,1,3))
    return {
        (0,0,0): "Normal Aging",
        (0,1,0): "Partial Discharge (low energy)",
        (1,1,0): "Partial Discharge (high energy)",
        (1,0,1): "Low Energy Discharge (arcing)",
        (1,0,2): "High Energy Discharge (arcing)",
        (0,0,1): "Thermal Fault < 150°C",
        (0,0,2): "Thermal Fault 150–200°C",
        (0,1,1): "Thermal Fault 200–300°C",
        (0,1,2): "Thermal Fault > 300°C",
    }.get(key, f"Unclassified {key}")

def iec60599(h2, ch4, c2h2, c2h4, c2h6):
    total = h2 + ch4 + c2h2 + c2h4 + c2h6
    if total < 50:
        return "Normal / Trace Gas"
    r1 = c2h2/c2h4 if c2h4 > 0 else 0
    r2 = ch4 /h2   if h2   > 0 else 0
    if r1 < 0.1:
        if r2 < 0.1:   return "Partial Discharge"
        if r2 < 1.0:   return "Thermal Fault (low temp)"
        return             "Thermal Fault (>300°C)"
    if r1 < 3.0:
        return "Low Energy Discharge" if r2 < 0.1 else "High Energy Discharge"
    return "High Energy Discharge (severe arcing)"


# ═════════════════════════════════════════════════════════════════════════════
# PLOT
# ═════════════════════════════════════════════════════════════════════════════

def draw(gases, meta, limits, out_path):
    g = lambda k: gases.get(k) or 0.0

    H2   = g("H2");  O2   = g("O2");  N2  = g("N2")
    CO   = g("CO");  CH4  = g("CH4"); CO2 = g("CO2")
    C2H4 = g("C2H4"); C2H6 = g("C2H6"); C2H2 = g("C2H2")
    C3H6 = g("C3H6"); C3H8 = g("C3H8")
    BDV  = gases.get("BDV")
    WC   = gases.get("Water")

    tri_sum = CH4 + C2H4 + C2H2
    # Very low gas generation → treat as healthy
    TDCG = gases.get("TDCG") or (
        H2 + CO + CH4 + C2H4 + C2H6 + C2H2
    )
    if tri_sum == 0:
        pCH4, pC2H4, pC2H2 = 33.3, 33.3, 33.3
    else:
        pCH4  = CH4  / tri_sum * 100
        pC2H4 = C2H4 / tri_sum * 100
        pC2H2 = C2H2 / tri_sum * 100
    if tri_sum < 10 and TDCG < 100:

        zone = "NORMAL"
        zone_name = "No Significant Fault / Trace Gas"

    else:

        zone = classify(pCH4, pC2H4, pC2H2)
        zone_name = DUVAL_MEANINGS[zone]
    
    
    
    
    tdcg_lv, tdcg_st = tdcg_level(TDCG)
    rog        = rogers(H2, CH4, C2H2, C2H4, C2H6)
    iec        = iec60599(H2, CH4, C2H2, C2H4, C2H6)
    co2co      = CO2/CO if CO > 0 else None
        if co2co is not None:

        if co2co < 3:
            co2co_text = "Possible Paper Degradation"

        elif co2co > 10:
            co2co_text = "Normal"

        else:
            co2co_text = "Moderate Aging"

    else:

        co2co_text = "N/A"

    dz = DUVAL_ZONES[zone]

    # ── figure ───────────────────────────────────────────────────────────────
    fig = plt.figure(figsize=(22, 13), facecolor="white")
    gs  = gridspec.GridSpec(1, 2, width_ratios=[1.5, 1], wspace=0.04)
    axT = fig.add_subplot(gs[0])   # triangle
    axI = fig.add_subplot(gs[1])   # info
    for ax in (axT, axI):
        ax.set_facecolor("white")
        ax.axis("off")

    # ── triangle axes ─────────────────────────────────────────────────────────
    axT.set_aspect("equal")
    axT.set_xlim(-0.34, 1.34)
    axT.set_ylim(-0.30, 1.10)
    H_tri = np.sqrt(3) / 2

    # filled zones
    for zname, zd in DUVAL_ZONES.items():
        pts = [t2c(*p) for p in zd["points"]]
        lw  = 3.5 if zname == zone else 1.2
        ec  = "red" if zname == zone else "#2A2A2A"
        axT.add_patch(Polygon(pts, closed=True,
                               facecolor=zd["color"], edgecolor=ec,
                               linewidth=lw, zorder=2))
        cx = np.mean([p[0] for p in pts])
        cy = np.mean([p[1] for p in pts])
        axT.text(cx, cy, zname,
                 ha="center", va="center",
                 fontsize=12, fontweight="bold", color=zd["tc"], zorder=6,
                 bbox=dict(boxstyle="square,pad=0.25",
                           facecolor=zd["color"],
                           edgecolor="#111111", linewidth=1.4))

    # outer border (drawn on top so it covers zone edges cleanly)
    axT.add_patch(Polygon([(0,0),(1,0),(0.5,H_tri)],
                           closed=True, fill=False,
                           edgecolor="black", linewidth=4.5, zorder=7))

    # grid lines + tick marks + labels
    for tv in [20, 40, 60, 80]:
        tk = 0.020   # tick length

        # ── left edge (CH4 axis) ──
        xl, yl = t2c(tv, 0, 100-tv)
        # tick inward (rotated 60°)
        axT.plot([xl, xl + tk*np.cos(np.radians(0))],
                 [yl, yl + tk*np.sin(np.radians(0))],
                 color="black", lw=1.6, zorder=8)
        axT.text(xl - 0.06, yl, str(tv),
                 ha="right", va="center",
                 fontsize=10, fontweight="bold", color="black")

        # ── right edge (C2H4 axis) ──
        xr, yr = t2c(0, 100-tv, tv)
        axT.plot([xr, xr - tk*np.cos(np.radians(0))],
                 [yr, yr + tk*np.sin(np.radians(0))],
                 color="black", lw=1.6, zorder=8)
        axT.text(xr + 0.06, yr, str(tv),
                 ha="left", va="center",
                 fontsize=10, fontweight="bold", color="black")

        # ── bottom edge (C2H2 axis) ──
        xb, yb = t2c(100-tv, 0, tv)
        axT.plot([xb, xb], [yb, yb + tk],
                 color="black", lw=1.6, zorder=8)
        axT.text(xb, yb - 0.045, str(tv),
                 ha="center", va="top",
                 fontsize=10, fontweight="bold", color="black")

        # dashed interior grid lines (all 3 families)
        for p1, p2 in [
            (t2c(tv,   0, 100-tv), t2c(tv,  100-tv, 0)),
            (t2c(0,   tv, 100-tv), t2c(100-tv, tv,  0)),
            (t2c(0, 100-tv,  tv),  t2c(100-tv,  0, tv)),
        ]:
            axT.plot([p1[0],p2[0]], [p1[1],p2[1]],
                     color="black", lw=0.5, ls="--", alpha=0.30, zorder=3)

    # ── axis arrows + rotated labels ────────────────────────────────────────
    # CH4: left edge, arrow pointing up-left (toward apex)
    axT.annotate("",
        xy=t2c(72, 0, 28), xytext=t2c(22, 0, 78),
        arrowprops=dict(arrowstyle="-|>", color="black",
                        lw=2.2, mutation_scale=18), zorder=9)
    mx, my = t2c(47, 0, 53)
    axT.text(mx - 0.14, my + 0.01, "% CH₄",
             ha="center", va="center", fontsize=13, fontweight="bold",
             color="black", rotation=60)

    # C2H4: right edge, arrow pointing down (toward bottom-right)
    axT.annotate("",
        xy=t2c(0, 78, 22), xytext=t2c(0, 28, 72),
        arrowprops=dict(arrowstyle="-|>", color="black",
                        lw=2.2, mutation_scale=18), zorder=9)
    rx, ry = t2c(0, 53, 47)
    axT.text(rx + 0.14, ry + 0.01, "% C₂H₄",
             ha="center", va="center", fontsize=13, fontweight="bold",
             color="black", rotation=-60)

    # C2H2: bottom edge, arrow pointing left
    axT.annotate("",
        xy=(0.16, -0.10), xytext=(0.84, -0.10),
        arrowprops=dict(arrowstyle="-|>", color="black",
                        lw=2.2, mutation_scale=18), zorder=9)
    axT.text(0.5, -0.155, "% C₂H₂",
             ha="center", va="top", fontsize=13, fontweight="bold", color="black")

    # ── corner labels ────────────────────────────────────────────────────────
    axT.text(0.5, H_tri + 0.06, "C₂H₂",
             ha="center", va="bottom", fontsize=14, fontweight="bold", color="black")
    axT.text(-0.03, -0.07, "CH₄",
             ha="right", va="top", fontsize=14, fontweight="bold", color="black")
    axT.text(1.03, -0.07, "C₂H₄",
             ha="left", va="top", fontsize=14, fontweight="bold", color="black")

    # ── sample point ─────────────────────────────────────────────────────────
    sx, sy = t2c(pCH4, pC2H4, pC2H2)
    axT.scatter([sx], [sy], s=320, color="red", edgecolors="black",
                linewidths=2.2, zorder=14, marker="*")

    adx = 0.15 if sx < 0.62 else -0.22
    ady = 0.10 if sy < 0.52 else -0.13
    axT.annotate(
        f"  CH₄  = {pCH4:.2f}%\n  C₂H₄ = {pC2H4:.2f}%\n  C₂H₂ = {pC2H2:.2f}%",
        (sx, sy), xytext=(sx+adx, sy+ady),
        fontsize=10, color="#1A1A1A", fontweight="bold",
        arrowprops=dict(arrowstyle="-|>", color="red", lw=1.8, mutation_scale=12),
        bbox=dict(boxstyle="round,pad=0.45", facecolor="white",
                  edgecolor="red", linewidth=2.0, alpha=0.97),
        zorder=15)

    # ── legend (top-left, outside triangle) ─────────────────────────────────
    lx, ly = -0.32, 1.00
    for code, desc in LEGEND_LABELS.items():
        col = DUVAL_ZONES[code]["color"]
        axT.add_patch(plt.Rectangle(
            (lx, ly-0.015), 0.028, 0.027,
            color=col, ec="black", lw=0.9,
            transform=axT.transData, zorder=10))
        axT.text(lx+0.038, ly+0.001, f"{code} = {desc}",
                 ha="left", va="center",
                 fontsize=8.5, color="black",
                 transform=axT.transData)
        ly -= 0.060

    # ── diagnosis strip ──────────────────────────────────────────────────────
    axT.text(0.5, -0.225,
             f"Diagnosis Result :   {zone}  =  {zone_name}",
             ha="center", va="center",
             fontsize=12.5, fontweight="bold", color=dz["tc"],
             transform=axT.transData, zorder=16,
             bbox=dict(boxstyle="round,pad=0.55",
                       facecolor=dz["color"],
                       edgecolor="black", linewidth=2.2))

    # ── title ─────────────────────────────────────────────────────────────────
    axT.set_title("Duval's Triangle DGA  —  IEC 60599",
                  fontsize=17, fontweight="bold", color="black", pad=12)

    # ══════════════════════════════════════════════════════════════════════════
    # INFO PANEL
    # ══════════════════════════════════════════════════════════════════════════

    def section(ax, y, title):
        ax.plot([0.02, 0.98], [y+0.006, y+0.006],
                color="#AAAAAA", lw=0.9, transform=ax.transAxes)
        ax.text(0.50, y-0.013, title,
                ha="center", va="top",
                fontsize=10, color="#1A4488", fontweight="bold",
                transform=ax.transAxes)

    def row(ax, y, label, value, unit="", hi=False, warn=False):
        lc = "#444444"
        vc = "#7A3800" if hi else ("#AA0000" if warn else "#111111")
        fw = "bold" if (hi or warn) else "normal"
        ax.text(0.04, y, label,
                ha="left", va="center", fontsize=9, color=lc,
                transform=ax.transAxes)
        ax.text(0.97, y, f"{value} {unit}".strip(),
                ha="right", va="center", fontsize=9,
                color=vc, fontweight=fw, transform=ax.transAxes)

    y = 0.975

    axI.text(0.5, y, "TRANSFORMER DGA PDF ANALYSIS",
             ha="center", va="top", fontsize=13, fontweight="bold",
             color="#111111", transform=axI.transAxes)
    y -= 0.055

    # ── basic details ─────────────────────────────────────────────────────────
    section(axI, y, "──────── BASIC DETAILS ────────");  y -= 0.053
    row(axI, y, "Equipment:",              meta.get("equipment","N/A")[:40]); y -= 0.038
    row(axI, y, "Manufacturer Serial No:", meta.get("serial","N/A"));         y -= 0.038
    row(axI, y, "Manufacturer:",           meta.get("manufacturer","N/A")[:35]); y -= 0.038
    row(axI, y, "Location:",               meta.get("location","N/A")[:40]); y -= 0.038
    row(axI, y, "Date:",                   meta.get("date","N/A"));           y -= 0.038
    row(axI, y, "Voltage Rating:",         meta.get("voltage","N/A"));        y -= 0.038
    row(axI, y, "MVA Rating:",             meta.get("mva","N/A"));            y -= 0.042

    # ── oil parameters ────────────────────────────────────────────────────────
    section(axI, y, "──────── OIL PARAMETERS ────────"); y -= 0.053
    if BDV is not None:
        row(axI, y, "BDV:", f"{BDV:.1f}", "kV", warn=BDV < limits.get("BDV_MIN", 40)); y -= 0.038
    else:
        row(axI, y, "BDV:", "N/A"); y -= 0.038
    if WC is not None:
        row(axI, y, "Water Content:", f"{WC:.1f}", "ppm", warn=WC > limits.get("Water_MAX", 30)); y -= 0.038
    else:
        row(axI, y, "Water Content:", "N/A"); y -= 0.038
    y -= 0.004
    # ── DGA gases ────────────────────────────────────────────────────────────
    section(axI, y, "──────── DGA GASES ────────"); y -= 0.053
    gas_rows = [
    ("H₂",   H2,   limits.get("H2_MAX")),
    ("O₂",   O2,   None),
    ("N₂",   N2,   None),
    ("CO",   CO,   limits.get("CO_MAX")),
    ("CH₄",  CH4,  limits.get("CH4_MAX")),
    ("CO₂",  CO2,  limits.get("CO2_MAX")),
    ("C₂H₄", C2H4, limits.get("C2H4_MAX")),
    ("C₂H₆", C2H6, limits.get("C2H6_MAX")),
    ("C₂H₂", C2H2, limits.get("C2H2_MAX")),
    ("C₃H₆", C3H6, None),
    ("C₃H₈", C3H8, None),
]
    for name, val, lim in gas_rows:
        over = lim is not None and val > lim
        limit_text = f"(Limit: {lim})" if lim is not None else ""
        row(
            axI,
            y,
            f"  {name}",
            f"{val:.2f} ppm {limit_text}",
            "",
            warn=over
        )
        y -= 0.035
    row(axI, y, "TDCG", f"{TDCG:.2f}", "ppm", hi=True); y -= 0.042

    # ── Duval percentages ─────────────────────────────────────────────────────
    section(axI, y, "──────── DUVAL PERCENTAGES ────────"); y -= 0.053
    row(axI, y, "  CH₄",  f"{CH4:.2f} ppm",  f"→ {pCH4:.2f}%",  hi=True); y -= 0.035
    row(axI, y, "  C₂H₄", f"{C2H4:.2f} ppm", f"→ {pC2H4:.2f}%", hi=True); y -= 0.035
    row(axI, y, "  C₂H₂", f"{C2H2:.2f} ppm", f"→ {pC2H2:.2f}%", hi=True); y -= 0.035
    row(axI, y, "  Sum",   f"{tri_sum:.4f}",  "ppm");                       y -= 0.042

    # ── supporting methods ───────────────────────────────────────────────────
    section(axI, y, "──────── SUPPORTING METHODS ────────"); y -= 0.053
    row(axI, y, "  Rogers Ratio:", rog[:38]); y -= 0.035
    row(axI, y, "  IEC 60599:", iec[:38]); y -= 0.035

    if co2co is not None:

        row(
            axI,
            y,
            "  CO₂/CO Ratio:",
            f"{co2co:.2f}"
        )
        y -= 0.030

        row(
            axI,
            y,
            "  Interpretation:",
            co2co_text,
            warn=co2co < 3
        )
        y -= 0.040

    row(
        axI,
        y,
        "  TDCG Level:",
        f"Level {tdcg_lv}  —  {tdcg_st}",
        warn=tdcg_lv >= 3,
        hi=tdcg_lv < 3
    )
    y -= 0.050

    # ── diagnosis box ────────────────────────────────────────────────────────
    bh = 0.092
    axI.add_patch(FancyBboxPatch(
        (0.03, y - bh), 0.94, bh,
        boxstyle="round,pad=0.01",
        transform=axI.transAxes,
        facecolor=dz["color"], edgecolor="black",
        linewidth=2.8, zorder=5))
    axI.text(0.50, y - 0.027, "DIAGNOSIS RESULT",
             ha="center", va="center",
             fontsize=9.5, color=dz["tc"], fontweight="bold",
             transform=axI.transAxes)
    axI.text(0.50, y - 0.067, f"{zone}  =  {zone_name}",
             ha="center", va="center",
             fontsize=12.5, color=dz["tc"], fontweight="bold",
             transform=axI.transAxes)

    # ── save ─────────────────────────────────────────────────────────────────
    plt.savefig(out_path, dpi=300, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"\n✅  Saved → {out_path}")


# ═════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════════

def main():
    # ── open file picker ─────────────────────────────────────────────────────
    root = Tk()
    root.withdraw()                    # hide the empty Tk window
    root.attributes("-topmost", True)  # dialog appears on top

    file_path = askopenfilename(
        title="Select DGA Report (PDF or Excel)",
        filetypes=[
            ("Supported files", "*.pdf *.xlsx *.xls *.xlsm"),
            ("PDF files",       "*.pdf"),
            ("Excel files",     "*.xlsx *.xls *.xlsm"),
            ("All files",       "*.*"),
        ]
    )
    root.destroy()

    if not file_path:
        print("No file selected. Exiting.")
        sys.exit(0)

    print(f"\nLoading: {file_path}")

    # ── extract data ─────────────────────────────────────────────────────────
    gases, meta, limits = load_file(file_path)

    # ── print what was found ─────────────────────────────────────────────────
    print("\n══════════════════════════════════════════════")
    print("         TRANSFORMER DGA PDF ANALYSIS")
    print("══════════════════════════════════════════════")
    print("\n──────── BASIC DETAILS ────────")
    print(f"  Equipment            : {meta.get('equipment','N/A')}")
    print(f"  Manufacturer Serial  : {meta.get('serial','N/A')}")
    print(f"  Manufacturer         : {meta.get('manufacturer','N/A')}")
    print(f"  Location             : {meta.get('location','N/A')}")
    print(f"  Date                 : {meta.get('date','N/A')}")

    print("\n──────── OIL PARAMETERS ────────")
    print(f"  BDV                  : {gases.get('BDV','N/A')} kV")
    print(f"  Water Content        : {gases.get('Water','N/A')} ppm")

    print("\n──────── DGA GASES ────────")
    gas_display = [
        ("H2","H₂"),("O2","O₂"),("N2","N₂"),
        ("CO","CO"),("CH4","CH₄"),("CO2","CO₂"),
        ("C2H4","C₂H₄"),("C2H6","C₂H₆"),("C2H2","C₂H₂"),
        ("C3H6","C₃H₆"),("C3H8","C₃H₈"),
    ]
    for key, label in gas_display:
        v = gases.get(key)
        print(f"  {label:6s}  =  {v:.4f} ppm" if v is not None else f"  {label:6s}  =  N/A")

    CH4  = gases.get("CH4")  or 0.0
    C2H4 = gases.get("C2H4") or 0.0
    C2H2 = gases.get("C2H2") or 0.0
    tri_sum = CH4 + C2H4 + C2H2
    if tri_sum > 0:
        print(f"\n  CH₄  = {CH4:.2f} ppm  → {CH4/tri_sum*100:.2f}%")
        print(f"  C₂H₄ = {C2H4:.2f} ppm  → {C2H4/tri_sum*100:.2f}%")
        print(f"  C₂H₂ = {C2H2:.2f} ppm  → {C2H2/tri_sum*100:.2f}%")
        print(f"  Sum  = {tri_sum:.4f} ppm")

    # ── output path next to source file ──────────────────────────────────────
    base     = os.path.splitext(file_path)[0]
    out_path = base + "_duval_triangle.png"

    # ── draw and save ─────────────────────────────────────────────────────────
    draw(gases, meta, limits, out_path)


if __name__ == "__main__":
    main()
