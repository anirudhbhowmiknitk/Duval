import os
import pandas as pd

# ============================================================
# FILE PATH
# ============================================================

ch4 = 1.15
c2h4 = 1.85
c2h2 = 1.65
total = CH4 + C2H4 + C2H2

# ============================================================
# PERCENTAGE CALCULATION
# ============================================================

pCH4 = (CH4 / total) * 100 if total > 0 else 0
pC2H4 = (C2H4 / total) * 100 if total > 0 else 0
pC2H2 = (C2H2 / total) * 100 if total > 0 else 0

# ============================================================
# DUVAL TRIANGLE CLASSIFICATION
# ============================================================

def classify_duval(pCH4, pC2H4, pC2H2):

    if pC2H2 >= 29:
        if pCH4 <= 23:
            return "DT"
        return "D2"

    if pC2H2 >= 13:
        if pC2H4 <= 60:
            return "D2"
        return "DT"

    if pC2H2 >= 2:

        if pC2H4 < 24:
            return "T1"

        if pCH4 >= 87:
            return "D1"

        return "D2"

    if pC2H4 >= 60:
        return "T3"

    if pC2H4 >= 24:
        return "T2"

    if pCH4 >= 98:
        return "PD"

    return "T1"

# ============================================================
# FAULT RESULT
# ============================================================

fault_zone = classify_duval(pCH4, pC2H4, pC2H2)

fault_meanings = {
    "PD": "Partial Discharge",
    "T1": "Thermal Fault < 300°C",
    "T2": "Thermal Fault 300–700°C",
    "T3": "Thermal Fault > 700°C",
    "D1": "Low Energy Electrical Discharge",
    "D2": "High Energy Electrical Discharge",
    "DT": "Mixed Thermal + Discharge Fault"
}

# ============================================================
# TERMINAL OUTPUT
# ============================================================

print("\n========== DUVAL ANALYSIS ==========")

print("\nGas Values")

print(f"CH4   : {CH4:.2f} ppm")
print(f"C2H4 : {C2H4:.2f} ppm")
print(f"C2H2 : {C2H2:.2f} ppm")

print("\nPercentage Contribution")

print(f"CH4   : {pCH4:.2f}%")
print(f"C2H4 : {pC2H4:.2f}%")
print(f"C2H2 : {pC2H2:.2f}%")

print("\nFault Classification")

print(f"Fault Zone : {fault_zone}")
print(f"Fault Type : {fault_meanings[fault_zone]}")

print("\n====================================")