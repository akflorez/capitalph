import os
import pandas as pd
import json
import re
from pathlib import Path
import http.server
import socketserver
import webbrowser
import sys

# Configuración de Rutas
BASE_DIR = r"c:\Users\ANA KARINA\Desktop\CARTERAS PROPIEDAD HORIZONTAL"
TEMPLATE_FILE = r"C:\Users\ANA KARINA\.gemini\antigravity\scratch\dashboard_template.html"
OUTPUT_HTML = r"c:\Users\ANA KARINA\Desktop\DASHBOARD_CARTERAS.html"
PORT = 3001

def log(msg):
    print(msg)

def process_file(file_path):
    results = []
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = [s for s in xls.sheet_names if s.strip().isdigit()]
        for sheet in sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            debt_sum = 0
            concepto_row = -1
            deuda_col = -1
            for idx, row in df.iterrows():
                row_vals = [str(val).strip().upper() for val in row]
                if "CONCEPTO" in row_vals:
                    concepto_row = idx
                    for col_idx, col_val in enumerate(row_vals):
                        if col_val in ["VALOR DEUDA", "VALOR"]:
                            deuda_col = col_idx
                            break
                    break
            if concepto_row != -1 and deuda_col != -1:
                for i in range(concepto_row + 1, len(df)):
                    val = df.iloc[i, deuda_col]
                    row_label = str(df.iloc[i, 0]).strip().upper()
                    if row_label == "TOTAL": break
                    if pd.notnull(val) and isinstance(val, (int, float)):
                        debt_sum += val
            abono_sum = 0
            payment_records = []
            factura_row = -1
            fecha_col = -1
            abono_cap_col = -1
            for idx, row in df.iterrows():
                row_vals = [str(val).strip().upper() for val in row]
                if "FACTURA" in row_vals:
                    factura_row = idx
                    for col_idx, col_val in enumerate(row_vals):
                        if "FECHA" in col_val: fecha_col = col_idx
                        if "ABONO CAPITAL" in col_val: abono_cap_col = col_idx
                    break
            if factura_row != -1 and abono_cap_col != -1:
                for i in range(factura_row + 1, len(df)):
                    row_data = df.iloc[i]
                    row_label = str(row_data[0]).strip().upper()
                    if "TOTAL" in row_label: break
                    val = row_data[abono_cap_col]
                    date_val = row_data[fecha_col] if fecha_col != -1 else None
                    if pd.notnull(val) and isinstance(val, (int, float)) and val > 0:
                        abono_sum += val
                        if pd.notnull(date_val):
                            try:
                                dt = pd.to_datetime(date_val)
                                if dt <= pd.Timestamp.now():
                                    payment_records.append({"d": dt.strftime('%Y-%m-%d'), "v": float(val)})
                            except: pass
            matricula = None
            for idx in range(min(12, len(df))):
                row_vals = [str(v).strip().upper() for v in df.iloc[idx]]
                if "MATRICULA INMOBILIARIA" in row_vals:
                    lab_idx = row_vals.index("MATRICULA INMOBILIARIA")
                    for c_idx in range(lab_idx + 1, min(lab_idx + 5, len(df.columns))):
                        candidate = str(df.iloc[idx, c_idx]).strip()
                        if candidate and candidate != "nan" and "-" in candidate:
                            matricula = candidate
                            break
                    break
            results.append({
                "unit": sheet,
                "matricula": matricula,
                "debt": float(debt_sum),
                "abonos": float(abono_sum),
                "balance": float(debt_sum - abono_sum),
                "history": payment_records
            })
    except Exception as e:
        log(f"Error en archivo {file_path}: {e}")
    return results

def update_data():
    log("Actualizando datos...")
    all_data = []
    folders = [f for f in os.listdir(BASE_DIR) if os.path.isdir(os.path.join(BASE_DIR, f))]
    folders.sort()
    for folder in folders:
        folder_path = os.path.join(BASE_DIR, folder)
        files = [f for f in os.listdir(folder_path) if f.lower().endswith(".xlsx") and not f.startswith("~$")]
        if not files: continue
        bldg_units = []
        for excel_file in files:
            file_path = os.path.join(folder_path, excel_file)
            bldg_units.extend(process_file(file_path))
        if not bldg_units: continue
        unique_mats = {u["matricula"] for u in bldg_units if u["matricula"]}
        bldg_payments = []
        for u in bldg_units: bldg_payments.extend(u["history"])
        all_data.append({
            "conjunto": folder,
            "units": bldg_units,
            "total_matriculas_unicas": len(unique_mats),
            "total_debt": sum(u["debt"] for u in bldg_units),
            "total_abonos": sum(u["abonos"] for u in bldg_units),
            "total_balance": sum(u["balance"] for u in bldg_units),
            "payment_history": bldg_payments
        })
    with open(TEMPLATE_FILE, "r", encoding="utf-8") as f:
        template = f.read()
    json_data = json.dumps(all_data, ensure_ascii=False, indent=2)
    final_html = template.replace("__DATA_JSON__", json_data)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(final_html)
    log("¡Dashboard actualizado!")

if __name__ == "__main__":
    try:
        update_data()
        log(f"Abriendo dashboard directamente...")
        webbrowser.open("file:///" + OUTPUT_HTML.replace("\\", "/"))
    except Exception as e:
        log(f"Error: {e}")
