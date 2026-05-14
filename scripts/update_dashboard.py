import os
import pandas as pd
import json
import re
from pathlib import Path
import http.server
import socketserver
import webbrowser
import sys
import openpyxl

# Configuración de Rutas
BASE_DIR = r"c:\Users\ANA KARINA\Desktop\CARTERAS PROPIEDAD HORIZONTAL"
TEMPLATE_FILE = "template.html"
OUTPUT_HTML = "../index.html"
PORT = 3001

def log(msg):
    print(msg, flush=True)

def process_file(file_path):
    results = []
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names_all = xls.sheet_names
        
        # 1. Identificar unidades a ignorar desde 'BASE DE DATOS'
        ignore_units = set()
        bd_sheet_name = next((n for n in sheet_names_all if "BASE DE DATOS" in n.upper()), None)
        if bd_sheet_name:
            try:
                # Leemos la pestaña de base de datos (header en fila 2, que es índice 1)
                df_bd = pd.read_excel(xls, sheet_name=bd_sheet_name, header=1)
                concepto_col = next((c for c in df_bd.columns if "CONCEPTO FINAL" in str(c).upper()), None)
                if concepto_col is not None:
                    for _, row in df_bd.iterrows():
                        unit_code = str(row.iloc[0]).strip()
                        if not unit_code or unit_code == "nan": continue
                        
                        concepto = str(row[concepto_col]).strip().upper()
                        if concepto in ["PAZ Y SALVO", "RETIRADO"]:
                            ignore_units.add(unit_code)
                            # Asegurar formatos comunes (001, 01, 1)
                            if unit_code.isdigit():
                                ignore_units.add(str(int(unit_code)))
                                ignore_units.add(unit_code.zfill(2))
                                ignore_units.add(unit_code.zfill(3))
            except Exception as e:
                log(f"Aviso: No se pudo procesar filtro en {file_path}: {e}")

        # 2. Procesar pestañas (excluyendo las ignoradas)
        sheet_names = [s for s in sheet_names_all if s.strip().isdigit() and s.strip() not in ignore_units]
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
        log(f"Generando index.html para despliegue...")
        # webbrowser.open("file:///" + os.path.abspath(OUTPUT_HTML).replace("\\", "/"))
    except Exception as e:
        log(f"Error: {e}")
