import pandas as pd
import numpy as np
import xlwings as xw
from PyQt5.QtCore import QThread, pyqtSignal
import os
import pickle
import config
import re
import traceback
from datetime import datetime

# --- Caching Helper ---
def load_or_cache(source, cache_key, read_func, progress_callback=None):
    if not os.path.exists(config.CACHE_DIR):
        os.makedirs(config.CACHE_DIR)

    if isinstance(source, str):
        file_name = os.path.basename(source)
        mtime = os.path.getmtime(source)
        cache_file = os.path.join(config.CACHE_DIR, f"cache_{cache_key}_{file_name}_{mtime}.pkl")

        if os.path.exists(cache_file):
            if progress_callback: progress_callback(50)
            try:
                with open(cache_file, 'rb') as f:
                    data = pickle.load(f)
                if progress_callback: progress_callback(100)
                print(f"[DEBUG] Cache loaded for {cache_key}")
                return data
            except Exception: pass

    if progress_callback: progress_callback(10)
    print(f"[DEBUG] Reading fresh data for {cache_key}...")
    data = read_func(source)
    if progress_callback: progress_callback(90)

    if isinstance(source, str):
        try:
            for f in os.listdir(config.CACHE_DIR):
                if f.startswith(f"cache_{cache_key}_{file_name}"):
                    os.remove(os.path.join(config.CACHE_DIR, f))
            with open(cache_file, 'wb') as f:
                pickle.dump(data, f)
        except Exception: pass

    if progress_callback: progress_callback(100)
    return data

# --- Helpers ---
def ensure_year(df):
    if "Year" not in df and "Month" in df:
        df["Year"] = pd.to_datetime(df["Month"], errors="coerce").dt.year
    return df

def compare_df(o, n, k, v):
    m = pd.merge(o, n, how="outer", on=k, suffixes=("_old", "_new"), indicator=True)
    return m[m["_merge"] == "left_only"], m[(m["_merge"] == "both") & (m[f"{v}_old"] != m[f"{v}_new"])]

def monthly_delta(o, n, r):
    if "Month" not in o or "Sales" not in o: return None
    for d in (o, n): d["Month"] = pd.to_datetime(d["Month"], errors="coerce").dt.strftime("%Y-%m")
    m = sorted(n["Month"].dropna().unique())
    if not m: return None
    l, p = m[-1], m[-2] if len(m) > 1 else None
    s = lambda d, x: d[d["Month"] == x]["Sales"].sum() if x else 0
    return {"Region": r, "Latest Month": l, "Prev Month": p or "",
            "Latest Δ": int(s(n, l) - s(o, l)), "Prev Δ": int(s(n, p) - s(o, p)) if p else ""}

def normalize_brand(name):
    if not isinstance(name, str): return str(name)
    u_name = name.strip().upper()
    if u_name in ['OPPO', 'REALME', 'ONEPLUS']:
        return 'Oppo'
    return name.strip()

# --- Readers (Existing) ---
def _read_weekly_impl(path):
    d = {}
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        for s in config.WEEKLY_SHEETS:
            try:
                df = wb.sheets[s].range(config.WEEKLY_START).expand("table").options(pd.DataFrame, header=1, index=False).value
                if s != "Basefile_Europe" and "Region" in df: df = df.drop(columns=["Region"])
                if "Brand" in df.columns: df["Brand"] = df["Brand"].apply(normalize_brand)
                d[s] = df
            except: d[s] = pd.DataFrame()
        wb.close()
    return d

def _read_monthly_impl(path):
    data_dict = {}
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        for sheet_name in config.MONTHLY_SHEETS:
            try:
                ws = wb.sheets[sheet_name]
                dates = ws.range((config.MONTHLY_DATE_ROW, 4)).expand('right').value
                brands_raw = ws.range(f"{config.MONTHLY_BRAND_COL}{config.MONTHLY_DATE_ROW+1}").expand('down').value
                valid_brands = []
                for b in brands_raw:
                    if str(b).strip().lower() == "total market": break
                    valid_brands.append(b)
                if not valid_brands or not dates: continue
                values = ws.range((config.MONTHLY_DATE_ROW+1, 4), (config.MONTHLY_DATE_ROW+len(valid_brands), 3+len(dates))).value
                df = pd.DataFrame(values, columns=dates)
                df.insert(0, "Brand", valid_brands)
                df_melt = df.melt(id_vars=["Brand"], var_name="Date", value_name="Sales")
                df_melt["Sales"] = pd.to_numeric(df_melt["Sales"], errors='coerce').fillna(0) * 1000000 
                df_melt["Date"] = pd.to_datetime(df_melt["Date"], errors='coerce')
                df_melt = df_melt.dropna(subset=["Date"])
                df_melt["Year"] = df_melt["Date"].dt.year; df_melt["Month"] = df_melt["Date"].dt.month; df_melt["Region"] = sheet_name
                if "Brand" in df_melt.columns: df_melt["Brand"] = df_melt["Brand"].apply(normalize_brand)
                data_dict[sheet_name] = df_melt[["Year", "Month", "Brand", "Region", "Sales"]]
            except: pass
        wb.close()
    return data_dict

def _read_flagship_impl(path):
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        try:
            ws = wb.sheets[config.FLAGSHIP_SHEET]
            data_range = ws.range((config.FLAGSHIP_HEADER_ROW, 3)).expand('table')
            df = data_range.options(pd.DataFrame, header=1, index=False).value
            df.columns = [str(c).strip() for c in df.columns]
            col_map = {}
            for c in df.columns:
                upper_c = c.upper()
                if "VENDOR" in upper_c or "BRAND" in upper_c: col_map[c] = "Brand"
                elif "MODEL" in upper_c: col_map[c] = "Model"
                elif "CATEGORY" in upper_c: col_map[c] = "Category"
            df.rename(columns=col_map, inplace=True)
            if 'Brand' in df.columns: df['Brand'] = df['Brand'].apply(normalize_brand)
            date_cols = [c for c in df.columns if c not in ['Brand', 'Model', 'Category']]
            df_melt = df.melt(id_vars=['Brand', 'Model', 'Category'], value_vars=date_cols, var_name="Date", value_name="Sales")
            df_melt['Date'] = pd.to_datetime(df_melt['Date'], errors='coerce')
            df_melt = df_melt.dropna(subset=['Date'])
            df_melt['Sales'] = pd.to_numeric(df_melt['Sales'], errors='coerce').fillna(0) * 1000000
            launch_dates = df_melt[df_melt['Sales'] > 0].groupby(['Brand', 'Model'])['Date'].min().reset_index()
            launch_dates.rename(columns={'Date': 'LaunchDate'}, inplace=True)
            df_final = pd.merge(df_melt, launch_dates, on=['Brand', 'Model'], how='left')
            df_final['MonthsSinceLaunch'] = (df_final['Date'].dt.year - df_final['LaunchDate'].dt.year) * 12 + (df_final['Date'].dt.month - df_final['LaunchDate'].dt.month)
            return df_final[df_final['MonthsSinceLaunch'] >= 0]
        finally: wb.close()

def _read_region_brand_impl(path):
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        try:
            ws = wb.sheets[config.REGION_BRAND_SHEET]
            data_range = ws.range(config.REGION_BRAND_START).expand('table')
            df = data_range.options(pd.DataFrame, header=1, index=False).value
            df.columns = [str(c).strip() for c in df.columns]
            for c in df.columns:
                if "Sell Through" in c: df.rename(columns={c: "Sales"}, inplace=True); break
            df["Sales"] = pd.to_numeric(df["Sales"], errors='coerce').fillna(0) * 1000000
            if "Month" in df.columns:
                df["Date_Obj"] = pd.to_datetime(df["Month"], format='%b %Y', errors='coerce')
                df["Year"] = df["Date_Obj"].dt.year; df["Month"] = df["Date_Obj"].dt.month
            return df
        finally: wb.close()

def _read_omdia_impl(path):
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        try:
            ws = wb.sheets[config.OMDIA_SHEET]
            df = ws.range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value
            df.columns = [str(c).strip() for c in df.columns]
            if 'Unit (Million)' not in df.columns and 'Unit (Thousand)' in df.columns:
                df['Unit (Million)'] = pd.to_numeric(df['Unit (Thousand)'], errors='coerce') / 1000.0
            
            df['Brand'] = df['Vendor'].apply(normalize_brand)
            df['Category'] = df.get('Form factor', 'Smartphone').apply(lambda x: 'Foldable' if 'foldable' in str(x).lower() else 'Smartphone')
            col_unit = 'Unit (Million)' if 'Unit (Million)' in df.columns else 'Unit (Thousand)'
            multiplier = 1000000 if 'Unit (Million)' in df.columns else 1000
            df['Sales'] = pd.to_numeric(df[col_unit], errors='coerce').fillna(0) * multiplier
            df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
            df['Quarter'] = df['Quarter'].astype(str).str.extract(r'(\d)').astype(float).fillna(0).astype(int)
            
            launch_data = df[df['Sales'] > 0].groupby(['Brand', 'Model'])[['Year', 'Quarter']].min().reset_index()
            launch_data.rename(columns={'Year': 'LaunchYear', 'Quarter': 'LaunchQuarter'}, inplace=True)
            df_final = pd.merge(df, launch_data, on=['Brand', 'Model'], how='left')
            df_final['QuartersSinceLaunch'] = (df_final['Year'] - df_final['LaunchYear']) * 4 + (df_final['Quarter'] - df_final['LaunchQuarter'])
            return df_final[df_final['QuartersSinceLaunch'] >= 0]
        finally: wb.close()

def _read_ti_impl(source):
    wb = None; app = None; close_after = False
    try:
        if isinstance(source, str):
            app = xw.App(visible=False); wb = app.books.open(source, read_only=True); close_after = True
        else: wb = source
        ws = wb.sheets[config.TI_SHEET]
        df = ws.range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value
        df.columns = [str(c).strip() for c in df.columns]
        df['Sales'] = pd.to_numeric(df['Value (M)'], errors='coerce').fillna(0) * 1000000
        month_map = {'January':1,'February':2,'March':3,'April':4,'May':5,'June':6,'July':7,'August':8,'September':9,'October':10,'November':11,'December':12}
        df['Month'] = df['Month'].map(month_map).fillna(0).astype(int)
        df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
        df['Brand'] = df['Vendor'].apply(normalize_brand)
        return df
    finally:
        if close_after and wb: wb.close()
        if close_after and app: app.quit()

def _read_generic_impl(source):
    wb = None; app = None; close_after = False
    try:
        if isinstance(source, str):
            app = xw.App(visible=False); wb = app.books.open(source, read_only=True); close_after = True
        else: wb = source
        ws = wb.sheets.active
        df = ws.range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value
        return df
    except Exception as e:
        raise Exception(f"Excel Read Error: {e}")
    finally:
        if close_after and wb: wb.close()
        if close_after and app: app.quit()

def _read_ti_shipment_impl(source):
    wb = None; app = None; close_after = False
    try:
        if isinstance(source, str): app = xw.App(visible=False); wb = app.books.open(source, read_only=True); close_after = True
        else: wb = source
        ws = wb.sheets[config.TI_SHIPMENT_SHEET]
        df = ws.range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value
        df.columns = [str(c).strip() for c in df.columns]
        if 'Metric Name' in df.columns:
            df = df[df['Metric Name'].astype(str).str.lower() == 'shipments']
        df = df[df['Brand'] == 'Apple']
        df['Year'] = pd.to_numeric(df['Year'], errors='coerce').fillna(0).astype(int)
        df = df[df['Year'] >= 2020]
        df['Model'] = df['Brand and Model Name'].astype(str).str.replace('Apple ', '').str.strip()
        df['Date'] = df['Year'].astype(str) + " Q" + df['Quarter'].astype(str)
        df['Value'] = pd.to_numeric(df['Metric Value'], errors='coerce').fillna(0)
        df['Firm'] = 'TI'
        return df[['Model', 'Date', 'Value', 'Firm']]
    except Exception as e: raise Exception(f"TI Shipment Read Error: {e}")
    finally:
        if close_after and wb: wb.close()
        if close_after and app: app.quit()

def _read_gfk_impl(source):
    wb = None; app = None; close_after = False
    try:
        if isinstance(source, str): app = xw.App(visible=False); wb = app.books.open(source, read_only=True); close_after = True
        else: wb = source
        ws = wb.sheets[config.GFK_SHEET]
        years_row = ws.range((2, 1)).expand('right').value
        quarters_row = ws.range((3, 1)).expand('right').value
        col_map = {}
        current_year = None
        for i, y in enumerate(years_row):
            if y is not None: current_year = str(int(y)) if isinstance(y, (int, float)) else str(y).strip()
            if "Total" in str(current_year) or "Total" in str(y): continue
            q_raw = quarters_row[i]
            if q_raw and "Q" in str(q_raw):
                q_num = str(q_raw).replace("Q", "").strip()
                date_str = f"{current_year} Q{q_num}"
                try:
                    if int(current_year) >= 2020: col_map[i] = date_str
                except: pass
        data_start_row = 4
        last_row = ws.range(f'B{ws.cells.last_cell.row}').end('up').row
        models = ws.range((data_start_row, 2), (last_row, 2)).value
        valid_indices = sorted(col_map.keys())
        if not valid_indices: return pd.DataFrame()
        min_col = min(valid_indices); max_col = max(valid_indices)
        val_block = ws.range((data_start_row, min_col+1), (last_row, max_col+1)).value
        records = []
        for r_idx, model_name in enumerate(models):
            if not model_name: continue
            m_str = str(model_name).strip()
            if "iPhone" not in m_str: continue 
            row_vals = val_block[r_idx]
            for c_idx in valid_indices:
                offset = c_idx - min_col
                if offset < len(row_vals):
                    val = row_vals[offset]
                    date = col_map[c_idx]
                    try: v_float = float(val)
                    except: v_float = 0.0
                    records.append({"Model": m_str, "Date": date, "Value": v_float, "Firm": "GfK"})
        return pd.DataFrame(records)
    except Exception as e: raise Exception(f"GfK Read Error: {e}")
    finally:
        if close_after and wb: wb.close()
        if close_after and app: app.quit()

def _read_sellin_new_impl(path):
    print(f"\n[DEBUG] === Starting Sell-in Read from: {os.path.basename(path)} ===")
    data_list = []
    app = xw.App(visible=False)
    try:
        wb = app.books.open(path, read_only=True)
        
        for sheet_name in config.SELLIN_SHEETS:
            print(f"[DEBUG] Target Sheet: '{sheet_name}'")
            try:
                ws = wb.sheets[sheet_name]
                print(f"[DEBUG] -> Sheet '{sheet_name}' Found.")
            except:
                print(f"[DEBUG] -> Sheet '{sheet_name}' NOT found. Skipping.")
                continue

            region_name = config.SELLIN_SHEET_MAP.get(sheet_name, sheet_name)

            # 1. Read Vendors
            print(f"[DEBUG] -> Reading Vendors from Column {config.SELLIN_VENDOR_COL}, starting Row {config.SELLIN_START_ROW}...")
            vendor_range_vals = ws.range(f"{config.SELLIN_VENDOR_COL}{config.SELLIN_START_ROW}:{config.SELLIN_VENDOR_COL}500").value
            
            valid_vendors = []
            row_count = 0
            
            for v in vendor_range_vals:
                if v is None:
                    valid_vendors.append(None) 
                    row_count += 1
                    continue
                v_str = str(v).strip()
                if v_str.lower() == "total market":
                    print(f"[DEBUG] -> Found 'Total Market' at relative row {row_count}. Stopping vendor read.")
                    break
                valid_vendors.append(v_str)
                row_count += 1
            
            clean_vendors = [v for v in valid_vendors if v is not None]
            print(f"[DEBUG] -> Recognized {len(clean_vendors)} Vendors: {clean_vendors[:5]} ...")
            
            if not clean_vendors: 
                print("[DEBUG] -> No valid vendors found. Skipping.")
                continue

            # 2. Read Date Headers
            print(f"[DEBUG] -> Reading Dates from Row {config.SELLIN_DATE_ROW}...")
            date_vals = ws.range((config.SELLIN_DATE_ROW, 3), (config.SELLIN_DATE_ROW, 200)).value
            
            valid_dates = []
            for d in date_vals:
                if d is None: break 
                valid_dates.append(d)
            
            col_count = len(valid_dates)
            print(f"[DEBUG] -> Found {col_count} Date Columns.")

            if col_count == 0:
                print("[DEBUG] -> No date columns found. Skipping.")
                continue

            # 3. Read Data Values
            start_row = config.SELLIN_START_ROW
            end_row = start_row + row_count - 1
            start_col = 3
            end_col = start_col + col_count - 1
            
            print(f"[DEBUG] -> Reading Data Block: Rows {start_row}~{end_row}, Cols {start_col}~{end_col}")
            values = ws.range((start_row, start_col), (end_row, end_col)).value
            if row_count == 1: values = [values]
            
            # 4. Construct Data
            temp_data = []
            for i, vendor in enumerate(valid_vendors):
                if vendor is None: continue 
                if i < len(values):
                    row_data = values[i]
                    if len(row_data) < col_count:
                        row_data += [None] * (col_count - len(row_data))
                    elif len(row_data) > col_count:
                        row_data = row_data[:col_count]
                        
                    record = {"Brand": vendor}
                    for j, date_val in enumerate(valid_dates):
                        record[date_val] = row_data[j]
                    temp_data.append(record)
            
            if not temp_data: continue

            df = pd.DataFrame(temp_data)
            
            # Melt
            df_melt = df.melt(id_vars=["Brand"], var_name="Date", value_name="Sales")
            
            # Conversions
            df_melt["Date_Obj"] = pd.to_datetime(df_melt["Date"], errors='coerce')
            df_melt = df_melt.dropna(subset=["Date_Obj"])
            
            df_melt["Year"] = df_melt["Date_Obj"].dt.year
            df_melt["Month"] = df_melt["Date_Obj"].dt.month
            
            # [FIXED] Add Region Column
            df_melt["Region"] = region_name
            
            # [MODIFIED] Multiply by 1M to store as Units, so display logic (which divides by 1M) works correct
            df_melt["Sales"] = pd.to_numeric(df_melt["Sales"], errors='coerce').fillna(0) * 1000000
            
            print(f"[DEBUG] -> Sheet '{sheet_name}' Processed. {len(df_melt)} rows created.")
            data_list.append(df_melt)
            
    except Exception as e:
        import traceback
        print(f"[ERROR] Exception in _read_sellin_impl:\n{traceback.format_exc()}")
        raise Exception(f"Sell-in Read Error: {e}")
    finally:
        try: wb.close()
        except: pass
        app.quit()
        
    if data_list:
        final_df = pd.concat(data_list, ignore_index=True)
        print(f"[DEBUG] === Sell-in Read Complete. Total Rows: {len(final_df)} ===\n")
        return final_df
    else:
        print("[DEBUG] === Sell-in Read Failed: No data collected ===\n")
        return pd.DataFrame()


# --- Threads ---
class CompareThread(QThread):
    result = pyqtSignal(pd.DataFrame, list, dict); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, o, n): super().__init__(); self.o=o; self.n=n
    def run(self):
        try:
            self.progress.emit(10); nd = load_or_cache(self.n, "weekly", _read_weekly_impl, lambda x: self.progress.emit(10+int(x*0.4)))
            od = None
            if self.o: self.progress.emit(50); od = load_or_cache(self.o, "weekly", _read_weekly_impl, lambda x: self.progress.emit(50+int(x*0.4)))
            self.progress.emit(90); rows=[]; sumy=[]
            if od:
                for s in config.WEEKLY_SHEETS:
                    if s in od and s in nd:
                        r=config.WEEKLY_MAP[s]; o_df=ensure_year(od[s]); n_df=ensure_year(nd[s])
                        k=["Brand","Model","Month","Week"]; o_df["Sales"]=pd.to_numeric(o_df["Sales"], errors="coerce"); n_df["Sales"]=pd.to_numeric(n_df["Sales"], errors="coerce")
                        if s=="Basefile_Europe": k=["Region"]+k
                        elif "Region" in o_df.columns: o_df=o_df.drop(columns=["Region"]); n_df=n_df.drop(columns=["Region"])
                        rem, chg = compare_df(o_df, n_df, k, "Sales")
                        for _,x in rem.iterrows(): rows.append({"Sheet":s,"Brand":x.get("Brand",""),"Model":x.get("Model",""),"Region":r,"Type":"Deleted","Sales_old":x.get("Sales_old",""),"Sales_new":""})
                        for _,x in chg.iterrows(): rows.append({"Sheet":s,"Brand":x.get("Brand",""),"Model":x.get("Model",""),"Region":r,"Type":"Changed","Sales_old":x.get("Sales_old",""),"Sales_new":x.get("Sales_new","")})
                        d=monthly_delta(o_df, n_df, r); 
                        if d: sumy.append(d)
            self.progress.emit(100); self.result.emit(pd.DataFrame(rows), sumy, nd)
        except Exception as e: self.error.emit(str(e))

class MonthlyCompareThread(QThread):
    result = pyqtSignal(pd.DataFrame, list, dict); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, o, n): super().__init__(); self.o=o; self.n=n
    def run(self):
        try:
            self.progress.emit(10); nd = load_or_cache(self.n, "monthly", _read_monthly_impl, lambda x: self.progress.emit(10+int(x*0.4)))
            od = None
            if self.o: self.progress.emit(50); od = load_or_cache(self.o, "monthly", _read_monthly_impl, lambda x: self.progress.emit(50+int(x*0.4)))
            self.progress.emit(90); rows=[]; sumy=[]
            if od:
                for s in config.MONTHLY_SHEETS:
                    if s in od and s in nd:
                        r=s; o_df=ensure_year(od[s]); n_df=ensure_year(nd[s]); k=["Brand","Month","Year","Region"]
                        o_df["Sales"]=pd.to_numeric(o_df["Sales"], errors="coerce"); n_df["Sales"]=pd.to_numeric(n_df["Sales"], errors="coerce")
                        rem, chg = compare_df(o_df, n_df, k, "Sales")
                        for _,x in rem.iterrows(): rows.append({"Sheet":s,"Brand":x.get("Brand",""),"Model":"","Region":r,"Type":"Deleted","Sales_old":x.get("Sales_old",""),"Sales_new":""})
                        for _,x in chg.iterrows(): rows.append({"Sheet":s,"Brand":x.get("Brand",""),"Model":"","Region":r,"Type":"Changed","Sales_old":x.get("Sales_old",""),"Sales_new":x.get("Sales_new","")})
                        d=monthly_delta(o_df, n_df, r); 
                        if d: sumy.append(d)
            self.progress.emit(100); self.result.emit(pd.DataFrame(rows), sumy, nd)
        except Exception as e: self.error.emit(str(e))

class FlagshipThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, p): super().__init__(); self.path=p
    def run(self):
        try: self.progress.emit(10); data = load_or_cache(self.path, "flagship", _read_flagship_impl, lambda x: self.progress.emit(x)); self.progress.emit(100); self.result.emit(data)
        except Exception as e: self.error.emit(str(e))

class RegionBrandThread(QThread):
    result = pyqtSignal(dict); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, p): super().__init__(); self.path=p
    def run(self):
        try: self.progress.emit(10); df = load_or_cache(self.path, "region", _read_region_brand_impl, lambda x: self.progress.emit(x)); self.progress.emit(100); self.result.emit({'AllData': df})
        except Exception as e: self.error.emit(str(e))

class OmdiaThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, p): super().__init__(); self.path=p
    def run(self):
        try: self.progress.emit(10); df = load_or_cache(self.path, "omdia", _read_omdia_impl, lambda x: self.progress.emit(x)); self.progress.emit(100); self.result.emit(df)
        except Exception as e: self.error.emit(str(e))

class TIThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, source): super().__init__(); self.source = source 
    def run(self):
        try: self.progress.emit(10); df = load_or_cache(self.source, "ti", _read_ti_impl, lambda x: self.progress.emit(x)); self.progress.emit(100); self.result.emit(df)
        except Exception as e: self.error.emit(str(e))

class GenericThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, source): super().__init__(); self.source = source 
    def run(self):
        try: self.progress.emit(10); df = load_or_cache(self.source, "generic", _read_generic_impl, lambda x: self.progress.emit(x)); self.progress.emit(100); self.result.emit(df)
        except Exception as e: self.error.emit(str(e))

class ByModelLoader(QThread):
    result = pyqtSignal(pd.DataFrame, str); error = pyqtSignal(str) 
    def __init__(self, path, firm):
        super().__init__()
        self.path = path
        self.firm = firm
        
    def run(self):
        try:
            df = pd.DataFrame()
            if self.firm == 'Omdia':
                raw_df = load_or_cache(self.path, "omdia", _read_omdia_impl)
                raw_df = raw_df[(raw_df['Brand'] == 'Apple') & (raw_df['Year'] >= 2020)]
                raw_df['Value'] = raw_df['Sales'] / 1000000.0
                raw_df['Date'] = raw_df['Year'].astype(str) + " Q" + raw_df['Quarter'].astype(str)
                raw_df['Firm'] = 'Omdia'
                df = raw_df[['Model', 'Date', 'Value', 'Firm']]
                
            elif self.firm == 'TI':
                df = _read_ti_shipment_impl(self.path)
                
            elif self.firm == 'GfK':
                df = _read_gfk_impl(self.path)
            
            self.result.emit(df, self.firm)
            
        except Exception as e:
            self.error.emit(f"{self.firm} Load Error: {str(e)}")

# [NEW] Sell In Thread
class SellInThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, path): super().__init__(); self.path = path
    def run(self):
        try: 
            self.progress.emit(10)
            # [수정됨] 캐시 키를 "sellin_final_v1"으로 변경하여 강제 리로드 유도
            df = load_or_cache(self.path, "sellin_final_v1", _read_sellin_new_impl, lambda x: self.progress.emit(x))
            self.progress.emit(100)
            self.result.emit(df)
        except Exception as e: self.error.emit(str(e))

class WeeklySimpleThread(QThread):
    result = pyqtSignal(dict); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, path): super().__init__(); self.path = path
    def run(self):
        try:
            self.progress.emit(10)
            # 기존 _read_weekly_impl 함수 재사용 (Weekly 탭과 동일한 로직으로 읽음)
            # 캐시 키는 'weekly_simple'로 지정하여 충돌 방지
            data = load_or_cache(self.path, "weekly_simple", _read_weekly_impl, lambda x: self.progress.emit(x))
            self.progress.emit(100)
            self.result.emit(data)
        except Exception as e: self.error.emit(str(e))
