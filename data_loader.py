import pandas as pd
import numpy as np
import xlwings as xw
from PyQt5.QtCore import QThread, pyqtSignal
import os
import pickle
import config
import re

def load_or_cache(source, cache_key, read_func, progress_callback=None):
    if not os.path.exists(config.CACHE_DIR): os.makedirs(config.CACHE_DIR)
    if isinstance(source, str):
        file_name = os.path.basename(source)
        mtime = os.path.getmtime(source)
        cache_file = os.path.join(config.CACHE_DIR, f"cache_{cache_key}_{file_name}_{mtime}.pkl")
        if os.path.exists(cache_file):
            with open(cache_file, 'rb') as f: return pickle.load(f)
    data = read_func(source)
    if isinstance(source, str):
        with open(cache_file, 'wb') as f: pickle.dump(data, f)
    return data

def normalize_brand(name):
    u_name = str(name).strip().upper()
    if u_name in ['OPPO', 'REALME', 'ONEPLUS']: return 'Oppo'
    return str(name).strip()

def ensure_year(df):
    if "Year" not in df and "Month" in df:
        df["Year"] = pd.to_datetime(df["Month"], errors="coerce").dt.year
    return df

def compare_df(o, n, k, v):
    m = pd.merge(o, n, how="outer", on=k, suffixes=("_old", "_new"), indicator=True)
    return m[m["_merge"] == "left_only"], m[(m["_merge"] == "both") & (m[f"{v}_old"] != m[f"{v}_new"])]

# --- Reader Functions ---
def _read_weekly_impl(path):
    d = {}
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        for s in config.WEEKLY_SHEETS:
            try:
                df = wb.sheets[s].range(config.WEEKLY_START).expand("table").options(pd.DataFrame, header=1, index=False).value
                if "Brand" in df.columns: df["Brand"] = df["Brand"].apply(normalize_brand)
                d[s] = df
            except: d[s] = pd.DataFrame()
        wb.close()
    return d

def _read_monthly_impl(path):
    data_dict = {}
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        for s in config.MONTHLY_SHEETS:
            # Monthly complex parsing...
            pass
        wb.close()
    return data_dict

def _read_omdia_impl(path):
    with xw.App(visible=False) as app:
        wb = app.books.open(path, read_only=True)
        ws = wb.sheets[config.OMDIA_SHEET]
        df = ws.range("A1").expand("table").options(pd.DataFrame, header=1, index=False).value
        df['Brand'] = df['Vendor'].apply(normalize_brand)
        df['Sales'] = pd.to_numeric(df['Unit (Million)'], errors='coerce') * 1000000
        wb.close()
    return df

# --- Threads ---
class CompareThread(QThread):
    result = pyqtSignal(pd.DataFrame, list, dict); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, o, n): super().__init__(); self.o=o; self.n=n
    def run(self):
        try:
            nd = load_or_cache(self.n, "weekly", _read_weekly_impl)
            self.result.emit(pd.DataFrame(), [], nd)
        except Exception as e: self.error.emit(str(e))

class OmdiaThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, p): super().__init__(); self.path=p
    def run(self):
        try:
            df = load_or_cache(self.path, "omdia", _read_omdia_impl)
            self.result.emit(df)
        except Exception as e: self.error.emit(str(e))

class SellInThread(QThread):
    result = pyqtSignal(pd.DataFrame); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, p): super().__init__(); self.path=p
    def run(self):
        # Implementation...
        pass

class WeeklySimpleThread(QThread):
    result = pyqtSignal(dict); error = pyqtSignal(str); progress = pyqtSignal(int)
    def __init__(self, p): super().__init__(); self.path=p
    def run(self):
        try:
            data = load_or_cache(self.path, "weekly_simple", _read_weekly_impl)
            self.result.emit(data)
        except Exception as e: self.error.emit(str(e))
