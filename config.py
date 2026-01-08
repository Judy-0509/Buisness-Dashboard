import os
import sys
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib import rcParams

# --- 시스템 설정 ---
CACHE_DIR = "cache"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- 디자인 색상 ---
THEMES = {
    "Counterpoint": {"dark": "#39A2DB", "light": "#E8F0F2", "border": "#D0D0D0"},
    "Omdia":         {"dark": "#39A2DB", "light": "#E8F0F2", "border": "#D0D0D0"}, 
    "TechInsights": {"dark": "#39A2DB", "light": "#E8F0F2", "border": "#D0D0D0"},
    "Pivot":         {"dark": "#27AE60", "light": "#E9F7EF", "border": "#D5F5E3"},
    "ByModel":       {"dark": "#8E44AD", "light": "#F4ECF7", "border": "#D2B4DE"}
}

CURRENT_THEME = THEMES["Counterpoint"]
BG_MAIN = "#E8F0F2"
HEADER_BG = "#39A2DB"
CARD_BG = "#FFFFFF"
BORDER = "#E0E0E0"

BLACK = "#000000"
WHITE = "#FFFFFF"
POS = "#2563EB"
NEG = "#EF4444"

COLOR_23 = "#DDEFF9"
COLOR_24 = "#BADFF3"
COLOR_25 = "#40A6DD"

# --- 엑셀 설정 ---
WEEKLY_SHEETS = ["Basefile_US", "Basefile_China", "Basefile_Japan", "Basefile_Europe", "Basefile_India"]
WEEKLY_MAP = {"Basefile_US": "US", "Basefile_China": "China", "Basefile_Japan": "Japan", 
              "Basefile_Europe": "Europe", "Basefile_India": "India"}
WEEKLY_START = "B9"

MONTHLY_SHEETS = ["China", "USA", "India", "Europe", "Others"]
MONTHLY_DATE_ROW = 44
MONTHLY_BRAND_COL = "C"

FLAGSHIP_SHEET = "Bestsellers"
FLAGSHIP_HEADER_ROW = 13

REGION_BRAND_SHEET = "Basefile"
REGION_BRAND_START = "B9"

OMDIA_SHEET = "Raw"
TI_SHEET = "11. FlatFile"
TI_SHIPMENT_SHEET = "6. SP Shipments Flat File"
GFK_SHEET = "Global Sell-in Summary"

# Sell in Sell Thru Settings
SELLIN_SHEETS = ["Global SP"]
SELLIN_SHEET_MAP = {
    "Global SP": "Total", 
    "China SP": "China", 
    "India SP": "India",
    "USA SP": "US",
    "Europe SP": "W.Europe"
}
SELLIN_DATE_ROW = 30
SELLIN_START_ROW = 31
SELLIN_VENDOR_COL = "B"
SELLIN_DATA_START_COL = 3

# --- 폰트 설정 ---
rcParams['font.family'] = 'Malgun Gothic'
rcParams['axes.unicode_minus'] = False

def generate_gradient_colors(n):
    if n < 1: return []
    cmap = plt.get_cmap("tab20")
    return [mcolors.to_hex(cmap(i % 20)) for i in range(n)]
