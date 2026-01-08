import sys
import os
import re
import pandas as pd
import xlwings as xw
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame, 
                             QCheckBox, QButtonGroup, QFileDialog, QTableWidget, QMessageBox, QTableWidgetItem, 
                             QMenu, QAction, QListWidget, QListWidgetItem, QSplitter, QSpinBox, QProgressBar, 
                             QDialog, QComboBox, QTabWidget, QApplication, QSizePolicy)
from PyQt5.QtCore import Qt, QPropertyAnimation, QEasingCurve, QRectF, pyqtSignal, QTimer, pyqtProperty, QSettings
from PyQt5.QtGui import QColor, QPainter, QFont

import config

# --- User Modules Import ---
from data_loader import (CompareThread, MonthlyCompareThread, FlagshipThread, RegionBrandThread, 
                         OmdiaThread, TIThread, GenericThread, ByModelLoader, SellInThread, WeeklySimpleThread)
from charts import (HeatmapWidget, LineChartWidget, TrendWidget, LaunchTrendWidget, 
                    LaunchTableWidget, PivotWidget, AdvancedPivotWidget, ComparisonTableWidget, DetailChartWidget)

# --- Helper Functions ---
def extract_version(filepath):
    if not filepath: return (-1, -1, -1)
    filename = os.path.basename(filepath).lower()
    year = 0
    y4_match = re.search(r'(20[2-3]\d)', filename)
    if y4_match: year = int(y4_match.group(1))
    else:
        y2_match = re.search(r"(?:['qQ_]|^)(2[0-9])(?:[^\d]|$)", filename)
        if y2_match: year = 2000 + int(y2_match.group(1))
    sub_unit = 0; day = 0
    week_match = re.search(r'(\d{1,2})\s*weeks?', filename)
    if not week_match: week_match = re.search(r'w(\d{1,2})', filename)
    if week_match: sub_unit = int(week_match.group(1))
    else:
        months_map = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
        for m_str, m_int in months_map.items():
            if m_str in filename: sub_unit = m_int; break
        if sub_unit == 0:
            q_match = re.search(r'([1-4])q|q([1-4])', filename)
            if q_match: sub_unit = int(q_match.group(1) or q_match.group(2)) * 3 
    if sub_unit == 0 and year > 0:
        y2_str = str(year)[2:]
        mmdd_match = re.search(r'(\d{3,4})[\s_\'\-]*' + y2_str, filename)
        if mmdd_match:
            val = int(mmdd_match.group(1))
            if 100 <= val <= 1231: sub_unit = val // 100; day = val % 100
    if year == 0:
        numbers = re.findall(r'\d+', filename)
        if numbers: return tuple(map(int, numbers))
        return (0, 0, 0)
    return (year, sub_unit, day)

# --- Basic Widgets ---
class FileDrop(QFrame):
    def __init__(self, t, cb):
        super().__init__()
        self.cb = cb; self.setAcceptDrops(True); self.setFixedHeight(80)
        self.setStyleSheet(f"QFrame{{background:{config.CARD_BG};border:2px dashed {config.HEADER_BG};border-radius:10px;}}")
        l = QVBoxLayout(self)
        self.lb = QLabel(f"{t}\nDrag & Drop"); self.lb.setAlignment(Qt.AlignCenter)
        self.lb.setFont(QFont("나눔스퀘어 네오 Light", 8)); self.lb.setStyleSheet(f"color:{config.BLACK};") 
        l.addWidget(self.lb)
    def mousePressEvent(self, e):
        p, _ = QFileDialog.getOpenFileName(self, "Select Excel", "", "Excel Files (*.xlsx *.xls *.xlsb *.xlsm)"); 
        if p: self.set(p)
    def dragEnterEvent(self, e): 
        if e.mimeData().hasUrls(): e.acceptProposedAction()
    def dropEvent(self, e):
        urls = e.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(('.xlsx', '.xls', '.xlsb', '.xlsm'))]
        if paths:
            if len(paths) == 1: self.set(paths[0])
            else: self.set(paths) 
    def set(self, p): self.cb(p)
    def update_label(self, p):
        if isinstance(p, list): txt = f"{len(p)} Files"
        else: txt = os.path.basename(p)
        self.lb.setText(txt); self.lb.setFont(QFont("나눔스퀘어 네오 Light", 9)); self.lb.setStyleSheet(f"color:{config.BLACK};")

class MultiStateToggle(QFrame):
    mode_changed = pyqtSignal(str)
    def __init__(self, parent=None):
        super().__init__(parent); self.setFixedSize(220, 30); self.current_theme_color = config.HEADER_BG
        self.setStyleSheet(f"QFrame {{ background-color: #E0E0E0; border-radius: 15px; border: 1px solid #C0C0C0; }}")
        self.layout = QHBoxLayout(self); self.layout.setContentsMargins(2, 2, 2, 2); self.layout.setSpacing(0)
        self.btn_pct = self._create_btn("Gr %", "pct"); self.btn_diff = self._create_btn("Diff", "diff"); self.btn_raw = self._create_btn("Raw", "raw")
        self.layout.addWidget(self.btn_pct); self.layout.addWidget(self.btn_diff); self.layout.addWidget(self.btn_raw)
        self._active_mode = "diff"; self._update_styles()
    def _create_btn(self, text, mode):
        btn = QPushButton(text); btn.setCursor(Qt.PointingHandCursor); btn.setCheckable(True)
        btn.clicked.connect(lambda: self.set_mode(mode)); btn.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); return btn
    def set_mode(self, mode): self._active_mode = mode; self._update_styles(); self.mode_changed.emit(mode)
    def update_theme_color(self, color): self.current_theme_color = color; self._update_styles()
    def _update_styles(self):
        base = "QPushButton { border: none; border-radius: 13px; background-color: transparent; color: #555555; } QPushButton:hover { background-color: rgba(255, 255, 255, 0.5); }"
        active = f"QPushButton {{ border: none; border-radius: 13px; background-color: {self.current_theme_color}; color: white; font-weight: bold; }}"
        self.btn_pct.setStyleSheet(active if self._active_mode == 'pct' else base)
        self.btn_diff.setStyleSheet(active if self._active_mode == 'diff' else base)
        self.btn_raw.setStyleSheet(active if self._active_mode == 'raw' else base)

class CategoryButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent); self.setFixedHeight(40); self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(f"QPushButton {{ text-align: left; padding-left: 15px; border: none; background-color: transparent; color: {config.BLACK}; font-weight: bold; font-family: '나눔스퀘어 네오 ExtraBold'; font-size: 11pt; }} QPushButton:hover {{ background-color: #D0D0D0; }}")

class SubMenuButton(QPushButton):
    def __init__(self, text, index, parent=None):
        super().__init__(text, parent); self.index = index; self.setFixedHeight(35); self.setCheckable(True); self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(f"QPushButton {{ text-align: left; padding-left: 30px; border: none; background-color: transparent; color: #555555; font-family: '나눔스퀘어 네오 Light'; font-size: 10pt; }} QPushButton:hover {{ background-color: #E8F0F2; }} QPushButton:checked {{ background-color: {config.HEADER_BG}; color: white; font-weight: bold; border-left: 4px solid {config.POS}; }}")
    def update_theme(self, theme):
        self.setStyleSheet(f"QPushButton {{ text-align: left; padding-left: 30px; border: none; background-color: transparent; color: #555555; font-family: '나눔스퀘어 네오 Light'; font-size: 10pt; }} QPushButton:hover {{ background-color: #E8F0F2; }} QPushButton:checked {{ background-color: {theme['dark']}; color: white; font-weight: bold; border-left: 4px solid {config.POS}; }}")

class Sidebar(QFrame):
    page_changed = pyqtSignal(int); theme_changed = pyqtSignal(dict)
    def __init__(self, parent=None):
        super().__init__(parent); self.setStyleSheet(f"background-color: {config.BG_MAIN}; border-right: 1px solid {config.BORDER};"); self.setFixedWidth(280)
        self.layout = QVBoxLayout(self); self.layout.setContentsMargins(0, 20, 0, 20); self.layout.setSpacing(5)
        title_box = QHBoxLayout(); title_box.setContentsMargins(15, 0, 0, 20)
        self.lbl_title = QLabel("Market\nIntelligence"); self.lbl_title.setFont(QFont("나눔스퀘어 네오 ExtraBold", 14)); self.lbl_title.setStyleSheet(f"color: {config.HEADER_BG};")
        title_box.addWidget(self.lbl_title); title_box.addStretch(1); self.layout.addLayout(title_box)
        self.btn_group = QButtonGroup(self); self.btn_group.setExclusive(True); self.menu_items = []
        
        self.add_category("Counterpoint")
        self.add_submenu("Weekly Analysis", 0)
        self.add_submenu("Monthly Analysis", 1)
        self.add_submenu("Flagship Model", 2)
        self.add_submenu("Region Brand", 3)
        self.add_submenu("Sell in Sell Thru", 4)
        
        self.add_spacing()
        self.add_category("Omdia")
        self.add_submenu("Market Tracker", 5)
        
        self.layout.addStretch(1)
        if self.menu_items: self.menu_items[0].setChecked(True)
        
    def add_category(self, text): btn = CategoryButton(text); self.layout.addWidget(btn)
    def add_submenu(self, text, index):
        btn = SubMenuButton(text, index); btn.clicked.connect(lambda: self.on_menu_clicked(index))
        self.layout.addWidget(btn); self.btn_group.addButton(btn); self.menu_items.append(btn)
    def add_spacing(self): self.layout.addSpacing(15)
    def on_menu_clicked(self, index):
        self.page_changed.emit(index)
        theme = config.THEMES["Counterpoint"]
        if index == 5: theme = config.THEMES["Omdia"]
        
        self.setStyleSheet(f"background-color: {theme['light']}; border-right: 1px solid {theme['border']};")
        self.lbl_title.setStyleSheet(f"color: {theme['dark']};")
        for btn in self.menu_items: btn.update_theme(theme)
        self.theme_changed.emit(theme)

class SwitchButton(QCheckBox):
    def __init__(self, parent=None, left_text="%", right_text="Vol"):
        super().__init__(parent); self.setFixedSize(130, 30); self.setCursor(Qt.PointingHandCursor)
        self._circle_position = 3; self.left_text = left_text; self.right_text = right_text; self.current_theme_color = config.HEADER_BG
        self.animation = QPropertyAnimation(self, b"circle_position", self); self.animation.setEasingCurve(QEasingCurve.OutBounce); self.animation.setDuration(300); self.stateChanged.connect(self.start_transition)
    def get_circle_position(self): return self._circle_position
    def set_circle_position(self, pos): self._circle_position = pos; self.update()
    circle_position = pyqtProperty(float, get_circle_position, set_circle_position)
    def start_transition(self, state):
        self.animation.stop()
        if state: self.animation.setEndValue(self.width() - 26)
        else: self.animation.setEndValue(3)
        self.animation.start()
    def update_theme_color(self, color): self.current_theme_color = color; self.update()
    def hitButton(self, pos): return self.contentsRect().contains(pos)
    def paintEvent(self, e):
        p = QPainter(self); p.setRenderHint(QPainter.Antialiasing); rect = QRectF(0, 0, self.width(), self.height())
        track_color = QColor(config.WHITE) if self.isChecked() else QColor(0,0,0, 80)
        p.setPen(Qt.NoPen); p.setBrush(track_color); p.drawRoundedRect(0, 0, self.width(), self.height(), 15, 15)
        p.setPen(QColor(self.current_theme_color) if self.isChecked() else QColor(config.WHITE))
        font = QFont("나눔스퀘어 네오 ExtraBold", 9); p.setFont(font)
        if self.isChecked(): p.drawText(QRectF(5, 0, self.width() - 30, self.height()), Qt.AlignCenter, self.right_text)
        else: p.drawText(QRectF(30, 0, self.width() - 35, self.height()), Qt.AlignCenter, self.left_text)
        p.setBrush(QColor(self.current_theme_color) if self.isChecked() else QColor(config.WHITE))
        p.drawEllipse(int(self._circle_position), 3, 24, 24); p.end()

def create_card(t, w, extra_widget=None):
    f = QFrame(); f.setStyleSheet(f"QFrame{{background:{config.CARD_BG};border:1px solid {config.BORDER};border-radius:10px;}}")
    v = QVBoxLayout(f); v.setSpacing(0)
    header_container = QWidget(); header_container.setObjectName("card_header")
    header_container.setStyleSheet(f"background:{config.HEADER_BG};border-top-left-radius:10px;border-top-right-radius:10px;"); header_container.setFixedHeight(35)
    header_layout = QHBoxLayout(header_container); header_layout.setContentsMargins(15, 0, 10, 0)
    h_lbl = QLabel(t); h_lbl.setStyleSheet(f"color:{config.WHITE};border:none;background:transparent;"); h_lbl.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); header_layout.addWidget(h_lbl); header_layout.addStretch(1)
    if extra_widget: header_layout.addWidget(extra_widget)
    v.setContentsMargins(0,0,0,0); v.addWidget(header_container)
    body_container = QWidget(); body_container.setStyleSheet("background:transparent;") 
    body_layout = QVBoxLayout(body_container); body_layout.setContentsMargins(5, 5, 5, 5); body_layout.addWidget(w); v.addWidget(body_container)
    return f

class BasePage(QWidget):
    def apply_theme(self, theme):
        dark_col = theme['dark']
        headers = self.findChildren(QWidget, "card_header")
        for h in headers: h.setStyleSheet(f"background:{dark_col};border-top-left-radius:10px;border-top-right-radius:10px;")
        buttons = self.findChildren(QPushButton)
        for btn in buttons:
            if btn.text() in ["Run Comparison", "Load Data", "Run Analysis", "Copy", "Reset", "Import from Active Excel", "Select", "Pivot", "Update Pivot", "Clear Fields", "Select Years", "Reset Filter", "Copy Data", "Run", "Load Sell-in Data", "Load Weekly"]:
                btn.setStyleSheet(f"background:{dark_col};color:{config.WHITE};border-radius:5px;")
            elif btn.text() in ["Download Result", "Check All", "Clear"]:
                btn.setStyleSheet(f"background:{config.WHITE};color:{dark_col};border:1px solid {dark_col};border-radius:5px;")
        drops = self.findChildren(FileDrop)
        for d in drops: d.setStyleSheet(f"QFrame{{background:{config.CARD_BG};border:2px dashed {dark_col};border-radius:10px;}}")
        toggles = self.findChildren(MultiStateToggle)
        for t in toggles: t.update_theme_color(dark_col)
        switches = self.findChildren(SwitchButton)
        for s in switches: s.update_theme_color(dark_col)
        p_bars = self.findChildren(QProgressBar)
        for p in p_bars:
            p.setStyleSheet(f"""QProgressBar {{ border: 1px solid {config.BORDER}; border-radius: 5px; text-align: center; }} QProgressBar::chunk {{ background-color: {dark_col}; width: 10px; }}""")
        for cb in self.findChildren(QComboBox):
            cb.setStyleSheet(f"QComboBox {{ background-color: white; color: black; border: 1px solid {dark_col}; border-radius: 5px; padding: 5px; }} QComboBox::drop-down {{ border: 0px; }}")

class ExcelSelectorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Excel File")
        self.setFixedSize(300, 150)
        self.selected_book = None
        self.setStyleSheet(f"""
            QDialog {{ background-color: {config.CARD_BG}; }}
            QLabel {{ color: {config.BLACK}; font-weight: bold; font-family: '나눔스퀘어 네오 ExtraBold'; font-size: 10pt; }}
            QComboBox {{ border: 1px solid {config.HEADER_BG}; padding: 5px; border-radius: 5px; font-family: '나눔스퀘어 네오 Light'; background-color: {config.WHITE}; }}
            QComboBox::drop-down {{ border: 0px; }}
            QPushButton {{ background-color: {config.HEADER_BG}; color: white; border-radius: 5px; padding: 5px; font-family: '나눔스퀘어 네오 ExtraBold'; font-size: 10pt; }}
            QPushButton:hover {{ background-color: #2980b9; }}
        """)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("여러 개의 엑셀이 열려 있습니다.\n분석할 파일을 선택해주세요:"))
        self.combo = QComboBox()
        try:
            self.book_list = [b for b in xw.books] 
        except:
            self.book_list = []
        for bk in self.book_list:
            self.combo.addItem(bk.name)
        layout.addWidget(self.combo)
        btn = QPushButton("Select")
        btn.clicked.connect(self.on_select)
        layout.addWidget(btn)

    def on_select(self):
        idx = self.combo.currentIndex()
        if idx >= 0 and idx < len(self.book_list):
            self.selected_book = self.book_list[idx] 
            self.accept()
        else:
            self.reject()

# --- Page Classes ---

class WeeklyPage(BasePage):
    def __init__(self):
        super().__init__()
        self.old = None; self.new = None; self.df = None; self.raw_data = None
        self.step = 0; self.settings = QSettings("MyCompany", "ExcelTool")
        self.init_ui()
        self.timer = QTimer(); self.timer.timeout.connect(self.anim)
        QTimer.singleShot(100, self.load_cache)

    def init_ui(self):
        self.run = QPushButton("Run Comparison"); self.run.setFixedSize(220, 45); self.run.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.run.clicked.connect(self.exec)
        self.dl = QPushButton("Download Result"); self.dl.setFixedSize(220, 45); self.dl.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.dl.setEnabled(False); self.dl.clicked.connect(self.download)
        header_layout = QHBoxLayout(); header_layout.addStretch(1); header_layout.addWidget(self.dl); header_layout.addWidget(self.run)
        
        self.drop_old = FileDrop("OLD FILE", self.set_old); self.drop_new = FileDrop("NEW FILE", self.set_new); input_layout = QHBoxLayout(); input_layout.addWidget(self.drop_old); input_layout.addWidget(self.drop_new)
        
        self.heatmap = HeatmapWidget(time_col="Week")
        self.toggle_heat = MultiStateToggle(); self.toggle_heat.mode_changed.connect(self.heatmap.set_mode)
        
        # [MODIFIED] Copy button connected to copy_heatmap_data
        self.btn_copy_heat = QPushButton("Copy"); self.btn_copy_heat.setFixedSize(60, 30); self.btn_copy_heat.setCursor(Qt.PointingHandCursor); self.btn_copy_heat.clicked.connect(self.copy_heatmap_data); self.btn_copy_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
        self.btn_reset_heat = QPushButton("Reset"); self.btn_reset_heat.setFixedSize(60, 30); self.btn_reset_heat.setCursor(Qt.PointingHandCursor); self.btn_reset_heat.clicked.connect(self.heatmap.reset_state); self.btn_reset_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
        
        # Month Selector
        self.lbl_month = QLabel("Max Month:")
        self.lbl_month.setStyleSheet(f"color: black; font-family: '나눔스퀘어 네오 ExtraBold'; font-size: 10pt;")
        self.spin_month = QSpinBox()
        self.spin_month.setRange(1, 12); self.spin_month.setValue(12)
        self.spin_month.setFixedWidth(50); self.spin_month.setStyleSheet("background: white; color: black; border-radius: 3px;")
        
        # 값 변경 시 즉시 필터링
        self.spin_month.valueChanged.connect(self.filter_data_by_month)

        hh_widget = QWidget(); hh_layout = QHBoxLayout(hh_widget); hh_layout.setContentsMargins(0,0,0,0)
        hh_layout.addWidget(self.lbl_month); hh_layout.addWidget(self.spin_month); hh_layout.addSpacing(10)
        hh_layout.addWidget(self.toggle_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_copy_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_reset_heat)
        
        c_heat = create_card("Sales Heatmap", self.heatmap, extra_widget=hh_widget)
        self.line_chart = LineChartWidget(time_col="Week"); self.toggle_line_cum = SwitchButton(left_text="Weekly", right_text="Cumulative"); self.toggle_line_cum.toggled.connect(self.update_line_chart_view); self.btn_copy_graph = QPushButton("Copy Data"); self.btn_copy_graph.setFixedSize(120, 35); self.btn_copy_graph.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_graph.setCursor(Qt.PointingHandCursor); self.btn_copy_graph.clicked.connect(self.line_chart.copy_current_data)
        gh_widget = QWidget(); gh_layout = QHBoxLayout(gh_widget); gh_layout.setContentsMargins(0,0,0,0); gh_layout.addWidget(self.toggle_line_cum); gh_layout.addWidget(self.btn_copy_graph)
        c_graph = create_card("Graph", self.line_chart, extra_widget=gh_widget)
        self.trend_chart = TrendWidget(time_col="Week"); self.toggle_trend = SwitchButton(left_text="Share", right_text="Vol"); self.toggle_trend.toggled.connect(self.trend_chart.set_mode); self.btn_copy_trend = QPushButton("Copy Data"); self.btn_copy_trend.setFixedSize(120, 35); self.btn_copy_trend.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_trend.setCursor(Qt.PointingHandCursor); self.btn_copy_trend.clicked.connect(self.trend_chart.copy_current_data)
        th_widget = QWidget(); th_layout = QHBoxLayout(th_widget); th_layout.setContentsMargins(0,0,0,0); th_layout.addWidget(self.toggle_trend); th_layout.addWidget(self.btn_copy_trend)
        c_trend = create_card("Trend", self.trend_chart, extra_widget=th_widget)
        self.heatmap.cell_clicked.connect(self.handle_heatmap_click); dashboard_layout = QHBoxLayout(); dashboard_layout.addWidget(c_heat, 2); dashboard_layout.addWidget(c_graph, 1); dashboard_layout.addWidget(c_trend, 1)
        self.t1 = QTableWidget(); self.t1.setFont(QFont("나눔스퀘어 네오 Light", 10)); c_detail = create_card("Detailed Comparison Results", self.t1); main_layout = QVBoxLayout(self); main_layout.setContentsMargins(20,20,20,20); main_layout.addLayout(header_layout); main_layout.addLayout(input_layout); main_layout.addLayout(dashboard_layout); main_layout.addWidget(c_detail, 1); self.apply_theme(config.THEMES["Counterpoint"])

    def load_cache(self):
        o, n = self.settings.value("weekly_old", ""), self.settings.value("weekly_new", "")
        if o and os.path.exists(o): self.old=o; self.drop_old.update_label(o)
        if n and os.path.exists(n): self.new=n; self.drop_new.update_label(n)
        if self.old and self.new: self.exec()
    def set_old(self, p): self.old = p; self.drop_old.update_label(p); self.settings.setValue("weekly_old", p); self.exec()
    def set_new(self, p):
        current_ver = extract_version(self.new); incoming_ver = extract_version(p)
        if incoming_ver >= current_ver: self.new = p; self.drop_new.update_label(p); self.settings.setValue("weekly_new", p); self.exec()
        else: QMessageBox.warning(self, "Warning", "Uploaded file is older than current.")
    def handle_heatmap_click(self, brand, region):
        self.selected_brand = brand; self.selected_region = region
        if self.heatmap.full_df is not None: self.trend_chart.update_chart(self.heatmap.full_df, brand, region); self.update_line_chart_view()
    def update_line_chart_view(self):
        if hasattr(self, 'selected_brand') and self.selected_brand and self.heatmap.full_df is not None: self.line_chart.update_chart(self.heatmap.full_df, self.selected_brand, self.selected_region, self.heatmap.p24, is_cumulative=self.toggle_line_cum.isChecked())
        else: self.line_chart.clear_plot()
    def anim(self): self.run.setText("Running"+"."*(self.step%4)); self.step+=1
    def exec(self):
        if not self.new: return
        self.run.setEnabled(False); self.step=0; self.timer.start(500)
        self.th = CompareThread(self.old, self.new); self.th.result.connect(self.show_result); self.th.error.connect(self.err); self.th.start()
    def err(self, e): self.timer.stop(); self.run.setText("Run Comparison"); self.run.setEnabled(True); QMessageBox.critical(self, "Error", e)
    
    def _extract_month_safe(self, df):
        date_col = None
        for c in df.columns:
            if str(c).strip().lower() == "date": date_col = c; break
        if date_col:
            date_str_series = df[date_col].astype(str).str.split('-').str[0].str.strip()
            dt_series = pd.to_datetime(date_str_series, format='%y%m%d', errors='coerce')
            if dt_series.isna().any():
                mask = dt_series.isna()
                dt_series[mask] = pd.to_datetime(date_str_series[mask], errors='coerce')
            return dt_series.dt.month.fillna(0).astype(int)
        
        month_col = None
        for c in df.columns:
            if str(c).strip().lower() == "month": month_col = c; break
        if month_col:
            dt_series = pd.to_datetime(df[month_col], errors='coerce')
            m_series = dt_series.dt.month
            if m_series.isna().any():
                def manual_parser(x):
                    s = str(x).strip().lower()
                    maps = {'jan':1, 'feb':2, 'mar':3, 'apr':4, 'may':5, 'jun':6, 
                            'jul':7, 'aug':8, 'sep':9, 'oct':10, 'nov':11, 'dec':12}
                    for k, v in maps.items():
                        if k in s: return v
                    try: return int(float(s))
                    except: return 0
                mask = m_series.isna()
                m_series[mask] = df.loc[mask, month_col].apply(manual_parser)
            return m_series.fillna(0).astype(int)
        return pd.Series([0]*len(df), index=df.index)

    def show_result(self, df, sumy, raw_data):
        self.timer.stop(); self.run.setText("Run Comparison"); self.run.setEnabled(True); self.df = df; self.dl.setEnabled(not df.empty)
        self.t1.setRowCount(len(df)); self.t1.setColumnCount(len(df.columns)); self.t1.setHorizontalHeaderLabels(df.columns)
        for i, r in df.iterrows(): 
            for j, c in enumerate(df.columns): self.t1.setItem(i, j, QTableWidgetItem(str(r[c])))
        
        self.raw_data = raw_data
        
        detected_month = 12
        try:
            for _, d in raw_data.items():
                if not d.empty:
                    target_df = d
                    if "Year" in d.columns:
                        max_year = d["Year"].max()
                        target_df = d[d["Year"] == max_year]
                    if not target_df.empty:
                        m_series = self._extract_month_safe(target_df)
                        max_m = m_series.max()
                        if max_m > 0: 
                            detected_month = max_m
                            break
        except: pass
        
        self.spin_month.blockSignals(True)
        self.spin_month.setValue(int(detected_month))
        self.spin_month.blockSignals(False)
        self.filter_data_by_month()

    def filter_data_by_month(self):
        if not self.raw_data: return
        target_month = self.spin_month.value()
        filtered_data = {}
        for sheet, df in self.raw_data.items():
            if not df.empty:
                temp_df = df.copy()
                temp_df["__MonthNum"] = self._extract_month_safe(temp_df)
                filtered_data[sheet] = temp_df[(temp_df["__MonthNum"] > 0) & (temp_df["__MonthNum"] <= target_month)].copy()
                del filtered_data[sheet]["__MonthNum"]
            else:
                filtered_data[sheet] = df
        self.heatmap.update_data(filtered_data)
        self.line_chart.clear_plot()
        self.trend_chart.clear_plot()

    # [MODIFIED] Copy Data: Transposed & Sorted Stacked View
    def copy_heatmap_data(self):
        if self.heatmap.p25 is None: 
            QMessageBox.warning(self, "Warning", "No data to copy.")
            return
        
        mode = self.heatmap.current_mode
        
        # [UPDATED] Sort Lists including Google & Japan
        desired_brands = ["Total", "Apple", "Samsung", "Xiaomi", "Oppo", "vivo", "Honor", "Huawei", "Google", "Others"]
        desired_regions = ["Total", "China", "India", "US", "W.Europe", "Japan", "Others"]
    
  def process_df(df):
              if df is None: return pd.DataFrame()
              # 행(Brand) 정렬
              existing_brands = [b for b in desired_brands if b in df.index]
              remaining_brands = [b for b in df.index if b not in desired_brands]
              df = df.reindex(existing_brands + remaining_brands)
              
              # 열(Region) 정렬
              existing_regions = [r for r in desired_regions if r in df.columns]
              remaining_regions = [r for r in df.columns if r not in desired_regions]
              df = df.reindex(columns=existing_regions + remaining_regions)
              
              # 전치 -> 행: Region, 열: Brand
              return df.T
  
          if mode == "diff" and self.heatmap.p24 is not None:
              # Diff 모드는 기존대로 유지 (필요하면 여기도 바꿀 수 있음)
              df24 = self.heatmap.p24 / 1000000.0
              df25 = self.heatmap.p25 / 1000000.0
              df_diff = df25 - df24
              
              t24 = process_df(df24)
              t25 = process_df(df25)
              tdiff = process_df(df_diff)
              
              combined = pd.concat([t24, t25, tdiff], axis=1, keys=['2024', '2025', 'Diff'])
              combined = combined.swaplevel(0, 1, axis=1)
              brands_in_col = t24.columns.tolist() 
              new_columns = []
              for b in brands_in_col:
                  new_columns.append((b, '2024'))
                  new_columns.append((b, '2025'))
                  new_columns.append((b, 'Diff'))
              
              combined = combined.reindex(columns=new_columns)
              combined.to_clipboard()
              QMessageBox.information(self, "Info", "Copied 2024/2025/Diff Data!")
  
          elif mode == "pct" and self.heatmap.p24 is not None:
              safe_p24 = self.heatmap.p24.mask(self.heatmap.p24 == 0)
              df_pct = (self.heatmap.p25 - self.heatmap.p24) / safe_p24
              export_df = process_df(df_pct)
              export_df.to_clipboard()
              QMessageBox.information(self, "Info", "Copied Growth %!")
              
          else:
              # [MODIFIED] Raw Volume Copy Logic (2025 & 2024 Stacked)
              df25 = self.heatmap.p25 / 1000000.0
              t25 = process_df(df25)
              
              if self.heatmap.p24 is not None:
                  df24 = self.heatmap.p24 / 1000000.0
                  t24 = process_df(df24)
              else:
                  t24 = pd.DataFrame()
              
              # Use CSV format with tabs for Excel copy
              s_25 = t25.to_csv(sep='\t')
              s_24 = t24.to_csv(sep='\t') if not t24.empty else ""
              
              final_text = f"2025\n{s_25}\n\n2024\n{s_24}"
              
              QApplication.clipboard().setText(final_text)
              QMessageBox.information(self, "Info", "Copied 2025 & 2024 Volume Tables!")
  
      def download(self):
          if self.df is None or self.df.empty: return
          p, _ = QFileDialog.getSaveFileName(self, "Save", "changed.xlsx", ".xlsx"); 
          if p: self.df.to_excel(p, index=False)
  
  class MonthlyPage(BasePage):
      def __init__(self): super().__init__(); self.old=None; self.new=None; self.df=None; self.step=0; self.settings = QSettings("MyCompany", "ExcelTool"); self.init_ui(); self.timer = QTimer(); self.timer.timeout.connect(self.anim); QTimer.singleShot(100, self.load_cache)
      def init_ui(self):
          self.run = QPushButton("Run Comparison"); self.run.setFixedSize(220, 45); self.run.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.run.clicked.connect(self.exec)
          self.dl = QPushButton("Download Result"); self.dl.setFixedSize(220, 45); self.dl.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.dl.setEnabled(False); self.dl.clicked.connect(self.download)
          header_layout = QHBoxLayout(); header_layout.addStretch(1); header_layout.addWidget(self.dl); header_layout.addWidget(self.run)
          self.drop_old = FileDrop("OLD FILE", self.set_old); self.drop_new = FileDrop("NEW FILE", self.set_new); input_layout = QHBoxLayout(); input_layout.addWidget(self.drop_old); input_layout.addWidget(self.drop_new)
          self.heatmap = HeatmapWidget(time_col="Month"); self.toggle_heat = MultiStateToggle(); self.toggle_heat.mode_changed.connect(self.heatmap.set_mode)
          self.btn_copy_heat = QPushButton("Copy"); self.btn_copy_heat.setFixedSize(60, 30); self.btn_copy_heat.setCursor(Qt.PointingHandCursor); self.btn_copy_heat.clicked.connect(self.heatmap.copy_data); self.btn_copy_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
          self.btn_reset_heat = QPushButton("Reset"); self.btn_reset_heat.setFixedSize(60, 30); self.btn_reset_heat.setCursor(Qt.PointingHandCursor); self.btn_reset_heat.clicked.connect(self.heatmap.reset_state); self.btn_reset_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
          hh_widget = QWidget(); hh_layout = QHBoxLayout(hh_widget); hh_layout.setContentsMargins(0,0,0,0); hh_layout.addWidget(self.toggle_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_copy_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_reset_heat)
          c_heat = create_card("Sales Heatmap", self.heatmap, extra_widget=hh_widget)
          self.line_chart = LineChartWidget(time_col="Month"); self.toggle_line_cum = SwitchButton(left_text="Monthly", right_text="Cumulative"); self.toggle_line_cum.toggled.connect(self.update_line_chart_view); self.btn_copy_graph = QPushButton("Copy Data"); self.btn_copy_graph.setFixedSize(120, 35); self.btn_copy_graph.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_graph.setCursor(Qt.PointingHandCursor); self.btn_copy_graph.clicked.connect(self.line_chart.copy_current_data)
          gh_widget = QWidget(); gh_layout = QHBoxLayout(gh_widget); gh_layout.setContentsMargins(0,0,0,0); gh_layout.addWidget(self.toggle_line_cum); gh_layout.addWidget(self.btn_copy_graph)
          c_graph = create_card("Graph", self.line_chart, extra_widget=gh_widget)
          self.trend_chart = TrendWidget(time_col="Month"); self.toggle_trend = SwitchButton(left_text="Share", right_text="Vol"); self.toggle_trend.toggled.connect(self.trend_chart.set_mode); self.btn_copy_trend = QPushButton("Copy Data"); self.btn_copy_trend.setFixedSize(120, 35); self.btn_copy_trend.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_trend.setCursor(Qt.PointingHandCursor); self.btn_copy_trend.clicked.connect(self.trend_chart.copy_current_data)
          th_widget = QWidget(); th_layout = QHBoxLayout(th_widget); th_layout.setContentsMargins(0,0,0,0); th_layout.addWidget(self.toggle_trend); th_layout.addWidget(self.btn_copy_trend)
          c_trend = create_card("Trend", self.trend_chart, extra_widget=th_widget)
          self.heatmap.cell_clicked.connect(self.handle_heatmap_click); dashboard_layout = QHBoxLayout(); dashboard_layout.addWidget(c_heat, 2); dashboard_layout.addWidget(c_graph, 1); dashboard_layout.addWidget(c_trend, 1)
          self.t1 = QTableWidget(); self.t1.setFont(QFont("나눔스퀘어 네오 Light", 10)); c_detail = create_card("Detailed Comparison Results", self.t1); main_layout = QVBoxLayout(self); main_layout.setContentsMargins(20,20,20,20); main_layout.addLayout(header_layout); main_layout.addLayout(input_layout); main_layout.addLayout(dashboard_layout); main_layout.addWidget(c_detail, 1); self.apply_theme(config.THEMES["Counterpoint"])
      def load_cache(self):
          o, n = self.settings.value("monthly_old", ""), self.settings.value("monthly_new", "")
          if o: self.old=o; self.drop_old.update_label(o)
          if n: self.new=n; self.drop_new.update_label(n)
          if self.old and self.new: self.exec()
      def set_old(self, p): self.old = p; self.drop_old.update_label(p); self.exec()
      def set_new(self, p): self.new = p; self.drop_new.update_label(p); self.exec()
      def handle_heatmap_click(self, brand, region):
          self.selected_brand = brand; self.selected_region = region
          if self.heatmap.full_df is not None: self.line_chart.update_chart(self.heatmap.full_df, brand, region, self.heatmap.p24, is_cumulative=self.toggle_line_cum.isChecked()); self.trend_chart.update_chart(self.heatmap.full_df, brand, region)
          else: self.trend_chart.clear_plot()
      def update_line_chart_view(self):
          if hasattr(self, 'selected_brand') and self.selected_brand and self.heatmap.full_df is not None: self.line_chart.update_chart(self.heatmap.full_df, self.selected_brand, self.selected_region, self.heatmap.p24, is_cumulative=self.toggle_line_cum.isChecked())
      def anim(self): self.run.setText("Running"+"."*(self.step%4)); self.step+=1
      def exec(self):
          if not self.new: return
          self.run.setEnabled(False); self.step=0; self.timer.start(500)
          self.th = MonthlyCompareThread(self.old, self.new); self.th.result.connect(self.show_result); self.th.error.connect(self.err); self.th.start()
      def err(self, e): self.timer.stop(); self.run.setText("Run Comparison"); self.run.setEnabled(True); QMessageBox.critical(self, "Error", e)
      def show_result(self, df, sumy, raw_data):
          self.timer.stop(); self.run.setText("Run Comparison"); self.run.setEnabled(True); self.df = df; self.dl.setEnabled(not df.empty)
          self.t1.setRowCount(len(df)); self.t1.setColumnCount(len(df.columns)); self.t1.setHorizontalHeaderLabels(df.columns)
          for i, r in df.iterrows(): 
              for j, c in enumerate(df.columns): self.t1.setItem(i, j, QTableWidgetItem(str(r[c])))
          self.line_chart.clear_plot(); self.trend_chart.clear_plot(); self.heatmap.update_data(raw_data)
      def download(self):
          if self.df is None or self.df.empty: return
          p, _ = QFileDialog.getSaveFileName(self, "Save", "changed.xlsx", ".xlsx"); 
          if p: self.df.to_excel(p, index=False)
  
  class FlagshipPage(BasePage):
      def __init__(self): super().__init__(); self.path=None; self.step=0; self.settings = QSettings("MyCompany", "ExcelTool"); self.init_ui(); self.timer = QTimer(); self.timer.timeout.connect(self.anim); QTimer.singleShot(100, self.load_cache)
      def init_ui(self):
          self.run = QPushButton("Load Data"); self.run.setFixedSize(220, 45); self.run.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.run.clicked.connect(self.exec)
          self.toggle_cat = SwitchButton(left_text="Foldable", right_text="Smartphone"); self.toggle_cat.setChecked(True); self.toggle_cat.toggled.connect(self.update_views)
          self.btn_year_select = QPushButton("Select Years"); self.btn_year_select.setFixedSize(120, 30); self.btn_year_select.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_year_menu = QMenu(self); self.btn_year_select.setMenu(self.btn_year_menu)
          header_layout = QHBoxLayout(); header_layout.addStretch(1); header_layout.addWidget(QLabel("Category:", font=QFont("나눔스퀘어 네오 ExtraBold", 10))); header_layout.addWidget(self.toggle_cat); header_layout.addSpacing(20); header_layout.addWidget(self.btn_year_select); header_layout.addSpacing(10); header_layout.addWidget(self.run)
          self.drop_file = FileDrop("FLAGSHIP FILE", self.set_path); input_layout = QHBoxLayout(); input_layout.addWidget(self.drop_file)
          self.heatmap = HeatmapWidget(time_col="Month")
          hh_widget = QWidget(); hh_layout = QHBoxLayout(hh_widget); hh_layout.setContentsMargins(0,0,0,0); self.btn_copy_heat = QPushButton("Copy"); self.btn_copy_heat.setFixedSize(60, 30); self.btn_copy_heat.setCursor(Qt.PointingHandCursor); self.btn_copy_heat.clicked.connect(self.heatmap.copy_data); self.btn_copy_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); hh_layout.addWidget(self.btn_copy_heat)
          c_heat = create_card("Brand Volume (Mu)", self.heatmap, extra_widget=hh_widget)
          self.launch_chart = LaunchTrendWidget()
          launch_header_widget = QWidget(); lh_layout = QHBoxLayout(launch_header_widget); lh_layout.setContentsMargins(0, 0, 0, 0)
          self.spin_max_month = QSpinBox(); self.spin_max_month.setRange(1, 120); self.spin_max_month.setValue(24); self.spin_max_month.setPrefix("Max T+"); self.spin_max_month.valueChanged.connect(self.update_launch_chart)
          self.toggle_cumulative = SwitchButton(left_text="Monthly", right_text="Cumulative"); self.toggle_cumulative.toggled.connect(self.update_launch_chart)
          self.btn_copy_chart = QPushButton("Copy"); self.btn_copy_chart.setFixedSize(60, 30); self.btn_copy_chart.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_chart.setCursor(Qt.PointingHandCursor); self.btn_copy_chart.clicked.connect(self.launch_chart.copy_current_data)
          lh_layout.addWidget(self.spin_max_month); lh_layout.addSpacing(10); lh_layout.addWidget(self.toggle_cumulative); lh_layout.addSpacing(10); lh_layout.addWidget(self.btn_copy_chart)
          c_launch = create_card("Model Launch Trend", self.launch_chart, extra_widget=launch_header_widget)
          self.model_list = QListWidget(); self.model_list.setFont(QFont("나눔스퀘어 네오 Light", 9)); self.model_list.itemChanged.connect(self.on_item_changed)
          btn_box = QHBoxLayout(); self.btn_check_all = QPushButton("Check All"); self.btn_check_all.clicked.connect(self.check_all); self.btn_clear = QPushButton("Clear"); self.btn_clear.clicked.connect(self.clear_checks); btn_box.addWidget(self.btn_check_all); btn_box.addWidget(self.btn_clear)
          list_container = QWidget(); v_list = QVBoxLayout(list_container); v_list.setContentsMargins(0,0,0,0); v_list.addWidget(self.model_list); v_list.addLayout(btn_box); c_list = create_card("Model List", list_container)
          self.heatmap.cell_clicked.connect(self.handle_heatmap_click); dashboard_layout = QHBoxLayout(); dashboard_layout.addWidget(c_heat, 3); dashboard_layout.addWidget(c_launch, 4); dashboard_layout.addWidget(c_list, 1)
          main_layout = QVBoxLayout(self); main_layout.setContentsMargins(20,20,20,20); main_layout.addLayout(header_layout); main_layout.addLayout(input_layout); main_layout.addLayout(dashboard_layout)
          self.apply_theme(config.THEMES["Counterpoint"])
      def load_cache(self):
          p = self.settings.value("flagship_path", "")
          if p: self.path=p; self.drop_file.update_label(p); self.exec()
      def set_path(self, p): self.path = p; self.drop_file.update_label(p); self.exec()
      def anim(self): self.run.setText("Loading"+"."*(self.step%4)); self.step+=1
      def exec(self):
          if not self.path: return
          self.run.setEnabled(False); self.step=0; self.timer.start(500)
          self.th = FlagshipThread(self.path); self.th.result.connect(self.show_result); self.th.error.connect(self.err); self.th.start()
      def err(self, e): self.timer.stop(); self.run.setText("Load Data"); self.run.setEnabled(True); QMessageBox.critical(self, "Error", e)
      def show_result(self, df):
          self.timer.stop(); self.run.setText("Load Data"); self.run.setEnabled(True); self.full_df=df; self.all_years = sorted(df['Date'].dt.year.unique().astype(str), reverse=True); self.selected_years = self.all_years[:3]; self.update_year_menu(); self.update_views()
      def update_year_menu(self):
          self.btn_year_menu.clear()
          for year in self.all_years: action = QAction(year, self); action.setCheckable(True); action.setChecked(year in self.selected_years); action.triggered.connect(self.on_year_toggled); self.btn_year_menu.addAction(action)
      def on_year_toggled(self):
          selected = []; 
          for action in self.btn_year_menu.actions(): 
              if action.isChecked(): selected.append(action.text())
          self.selected_years = selected; self.update_views()
      def update_views(self):
          if self.full_df is None: return
          cat = "Smartphone" if self.toggle_cat.isChecked() else "Foldable"
          self.heatmap.update_data_flagship(self.full_df, cat, self.selected_years); self.launch_chart.clear_plot(); self.model_list.clear()
      def handle_heatmap_click(self, brand, _): 
          if brand is None or brand == "Total": self.launch_chart.clear_plot(); self.model_list.clear(); return
          self.current_brand = brand; self.populate_model_list(); self.update_launch_chart()
      def populate_model_list(self):
          self.model_list.clear()
          if self.full_df is None or not self.current_brand: return
          cat = "Smartphone" if self.toggle_cat.isChecked() else "Foldable"
          models = sorted(self.full_df[(self.full_df['Brand'] == self.current_brand) & (self.full_df['Category'] == cat)]['Model'].unique())
          for m in models: item = QListWidgetItem(m); item.setFlags(item.flags() | Qt.ItemIsUserCheckable); item.setCheckState(Qt.Checked); self.model_list.addItem(item)
      def on_item_changed(self, item): self.update_launch_chart()
      def check_all(self):
          for i in range(self.model_list.count()): self.model_list.item(i).setCheckState(Qt.Checked)
          self.update_launch_chart()
      def clear_checks(self):
          for i in range(self.model_list.count()): self.model_list.item(i).setCheckState(Qt.Unchecked)
          self.update_launch_chart()
      def update_launch_chart(self):
          if not hasattr(self, 'current_brand') or not self.current_brand: return
          cat = "Smartphone" if self.toggle_cat.isChecked() else "Foldable"
          visible_models = []
          for i in range(self.model_list.count()):
              if self.model_list.item(i).checkState() == Qt.Checked: visible_models.append(self.model_list.item(i).text())
          self.launch_chart.update_chart(self.full_df, self.current_brand, cat, visible_models, self.spin_max_month.value(), self.toggle_cumulative.isChecked())
    
class RegionBrandPage(BasePage):
    def __init__(self): super().__init__(); self.path=None; self.step=0; self.settings = QSettings("MyCompany", "ExcelTool"); self.init_ui(); self.timer = QTimer(); self.timer.timeout.connect(self.anim); QTimer.singleShot(100, self.load_cache)
    def init_ui(self):
        self.run = QPushButton("Run Analysis"); self.run.setFixedSize(220, 45); self.run.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.run.clicked.connect(self.exec)
        header_layout = QHBoxLayout(); header_layout.addStretch(1); header_layout.addWidget(self.run)
        self.drop_file = FileDrop("REGION BRAND FILE", self.set_path); input_layout = QHBoxLayout(); input_layout.addWidget(self.drop_file)
        self.heatmap = HeatmapWidget(time_col="Month"); self.toggle_heat = MultiStateToggle(); self.toggle_heat.mode_changed.connect(self.heatmap.set_mode)
        self.btn_copy_heat = QPushButton("Copy"); self.btn_copy_heat.setFixedSize(60, 30); self.btn_copy_heat.setCursor(Qt.PointingHandCursor); self.btn_copy_heat.clicked.connect(self.heatmap.copy_data); self.btn_copy_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
        self.btn_reset_heat = QPushButton("Reset"); self.btn_reset_heat.setFixedSize(60, 30); self.btn_reset_heat.setCursor(Qt.PointingHandCursor); self.btn_reset_heat.clicked.connect(self.heatmap.reset_state); self.btn_reset_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
        hh_widget = QWidget(); hh_layout = QHBoxLayout(hh_widget); hh_layout.setContentsMargins(0,0,0,0); hh_layout.addWidget(self.toggle_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_copy_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_reset_heat)
        c_heat = create_card("Sales Heatmap", self.heatmap, extra_widget=hh_widget)
        self.line_chart = LineChartWidget(time_col="Month"); self.btn_copy_graph = QPushButton("Copy Data"); self.btn_copy_graph.setFixedSize(120, 35); self.btn_copy_graph.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_graph.setCursor(Qt.PointingHandCursor); self.btn_copy_graph.clicked.connect(self.line_chart.copy_current_data); c_graph = create_card("Graph", self.line_chart, extra_widget=self.btn_copy_graph)
        self.trend_chart = TrendWidget(time_col="Month"); self.toggle_trend = SwitchButton(left_text="Share", right_text="Vol"); self.toggle_trend.toggled.connect(self.trend_chart.set_mode); self.btn_copy_trend = QPushButton("Copy Data"); self.btn_copy_trend.setFixedSize(120, 35); self.btn_copy_trend.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_trend.setCursor(Qt.PointingHandCursor); self.btn_copy_trend.clicked.connect(self.trend_chart.copy_current_data)
        th_widget = QWidget(); th_layout = QHBoxLayout(th_widget); th_layout.setContentsMargins(0,0,0,0); th_layout.addWidget(self.toggle_trend); th_layout.addWidget(self.btn_copy_trend)
        c_trend = create_card("Trend", self.trend_chart, extra_widget=th_widget)
        self.heatmap.cell_clicked.connect(self.handle_heatmap_click); dashboard_layout = QHBoxLayout(); dashboard_layout.addWidget(c_heat, 2); dashboard_layout.addWidget(c_graph, 1); dashboard_layout.addWidget(c_trend, 1)
        main_layout = QVBoxLayout(self); main_layout.setContentsMargins(20,20,20,20); main_layout.addLayout(header_layout); main_layout.addLayout(input_layout); main_layout.addLayout(dashboard_layout)
        self.apply_theme(config.THEMES["Counterpoint"])
    def set_path(self, p): self.path = p; self.drop_file.update_label(p); self.exec()
    def load_cache(self):
        p = self.settings.value("region_path", "")
        if p: self.path=p; self.drop_file.update_label(p); self.exec()
    def anim(self): self.run.setText("Analyzing"+"."*(self.step%4)); self.step+=1
    def exec(self):
        if not self.path: return
        self.run.setEnabled(False); self.step=0; self.timer.start(500)
        self.th = RegionBrandThread(self.path); self.th.result.connect(self.show_result); self.th.error.connect(self.err); self.th.start()
    def err(self, e): self.timer.stop(); self.run.setText("Run Analysis"); self.run.setEnabled(True); QMessageBox.critical(self, "Error", e)
    def show_result(self, data):
        self.timer.stop(); self.run.setText("Run Analysis"); self.run.setEnabled(True); self.heatmap.update_data(data)
    def handle_heatmap_click(self, brand, region):
        if self.heatmap.full_df is not None: self.line_chart.update_chart(self.heatmap.full_df, brand, region, self.heatmap.p24); self.trend_chart.update_chart(self.heatmap.full_df, brand, region)


class SellInPage(BasePage):
    def __init__(self):
        super().__init__()
        self.sellin_path = None; self.weekly_path = None
        self.sellin_df = None; self.weekly_data = None 
        self.settings = QSettings("MyCompany", "ExcelTool")
        self.init_ui()
        self.timer = QTimer(); self.timer.timeout.connect(self.anim)
        QTimer.singleShot(100, self.load_cache)

    def init_ui(self):
        # Header Buttons
        self.btn_load_sellin = QPushButton("Load Sell-in"); self.btn_load_sellin.setFixedSize(150, 45); self.btn_load_sellin.clicked.connect(lambda: self.exec('sellin'))
        self.btn_load_weekly = QPushButton("Load Weekly"); self.btn_load_weekly.setFixedSize(150, 45); self.btn_load_weekly.clicked.connect(lambda: self.exec('weekly'))
        
        # Apply Styles
        for btn in [self.btn_load_sellin, self.btn_load_weekly]:
            btn.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10))
            
        header_layout = QHBoxLayout(); header_layout.addStretch(1)
        header_layout.addWidget(self.btn_load_sellin); header_layout.addWidget(self.btn_load_weekly)
        
        # Dual Drop Zones
        input_layout = QHBoxLayout()
        self.drop_sellin = FileDrop("SELL-IN (Global SP)", lambda p: self.set_path(p, 'sellin'))
        self.drop_weekly = FileDrop("SELL-THRU (Weekly)", lambda p: self.set_path(p, 'weekly'))
        input_layout.addWidget(self.drop_sellin); input_layout.addWidget(self.drop_weekly)
        
        # Controls (Max Month, Copy)
        self.lbl_month = QLabel("Max Month:")
        self.lbl_month.setStyleSheet(f"color: black; font-family: '나눔스퀘어 네오 ExtraBold'; font-size: 10pt;")
        self.spin_month = QSpinBox(); self.spin_month.setRange(1, 12); self.spin_month.setValue(12); self.spin_month.setFixedWidth(50)
        self.spin_month.valueChanged.connect(self.update_view) # Trigger update on change
        
        self.heatmap = HeatmapWidget(time_col="Month") 
        self.toggle_heat = MultiStateToggle(); self.toggle_heat.mode_changed.connect(self.heatmap.set_mode)
        self.btn_copy_heat = QPushButton("Copy"); self.btn_copy_heat.setFixedSize(60, 30); self.btn_copy_heat.setCursor(Qt.PointingHandCursor); self.btn_copy_heat.clicked.connect(self.copy_heatmap_data); self.btn_copy_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9))
        
        hh_widget = QWidget(); hh_layout = QHBoxLayout(hh_widget); hh_layout.setContentsMargins(0,0,0,0)
        hh_layout.addWidget(self.lbl_month); hh_layout.addWidget(self.spin_month); hh_layout.addSpacing(10)
        hh_layout.addWidget(self.toggle_heat); hh_layout.addSpacing(5); hh_layout.addWidget(self.btn_copy_heat)
        
        c_heat = create_card("Sell-in YoY Heatmap (2024 vs 2025)", self.heatmap, extra_widget=hh_widget)
        
        main_layout = QVBoxLayout(self); main_layout.setContentsMargins(20,20,20,20)
        main_layout.addLayout(header_layout); main_layout.addLayout(input_layout); main_layout.addWidget(c_heat)
        self.apply_theme(config.THEMES["Counterpoint"])

    def load_cache(self):
        p_si = self.settings.value("sellin_path", "")
        p_wk = self.settings.value("sellin_weekly_path", "") 
        if p_si and os.path.exists(p_si): self.sellin_path=p_si; self.drop_sellin.update_label(p_si); self.exec('sellin')
        if p_wk and os.path.exists(p_wk): self.weekly_path=p_wk; self.drop_weekly.update_label(p_wk); self.exec('weekly')

    def set_path(self, p, type_):
        if type_ == 'sellin': self.sellin_path = p; self.drop_sellin.update_label(p); self.settings.setValue("sellin_path", p)
        else: self.weekly_path = p; self.drop_weekly.update_label(p); self.settings.setValue("sellin_weekly_path", p)
        self.exec(type_)

    def anim(self): 
        pass 

    def exec(self, type_):
        path = self.sellin_path if type_ == 'sellin' else self.weekly_path
        if not path: return
        
        if type_ == 'sellin':
            self.btn_load_sellin.setEnabled(False); self.btn_load_sellin.setText("Loading...")
            self.th_si = SellInThread(path)
            self.th_si.result.connect(self.on_sellin_loaded)
            self.th_si.error.connect(lambda e: self.err(e, 'sellin'))
            self.th_si.start()
        else:
            from data_loader import WeeklySimpleThread 
            self.btn_load_weekly.setEnabled(False); self.btn_load_weekly.setText("Loading...")
            self.th_wk = WeeklySimpleThread(path)
            self.th_wk.result.connect(self.on_weekly_loaded)
            self.th_wk.error.connect(lambda e: self.err(e, 'weekly'))
            self.th_wk.start()

    def err(self, e, type_):
        if type_ == 'sellin': self.btn_load_sellin.setText("Load Sell-in"); self.btn_load_sellin.setEnabled(True)
        else: self.btn_load_weekly.setText("Load Weekly"); self.btn_load_weekly.setEnabled(True)
        QMessageBox.critical(self, "Error", e)

    def on_sellin_loaded(self, df):
        self.btn_load_sellin.setText("Load Sell-in"); self.btn_load_sellin.setEnabled(True)
        if df.empty: return
        self.sellin_df = df
        
        max_year = df["Year"].max()
        if not pd.isna(max_year):
            max_month = df[df["Year"] == max_year]["Month"].max()
            self.spin_month.blockSignals(True)
            self.spin_month.setValue(int(max_month))
            self.spin_month.blockSignals(False)
        
        self.update_view()

    def on_weekly_loaded(self, data_dict):
        self.btn_load_weekly.setText("Load Weekly"); self.btn_load_weekly.setEnabled(True)
        if not data_dict: return
        self.weekly_data = data_dict
        self.update_view()

    def update_view(self):
        if self.sellin_df is not None:
            self.update_heatmap_logic()
        
        if self.weekly_data is not None:
            self.print_weekly_stats()

    def print_weekly_stats(self):
        target_month = self.spin_month.value()
        display_year = 2025
        print(f"\n[Weekly Sell-Through Analysis] Max Month: {target_month} (Target Year: {display_year})")
        print("-" * 60)
        
        target_regions = {
            "China": "Basefile_China",
            "India": "Basefile_India",
            "US": "Basefile_US",
            "W.Europe": "Basefile_Europe"
        }
        
        for region_name, sheet_name in target_regions.items():
            if sheet_name not in self.weekly_data:
                print(f"Region: {region_name:<10} | Data Not Found")
                continue
                
            df = self.weekly_data[sheet_name].copy()
            if "Sales" not in df.columns:
                print(f"Region: {region_name:<10} | 'Sales' Column Missing")
                continue

            df["Sales"] = pd.to_numeric(df["Sales"], errors='coerce').fillna(0)
            
            def parse_month(x):
                try:
                    s = str(x).strip().lower()
                    if s.isdigit(): return int(s)
                    try: return int(float(s))
                    except: pass
                    maps = {'jan':1, 'feb':2, 'mar':3, 'apr':4, 'may':5, 'jun':6, 
                            'jul':7, 'aug':8, 'sep':9, 'oct':10, 'nov':11, 'dec':12}
                    for k, v in maps.items():
                        if k in s: return v
                    return 0
                except: return 0

            if "Month" in df.columns:
                df["__MonthNum"] = df["Month"].apply(parse_month)
            else:
                df["__MonthNum"] = 0

            if "Year" in df.columns:
                df["Year"] = pd.to_numeric(df["Year"], errors='coerce').fillna(0)
                if 2025 in df["Year"].unique():
                    df = df[df["Year"] == 2025]
            
            df_target = df[(df["__MonthNum"] > 0) & (df["__MonthNum"] <= target_month)]
            total_sales = df_target["Sales"].sum()
            
            # [MODIFIED] Print both Raw and x1M for debugging
            print(f"Region: {region_name:<10} | Year: {display_year} | Month: 1~{target_month} | Total Sales(Raw): {total_sales:,.2f} | x1M: {total_sales * 1000000:,.2f}")

        print("-" * 60)

    def update_heatmap_logic(self):
        if self.sellin_df is None: return
        target_month = self.spin_month.value()
        
        df_filtered = self.sellin_df[self.sellin_df["Month"] <= target_month].copy()
        if df_filtered.empty: return

        max_year = int(self.sellin_df["Year"].max())
        prev_year = max_year - 1
        
        df_curr = df_filtered[df_filtered["Year"] == max_year]
        df_prev = df_filtered[df_filtered["Year"] == prev_year]
        
        def group_brand(name):
            n = str(name).strip().upper()
            if n == "SAMSUNG": return "MX"
            if n in ["OPPO", "REALME", "ONEPLUS"]: return "Oppo"
            if "TRANSSION" in n: return "Transsion"
            if n in ["APPLE", "XIAOMI", "VIVO", "HONOR", "HUAWEI"]: return name.strip()
            if n == "APPLE": return "Apple"
            if n == "XIAOMI": return "Xiaomi"
            if n == "VIVO": return "Vivo"
            if n == "HONOR": return "Honor"
            if n == "HUAWEI": return "Huawei"
            return "Others_Calc"

        def make_pivot(d):
            if d.empty: return pd.DataFrame()
            temp = d.copy()
            temp["Brand_Group"] = temp["Brand"].apply(group_brand)
            p = temp.pivot_table(index="Region", columns="Brand_Group", values="Sales", aggfunc="sum", fill_value=0)
            
            req_cols = ["Apple", "MX", "Xiaomi", "Oppo", "Vivo", "Transsion", "Honor", "Huawei"]
            for c in req_cols: 
                if c not in p.columns: p[c] = 0
            
            p["Total"] = p.sum(axis=1)
            specified_sum = p[req_cols].sum(axis=1)
            p["Others"] = p["Total"] - specified_sum
            
            final_cols = ["Total"] + req_cols + ["Others"]
            p = p.reindex(columns=final_cols, fill_value=0)
            
            if "Total" in p.index:
                sub_regions_sum = pd.Series(0, index=p.columns)
                for r in ["China", "India", "US", "W.Europe"]:
                    if r in p.index: sub_regions_sum += p.loc[r]
                p.loc["Others"] = p.loc["Total"] - sub_regions_sum
            
            final_rows = ["Total", "China", "India", "US", "W.Europe", "Others"]
            p = p.reindex(final_rows, fill_value=0)
            return p

        p24 = make_pivot(df_prev)
        p25 = make_pivot(df_curr)
        
        self.heatmap.p24 = p24
        self.heatmap.p25 = p25
        self.heatmap.refresh_view()

    def copy_heatmap_data(self):
        if self.heatmap.p25 is None: 
            QMessageBox.warning(self, "Warning", "No data to copy.")
            return
        
        mode = self.heatmap.current_mode
        
        desired_brands = ["Total", "Apple", "MX", "Xiaomi", "Oppo", "Vivo", "Transsion", "Honor", "Huawei", "Others"]
        desired_regions = ["Total", "China", "India", "US", "W.Europe", "Others"]

        def process_df(df):
            df = df.reindex(index=desired_regions, columns=desired_brands)
            return df.T

        if mode == "diff" and self.heatmap.p24 is not None:
            df24 = self.heatmap.p24 / 1000000.0
            df25 = self.heatmap.p25 / 1000000.0
            df_diff = df25 - df24
            
            t24 = process_df(df24)
            t25 = process_df(df25)
            tdiff = process_df(df_diff)
            
            combined = pd.concat([t24, t25, tdiff], axis=1, keys=['2024', '2025', 'Diff'])
            combined = combined.swaplevel(0, 1, axis=1)
            new_columns = []
            for r in desired_regions:
                if r in t24.columns:
                    new_columns.append((r, '2024'))
                    new_columns.append((r, '2025'))
                    new_columns.append((r, 'Diff'))
            combined = combined.reindex(columns=new_columns)
            combined.to_clipboard()
            QMessageBox.information(self, "Info", "Copied 2024/2025/Diff Data!")

        elif mode == "pct" and self.heatmap.p24 is not None:
            safe_p24 = self.heatmap.p24.mask(self.heatmap.p24 == 0)
            df_pct = (self.heatmap.p25 - self.heatmap.p24) / safe_p24
            export_df = process_df(df_pct)
            export_df.to_clipboard()
            QMessageBox.information(self, "Info", "Copied Growth %!")
            
        else:
            df_vol = self.heatmap.p25 / 1000000.0
            export_df = process_df(df_vol)
            export_df.to_clipboard()
            QMessageBox.information(self, "Info", "Copied Volume!")

  
class OmdiaPage(BasePage):
    def __init__(self): super().__init__(); self.path=None; self.full_df=None; self.step=0; self.all_years=[]; self.selected_years=[]; self.current_brand=None; self.settings = QSettings("MyCompany", "ExcelTool"); self.init_ui(); self.timer = QTimer(); self.timer.timeout.connect(self.anim); QTimer.singleShot(100, self.load_cache)
    def init_ui(self):
        self.run = QPushButton("Load Data"); self.run.setFixedSize(220, 45); self.run.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.run.clicked.connect(self.exec)
        self.toggle_cat = SwitchButton(left_text="Foldable", right_text="Smartphone"); self.toggle_cat.setChecked(True); self.toggle_cat.toggled.connect(self.update_views)
        self.btn_year_select = QPushButton("Select Years"); self.btn_year_select.setFixedSize(120, 30); self.btn_year_select.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_year_menu = QMenu(self); self.btn_year_select.setMenu(self.btn_year_menu)
        header_layout = QHBoxLayout(); header_layout.addStretch(1); header_layout.addWidget(QLabel("Category:", font=QFont("나눔스퀘어 네오 ExtraBold", 10))); header_layout.addWidget(self.toggle_cat); header_layout.addSpacing(20); header_layout.addWidget(self.btn_year_select); header_layout.addSpacing(10); header_layout.addWidget(self.run)
        self.drop_file = FileDrop("OMDIA RAW FILE", self.set_path); input_layout = QHBoxLayout(); input_layout.addWidget(self.drop_file)
        self.heatmap = HeatmapWidget(time_col="Quarter")
        hh_widget = QWidget(); hh_layout = QHBoxLayout(hh_widget); hh_layout.setContentsMargins(0,0,0,0); self.btn_copy_heat = QPushButton("Copy"); self.btn_copy_heat.setFixedSize(60, 30); self.btn_copy_heat.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_heat.setCursor(Qt.PointingHandCursor); self.btn_copy_heat.clicked.connect(self.heatmap.copy_data); hh_layout.addWidget(self.btn_copy_heat); c_heat = create_card("Vendor Volume (Mu) - Quarterly", self.heatmap, extra_widget=hh_widget)
        self.launch_table = LaunchTableWidget()
        launch_header_widget = QWidget(); lh_layout = QHBoxLayout(launch_header_widget); lh_layout.setContentsMargins(0, 0, 0, 0)
        self.toggle_view = SwitchButton(left_text="Release", right_text="Current"); self.toggle_view.toggled.connect(self.update_launch_table)
        self.btn_copy_table = QPushButton("Copy"); self.btn_copy_table.setFixedSize(60, 30); self.btn_copy_table.setFont(QFont("나눔스퀘어 네오 ExtraBold", 9)); self.btn_copy_table.setCursor(Qt.PointingHandCursor); self.btn_copy_table.clicked.connect(self.launch_table.copy_current_data)
        lh_layout.addWidget(self.toggle_view); lh_layout.addSpacing(10); lh_layout.addWidget(self.btn_copy_table)
        c_launch = create_card("Model Launch Table", self.launch_table, extra_widget=launch_header_widget)
        list_container = QWidget(); v_list = QVBoxLayout(list_container); v_list.setContentsMargins(0,0,0,0); self.model_list = QListWidget(); self.model_list.setFont(QFont("나눔스퀘어 네오 Light", 9)); self.model_list.itemChanged.connect(self.on_item_changed); btn_box = QHBoxLayout(); self.btn_check_all = QPushButton("Check All"); self.btn_check_all.clicked.connect(self.check_all); self.btn_clear = QPushButton("Clear"); self.btn_clear.clicked.connect(self.clear_checks); btn_box.addWidget(self.btn_check_all); btn_box.addWidget(self.btn_clear); v_list.addWidget(self.model_list); v_list.addLayout(btn_box); c_list = create_card("Model List", list_container)
        self.heatmap.cell_clicked.connect(self.handle_heatmap_click); dashboard_layout = QHBoxLayout(); dashboard_layout.addWidget(c_heat, 3); dashboard_layout.addWidget(c_launch, 4); dashboard_layout.addWidget(c_list, 1)    
        main_layout = QVBoxLayout(self); main_layout.setContentsMargins(20,20,20,20); main_layout.addLayout(header_layout); main_layout.addLayout(input_layout); main_layout.addLayout(dashboard_layout); self.apply_theme(config.THEMES["Omdia"])
    def load_cache(self):
        p = self.settings.value("omdia_path", "")
        if p: self.path=p; self.drop_file.update_label(p); self.exec()
    def set_path(self, p): self.path = p; self.drop_file.update_label(p); self.exec()
    def anim(self): self.run.setText("Loading"+"."*(self.step%4)); self.step+=1
    def exec(self):
        if not self.path: return
        self.run.setEnabled(False); self.step=0; self.timer.start(500)
        self.th = OmdiaThread(self.path); self.th.result.connect(self.show_result); self.th.error.connect(self.err); self.th.start()
    def err(self, e): self.timer.stop(); self.run.setText("Load Data"); self.run.setEnabled(True); QMessageBox.critical(self, "Error", e)
    def show_result(self, df):
        self.timer.stop(); self.run.setText("Load Data"); self.run.setEnabled(True)
        if df is None or df.empty: QMessageBox.warning(self, "Warning", "No data found."); return
        self.full_df = df; self.all_years = sorted(df['Year'].unique().astype(str), reverse=True); self.selected_years = self.all_years[:2] 
        self.update_year_menu(); self.update_views()
    def update_year_menu(self):
        self.btn_year_menu.clear()
        for year in self.all_years: action = QAction(year, self); action.setCheckable(True); action.setChecked(year in self.selected_years); action.triggered.connect(self.on_year_toggled); self.btn_year_menu.addAction(action)
    def on_year_toggled(self):
        selected = []
        for action in self.btn_year_menu.actions(): 
            if action.isChecked(): selected.append(action.text())
        self.selected_years = selected; self.update_views()
    def update_views(self):
        if self.full_df is None: return
        cat = "Smartphone" if self.toggle_cat.isChecked() else "Foldable"
        self.heatmap.update_data_omdia(self.full_df, cat, self.selected_years)
        self.launch_table.update_table(None, None, None); self.model_list.clear()
    def handle_heatmap_click(self, brand, _): 
        if brand is None or brand == "Total": self.launch_table.update_table(None, None, None); self.model_list.clear(); return
        self.current_brand = brand; self.populate_model_list(); self.update_launch_table()
    def populate_model_list(self):
        self.model_list.clear()
        if self.full_df is None or not self.current_brand: return
        cat = "Smartphone" if self.toggle_cat.isChecked() else "Foldable"
        models = sorted(self.full_df[(self.full_df['Brand'] == self.current_brand) & (self.full_df['Category'] == cat)]['Model'].unique())
        for m in models: item = QListWidgetItem(m); item.setFlags(item.flags() | Qt.ItemIsUserCheckable); item.setCheckState(Qt.Checked); self.model_list.addItem(item)
    def on_item_changed(self, item): self.update_launch_table()
    def check_all(self):
        for i in range(self.model_list.count()): self.model_list.item(i).setCheckState(Qt.Checked)
        self.update_launch_table()
    def clear_checks(self):
        for i in range(self.model_list.count()): self.model_list.item(i).setCheckState(Qt.Unchecked)
        self.update_launch_table()
    def update_launch_table(self):
        if not hasattr(self, 'current_brand') or not self.current_brand: return
        cat = "Smartphone" if self.toggle_cat.isChecked() else "Foldable"
        visible_models = []
        for i in range(self.model_list.count()): 
            if self.model_list.item(i).checkState() == Qt.Checked: visible_models.append(self.model_list.item(i).text())
        mode = "Current" if self.toggle_view.isChecked() else "Release"
        self.launch_table.update_table(self.full_df, self.current_brand, cat, visible_models, mode=mode, target_years=self.selected_years)
  
