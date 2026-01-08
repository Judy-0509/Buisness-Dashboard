import sys
import traceback
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QHBoxLayout, QStackedWidget, QMessageBox
from PyQt5.QtGui import QFont
import config

# [Import UI Components]
# ui_components.py 안에 SellInPage가 포함되어 있어야 합니다.
try:
    from ui_components import (Sidebar, WeeklyPage, MonthlyPage, FlagshipPage, 
                               RegionBrandPage, OmdiaPage, SellInPage) # SellInPage 추가 확인
except ImportError as e:
    print(f"[Critical Error] ui_components.py에서 페이지 클래스를 불러올 수 없습니다: {e}")
    sys.exit(1)

# --- Global Exception Hook ---
def exception_hook(exctype, value, tb):
    error_msg = "".join(traceback.format_exception(exctype, value, tb))
    print(error_msg)
    try:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText("An unexpected error occurred.")
        msg.setInformativeText(str(value))
        msg.setDetailedText(error_msg)
        msg.setWindowTitle("Critical Error")
        msg.exec_()
    except: pass
    sys.exit(1)

sys.excepthook = exception_hook

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(2200, 1000)
        self.setWindowTitle("Market Intelligence Dashboard")
        self.setStyleSheet(f"background:{config.BG_MAIN};")
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 1. Sidebar
        self.sidebar = Sidebar()
        self.sidebar.page_changed.connect(self.switch_page)
        if hasattr(self.sidebar, 'theme_changed'):
            self.sidebar.theme_changed.connect(self.apply_theme)
        
        main_layout.addWidget(self.sidebar)

        # 2. Main Content Stack
        self.stack = QStackedWidget()
        
        # [중요] Sidebar 메뉴 순서와 정확히 일치해야 합니다.
        # 0: Weekly
        # 1: Monthly
        # 2: Flagship
        # 3: Region Brand
        # 4: Sell in Sell Thru (NEW)
        # 5: Omdia
        
        self.stack.addWidget(WeeklyPage())      # Index 0
        self.stack.addWidget(MonthlyPage())     # Index 1
        self.stack.addWidget(FlagshipPage())    # Index 2
        self.stack.addWidget(RegionBrandPage()) # Index 3
        self.stack.addWidget(SellInPage())      # Index 4  <-- 여기가 새로 추가된 부분입니다!
        self.stack.addWidget(OmdiaPage())       # Index 5
        
        main_layout.addWidget(self.stack)

    def switch_page(self, index):
        self.stack.setCurrentIndex(index)

    def apply_theme(self, theme):
        for i in range(self.stack.count()):
            page = self.stack.widget(i)
            if hasattr(page, 'apply_theme'):
                page.apply_theme(theme)

if __name__=="__main__":
    app = QApplication(sys.argv)
    font = QFont("나눔스퀘어 네오 Light", 10)
    font.setStyleStrategy(QFont.PreferAntialias)
    app.setFont(font)
    
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
