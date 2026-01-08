import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.colors import LinearSegmentedColormap
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QMessageBox, QTableWidget, QTableWidgetItem, 
                             QAbstractItemView, QHBoxLayout, QLabel, QComboBox, QPushButton, 
                             QListWidget, QListWidgetItem, QFrame, QRadioButton, QButtonGroup, QHeaderView)
from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtGui import QFont, QColor
import config

# --- Base Chart Widget ---
class BaseChartWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.fig, self.ax = plt.subplots(1, 1, figsize=(10, 5))
        self.fig.patch.set_facecolor('none')
        self.ax.set_facecolor('none')
        self.canvas = FigureCanvas(self.fig)
        self.canvas.setStyleSheet("background:transparent;")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.addWidget(self.canvas)
        self.annot = None; self.highlight_dot = None; self.lines_dict = {}
        self.canvas.mpl_connect('button_press_event', self.on_click)
        self.canvas.mpl_connect('motion_notify_event', self.on_hover)
    def clear_plot(self, message="Ready to Analyze"):
        self.ax.clear(); self.ax.set_xticks([]); self.ax.set_yticks([])
        self.ax.text(0.5, 0.5, message, ha='center', va='center')
        self.annot = None; self.canvas.draw()
    def on_click(self, event): pass 
    def on_hover(self, event): pass 
    @staticmethod
    def group_brand_static(name):
        name_str = str(name).strip(); u_name = name_str.upper()
        if u_name in ["OPPO", "ONEPLUS", "REALME"]: return "Oppo"
        target_map = {"APPLE": "Apple", "GOOGLE": "Google", "HONOR": "Honor", "HUAWEI": "Huawei", "SAMSUNG": "Samsung", "XIAOMI": "Xiaomi", "VIVO": "vivo"}
        if u_name in target_map: return target_map[u_name]
        return "Others"
    def group_brand(self, name): return self.group_brand_static(name)

# --- Heatmap Widget ---
class HeatmapWidget(BaseChartWidget):
    cell_clicked = pyqtSignal(object, object)
    def __init__(self, time_col="Week"):
        super().__init__(); self.time_col = time_col
        self.fig.subplots_adjust(left=0.15, right=0.95, top=0.9, bottom=0.15)
        self.p24 = None; self.p25 = None; self.current_mode = "diff"; self.full_df = None; self.selected_idx = None
        self.clear_plot()
    def safe_remove_cbar(self):
        if hasattr(self, 'cbar') and self.cbar:
            try: self.cbar.remove()
            except: pass
            self.cbar = None
    def clear_plot(self, msg="Ready to Analyze"): self.safe_remove_cbar(); super().clear_plot(msg)
    def set_mode(self, mode): 
        self.current_mode = mode 
        if hasattr(self, 'p25') and self.p25 is not None: self.refresh_view()
    def copy_data(self):
        if self.p25 is None: QMessageBox.warning(self, "Warning", "No data to copy."); return
        export_df = pd.DataFrame()
        if self.current_mode == "pct" and self.p24 is not None:
             safe_p24 = self.p24.replace(0, np.nan); export_df = (self.p25 - self.p24) / safe_p24 * 100
        elif self.current_mode == "diff" and self.p24 is not None: export_df = (self.p25 - self.p24) / 1000000.0
        else: export_df = self.p25 / 1000000.0 
        if not export_df.empty: export_df.to_clipboard(); QMessageBox.information(self, "Info", f"Copied!")
        else: QMessageBox.warning(self, "Warning", "Data empty.")
    def reset_state(self): self.selected_idx = None; self.refresh_view(); self.cell_clicked.emit(None, None)
    def on_click(self, event):
        if self.p25 is None or event.inaxes != self.ax: self.reset_state(); return
        col_idx = int(round(event.xdata)); row_idx = int(round(event.ydata))
        if 0 <= row_idx < len(self.p25.index) and 0 <= col_idx < len(self.p25.columns):
            self.selected_idx = (row_idx, col_idx); self.refresh_view(); self.cell_clicked.emit(self.p25.index[row_idx], self.p25.columns[col_idx])
        else: self.reset_state()
    def on_hover(self, event):
        if event.inaxes != self.ax or self.p25 is None: 
            if self.annot and self.annot.get_visible(): self.annot.set_visible(False); self.canvas.draw_idle()
            return
        col_idx = int(round(event.xdata)); row_idx = int(round(event.ydata))
        if 0 <= row_idx < len(self.p25.index) and 0 <= col_idx < len(self.p25.columns):
            if self.annot:
                if self.p24 is not None and self.current_mode == "pct":
                    val_24 = self.p24.iloc[row_idx, col_idx] / 1000000.0; val_25 = self.p25.iloc[row_idx, col_idx] / 1000000.0
                    self.annot.set_text(f"Old: {val_24:.2f} Mu\nNew: {val_25:.2f} Mu")
                else:
                    val = self.p25.iloc[row_idx, col_idx]
                    try: 
                        fval = float(val)
                        # 화면 표시용 (/1M)
                        self.annot.set_text(f"{fval/1000000:.2f} Mu")
                    except: self.annot.set_text(str(val))
                self.annot.xy = (col_idx, row_idx); self.annot.set_visible(True); self.canvas.draw_idle()
        else:
            if self.annot and self.annot.get_visible(): self.annot.set_visible(False); self.canvas.draw_idle()
    def update_data(self, raw_data): # Weekly
        all_dfs = []
        for sheet_name, df in raw_data.items():
            temp = df.copy()
            if "Region" not in temp.columns: temp["Region"] = config.WEEKLY_MAP.get(sheet_name, sheet_name)
            temp["Sales"] = pd.to_numeric(temp["Sales"], errors='coerce').fillna(0)
            if self.time_col in temp.columns: temp[self.time_col] = pd.to_numeric(temp[self.time_col], errors='coerce')
            temp["Brand_Group"] = temp["Brand"].apply(self.group_brand); all_dfs.append(temp)
        self.full_df = pd.concat(all_dfs); exclude = ["East Europe", "E.Europe", "E. Europe", "East Europe "]; self.full_df = self.full_df[~self.full_df['Region'].isin(exclude)]
        if 2025 in self.full_df['Year'].unique(): max_time = self.full_df[self.full_df['Year'] == 2025][self.time_col].max()
        else: max_time = 52 if self.time_col == "Week" else 12
        df_ytd = self.full_df[self.full_df[self.time_col] <= max_time]
        self.p24 = df_ytd[df_ytd['Year'] == 2024].pivot_table(index="Brand_Group", columns="Region", values="Sales", aggfunc="sum", fill_value=0)
        self.p25 = df_ytd[df_ytd['Year'] == 2025].pivot_table(index="Brand_Group", columns="Region", values="Sales", aggfunc="sum", fill_value=0)
        self._process_others_and_total(); self.selected_idx = None; self.refresh_view()
    def update_data_flagship(self, df, category, target_years=None):
        filtered_df = df[df['Category'] == category].copy()
        if not filtered_df.empty:
            max_date = filtered_df['Date'].max(); max_month = max_date.month
            filtered_df = filtered_df[filtered_df['Date'].dt.month <= max_month]
        filtered_df['YearStr'] = filtered_df['Date'].dt.year.astype(str)
        if target_years: filtered_df = filtered_df[filtered_df['YearStr'].isin(target_years)]
        self.p25 = filtered_df.pivot_table(index="Brand", columns="YearStr", values="Sales", aggfunc="sum", fill_value=0)
        self.p24 = None
        if self.p25.empty: self.clear_plot("No Data"); return
        self.p25.loc['Total'] = self.p25.sum(axis=0); last_col = self.p25.columns[-1]; self.p25 = self.p25.sort_values(by=last_col, ascending=False)
        if 'Total' in self.p25.index: self.p25 = pd.concat([self.p25.loc[['Total']], self.p25.drop('Total')])
        self.selected_idx = None; self.refresh_view()
    def update_data_omdia(self, df, category, target_years=None):
        filtered_df = df[df['Category'] == category].copy()
        filtered_df['TimeLabel'] = filtered_df['Year'].astype(str) + " " + filtered_df['Quarter'].astype(str) + "Q"
        if target_years: filtered_df = filtered_df[filtered_df['Year'].astype(str).isin(target_years)]
        self.p25 = filtered_df.pivot_table(index="Brand", columns="TimeLabel", values="Sales", aggfunc="sum", fill_value=0)
        self.p24 = None
        if self.p25.empty: self.clear_plot("No Data"); return
        cols = sorted(self.p25.columns, key=lambda x: (int(x.split()[0]), int(x.split()[1][0])))
        self.p25 = self.p25[cols]; self.p25.loc['Total'] = self.p25.sum(axis=0)
        last_col = self.p25.columns[-1]; self.p25 = self.p25.sort_values(by=last_col, ascending=False)
        if 'Total' in self.p25.index: self.p25 = pd.concat([self.p25.loc[['Total']], self.p25.drop('Total')])
        self.selected_idx = None; self.refresh_view()
    def update_data_ti_ytd(self, df, measure_filter, target_years):
        target_df = df.copy()
        if measure_filter: target_df = target_df[target_df['Measure'] == measure_filter]
        if target_years: target_df = target_df[target_df['Year'].astype(str).isin(target_years)]
        if target_df.empty: self.clear_plot("No Data"); return
        max_year = target_df['Year'].max()
        max_month = target_df[target_df['Year'] == max_year]['Month'].max()
        target_df = target_df[target_df['Month'] <= max_month]
        self.p25 = target_df.pivot_table(index="Brand", columns="Year", values="Sales", aggfunc="sum", fill_value=0)
        self.ti_vol = self.p25.copy(); self.ti_diff = pd.DataFrame(index=self.p25.index); self.ti_yoy = pd.DataFrame(index=self.p25.index)
        cols = sorted(self.p25.columns)
        for i, col in enumerate(cols):
            if i == 0: self.ti_diff[col] = 0; self.ti_yoy[col] = 0
            else:
                prev_col = cols[i-1]; self.ti_diff[col] = self.ti_vol[col] - self.ti_vol[prev_col]
                prev_val = self.ti_vol[prev_col].replace(0, np.nan); self.ti_yoy[col] = (self.ti_vol[col] - self.ti_vol[prev_col]) / prev_val * 100
                self.ti_yoy[col] = self.ti_yoy[col].fillna(0)
        self.ti_vol.loc['Total'] = self.ti_vol.sum(axis=0); self.ti_diff.loc['Total'] = self.ti_diff.sum(axis=0)
        for i, col in enumerate(cols):
            if i > 0:
                prev = self.ti_vol.loc['Total', cols[i-1]]; curr = self.ti_vol.loc['Total', col]
                val = (curr - prev) / prev * 100 if prev != 0 else 0
                self.ti_yoy.loc['Total', col] = val
            else: self.ti_yoy.loc['Total', col] = 0
        last_col = cols[-1]; sorted_idx = self.ti_vol.sort_values(by=last_col, ascending=False).index
        if 'Total' in sorted_idx: sorted_idx = ['Total'] + [x for x in sorted_idx if x != 'Total']
        self.ti_vol = self.ti_vol.reindex(sorted_idx); self.ti_diff = self.ti_diff.reindex(sorted_idx); self.ti_yoy = self.ti_yoy.reindex(sorted_idx)
        self.p25 = self.ti_vol; self.selected_idx = None; self.refresh_view_ti() 
    def refresh_view_ti(self):
        self.safe_remove_cbar(); self.fig.clear(); self.ax = self.fig.add_subplot(111); self.ax.set_facecolor('none')
        self.fig.subplots_adjust(left=0.15, right=0.95, top=0.85, bottom=0.05)
        if self.current_mode == "pct": data = self.ti_yoy; fmt_type = "pct"; vmin, vmax = -50, 50
        elif self.current_mode == "diff": data = self.ti_diff; fmt_type = "diff"; mx = data.abs().max().max(); vmin, vmax = -mx, mx
        else: data = self.ti_vol; fmt_type = "vol"; vmin, vmax = 0, data.max().max()
        self.draw_heatmap(data, self.ti_vol, vmin, vmax, fmt_type); self.canvas.draw()
    def _process_others_and_total(self):
        if "Others" not in self.p24.index: self.p24.loc["Others"] = 0
        if "Others" not in self.p25.index: self.p25.loc["Others"] = 0
        target_brands = [b for b in self.p24.index if b != "Others"]
        for region in self.p24.columns:
            low_vol = self.p24.loc[target_brands, region]
            move = low_vol[low_vol < 1000000].index.tolist()
            if move:
                self.p24.loc["Others", region] += self.p24.loc[move, region].sum(); self.p24.loc[move, region] = 0
                self.p25.loc["Others", region] += self.p25.loc[move, region].sum(); self.p25.loc[move, region] = 0
        self.p24['Total'] = self.p24.sum(axis=1); self.p25['Total'] = self.p25.sum(axis=1)
        self.p24.loc['Total'] = self.p24.sum(axis=0); self.p25.loc['Total'] = self.p25.sum(axis=0)
        all_brands = sorted(list(set(self.p24.index) | set(self.p25.index)))
        if "Total" in all_brands: all_brands.remove("Total")
        if "Others" in all_brands: all_brands.remove("Others")
        final_idx = ["Total"] + all_brands + ["Others"]
        all_regions = sorted(list(set(self.p24.columns) | set(self.p25.columns)))
        if "Total" in all_regions: all_regions.remove("Total")
        final_cols = ["Total"] + all_regions
        self.p24 = self.p24.reindex(index=[x for x in final_idx if x in self.p24.index], columns=final_cols, fill_value=0)
        self.p25 = self.p25.reindex(index=[x for x in final_idx if x in self.p25.index], columns=final_cols, fill_value=0)
    def refresh_view(self):
        if hasattr(self, 'ti_vol') and self.ti_vol is not None: self.refresh_view_ti(); return
        self.safe_remove_cbar(); self.fig.clear(); self.ax = self.fig.add_subplot(111); self.ax.set_facecolor('none')
        self.fig.subplots_adjust(left=0.15, right=0.95, top=0.85, bottom=0.05)
        self.annot = self.ax.annotate("", xy=(0,0), xytext=(10,10), textcoords="offset points", bbox=dict(boxstyle="round", fc="w", alpha=0.9), arrowprops=dict(arrowstyle="->")); self.annot.set_visible(False)
        data = None; fmt_type = "vol"; vmin, vmax = 0, 1
        if self.p24 is None: data = self.p25; fmt_type = "vol"; vmin, vmax = 0, data.max().max()
        else:
            if self.current_mode == "pct":
                safe_p24 = self.p24.replace(0, np.nan); data = (self.p25 - self.p24) / safe_p24 * 100
                mask_zero = (self.p24 == 0); data[mask_zero] = np.where(self.p25[mask_zero] > 0, 100.0, 0.0); fmt_type = "pct"; vmin, vmax = -50, 50
            elif self.current_mode == "diff": data = self.p25 - self.p24; fmt_type = "diff"; max_val = data.abs().max().max(); vmin, vmax = -max_val, max_val if max_val > 0 else 1
            elif self.current_mode == "raw": data = self.p25; fmt_type = "vol"; vmin, vmax = 0, data.max().max()
        self.draw_heatmap(data, self.p25, vmin, vmax, fmt_type); self.canvas.draw()
    def draw_heatmap(self, data_df, vol_df, vmin, vmax, fmt_type):
        colors = ["#ffffff", "#2563EB"] if fmt_type == "vol" else ["#f44336", "#ffffff", "#90caf9"]
        custom_cmap = LinearSegmentedColormap.from_list("custom_cmap", colors); norm = plt.Normalize(vmin, vmax)
        mapped_data = custom_cmap(norm(data_df.values))
        if self.selected_idx: sel_r, sel_c = self.selected_idx; mapped_data[..., 3] = 0.3; mapped_data[sel_r, sel_c, 3] = 1.0
        im = self.ax.imshow(mapped_data, aspect='auto'); self.ax.xaxis.tick_top()
        self.ax.set_xticks(range(len(data_df.columns))); self.ax.set_yticks(range(len(data_df.index)))
        xt = self.ax.set_xticklabels(data_df.columns, fontsize=11, rotation=45, ha='left')
        yt = self.ax.set_yticklabels(data_df.index, fontsize=12)
        for label in xt + yt:
            if label.get_text() == "Total": label.set_fontweight('bold')
        sm = plt.cm.ScalarMappable(cmap=custom_cmap, norm=norm); sm.set_array([]); self.cbar = self.fig.colorbar(sm, ax=self.ax, fraction=0.046, pad=0.04)
        cbar_label = 'Volume (Mu)' if fmt_type == "vol" else ('Growth Rate (%)' if fmt_type == "pct" else 'Volume Diff (Mu)')
        self.cbar.set_label(cbar_label, rotation=270, labelpad=15)
        for i in range(len(data_df.index)):
            for j in range(len(data_df.columns)):
                val = data_df.iloc[i, j]; vol = vol_df.iloc[i, j]
                is_dark = abs(val) > 40 if fmt_type == "pct" else (abs(val) > (vmax * 0.6) if fmt_type == "diff" else val > (vmax * 0.5))
                text_color = "white" if is_dark else "black"
                text_alpha = 1.0 if not self.selected_idx or (i, j) == self.selected_idx else 0.3
                if fmt_type == "pct":
                    if vol == 0 and val == 0: txt = "-"
                    elif val == 100.0 and vol > 0 and data_df.iloc[i,j] == 100.0: txt = "New"
                    else: txt = f"{val:+.1f}%"
                else:
                    # [MODIFIED] Divide by 1M for Mu display
                    display_val = val / 1000000.0
                    txt = f"{display_val:,.2f}"
                self.ax.text(j, i, txt, ha="center", va="center", color=text_color, fontsize=11, alpha=text_alpha)
  class ComparisonTableWidget(QWidget):
    cellClicked = pyqtSignal(str, str) # model, quarter
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        self.table = QTableWidget()
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.cellClicked.connect(self.on_cell_clicked)
        self.table.setStyleSheet("""
            QTableWidget { background-color: transparent; gridline-color: #d0d0d0; font-family: 'Malgun Gothic'; font-size: 10pt; }
            QHeaderView::section { background-color: #f0f0f0; padding: 4px; border: 1px solid #d0d0d0; font-weight: bold; }
        """)
        layout.addWidget(self.table)
        self.df = None
        self.full_data = None

    def update_data(self, df):
        self.full_data = df
        if df.empty:
            self.table.clear(); self.table.setRowCount(0); self.table.setColumnCount(0); return
        
        # Calculate Average
        pivot = df.pivot_table(index='Model', columns='Date', values='Value', aggfunc='mean', fill_value=0)
        
        # Sort Columns
        try:
            cols = sorted(pivot.columns, key=lambda x: (int(x.split()[0]), int(x.split()[1][1]))) # YYYY Qn -> split by space? "2023 Q1" -> 2023, 1
            pivot = pivot[cols]
        except: pass
        
        self.df = pivot
        
        self.table.clear()
        self.table.setRowCount(len(pivot.index))
        self.table.setColumnCount(len(pivot.columns))
        
        self.table.setVerticalHeaderLabels(pivot.index.astype(str))
        self.table.setHorizontalHeaderLabels(pivot.columns.astype(str))
        
        for i in range(len(pivot.index)):
            for j in range(len(pivot.columns)):
                val = pivot.iloc[i, j]
                txt = f"{val:,.2f}" if val != 0 else "-"
                item = QTableWidgetItem(txt)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(i, j, item)
        self.table.resizeColumnsToContents()

    def on_cell_clicked(self, row, col):
        if self.df is None: return
        model = self.df.index[row]
        quarter = self.df.columns[col]
        self.cellClicked.emit(model, quarter)

    def copy_data(self):
        if self.df is not None:
            self.df.to_clipboard()
            QMessageBox.information(self, "Info", "Copied Average Data")

# [NEW] Detail Chart Widget (Sidebar)
class DetailChartWidget(BaseChartWidget):
    def __init__(self):
        super().__init__()
        self.fig.subplots_adjust(left=0.15, right=0.9, top=0.9, bottom=0.2)
        self.clear_plot()

    def clear_plot(self, message="Select a cell"):
        super().clear_plot(message)

    def update_chart(self, full_df, model, quarter):
        self.ax.clear()
        target = full_df[(full_df['Model'] == model) & (full_df['Date'] == quarter)]
        
        if target.empty:
            self.ax.text(0.5, 0.5, "No Detail Data", ha='center', va='center')
            self.canvas.draw()
            return
            
        # Bar Chart
        firms = target['Firm'].tolist()
        values = target['Value'].tolist()
        colors = [config.THEMES['Omdia']['dark'], config.THEMES['TechInsights']['dark'], '#F39C12'] # Colors for Omdia, TI, GfK
        
        bars = self.ax.bar(firms, values, color=colors[:len(firms)], width=0.5)
        
        self.ax.set_title(f"{model} - {quarter}", fontsize=11, fontweight='bold')
        self.ax.set_ylabel("Volume (Mu)")
        self.ax.grid(axis='y', linestyle='--', alpha=0.5)
        
        # Annotate
        for bar in bars:
            height = bar.get_height()
            self.ax.text(bar.get_x() + bar.get_width()/2., height,
                         f'{height:.2f}', ha='center', va='bottom')
            
        self.canvas.draw()

class LaunchTableWidget(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self); layout.setContentsMargins(0, 0, 0, 0)
        self.table = QTableWidget(); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setStyleSheet("QTableWidget { background-color: transparent; gridline-color: #d0d0d0; font-family: 'Malgun Gothic'; font-size: 10pt; } QHeaderView::section { background-color: #f0f0f0; padding: 4px; border: 1px solid #d0d0d0; font-weight: bold; }")
        layout.addWidget(self.table); self.current_df = None
    def copy_current_data(self):
        if self.current_df is not None and not self.current_df.empty: (self.current_df / 1000000).to_clipboard(); QMessageBox.information(self, "Info", "Copied!")
        else: QMessageBox.warning(self, "Warning", "No data.")
    def update_table(self, df, brand, category, models=None, mode="Release", target_years=None):
        if df is None or brand is None: self.table.clear(); return
        target = df[(df['Brand']==brand) & (df['Category']==category)]
        if models: target = target[target['Model'].isin(models)]
        if target.empty: self.table.clear(); self.table.setRowCount(0); self.table.setColumnCount(0); return
        if mode == "Release":
            pivot = target.pivot_table(index="Model", columns="QuartersSinceLaunch", values="Sales", aggfunc="sum")
            cols = sorted([c for c in pivot.columns if c >= 0]); pivot = pivot[cols]; pivot.columns = [f"Q+{int(c)}" for c in pivot.columns]
        else:
            if target_years: target = target[target['Year'].astype(str).isin(target_years)]
            target = target.copy(); target['YQ'] = target['Year'].astype(str) + " " + target['Quarter'].astype(str) + "Q"
            pivot = target.pivot_table(index="Model", columns="YQ", values="Sales", aggfunc="sum")
            sorted_cols = sorted(pivot.columns, key=lambda x: (int(x.split()[0]), int(x.split()[1][0]))); pivot = pivot[sorted_cols]
        self.current_df = pivot; self.table.clear(); self.table.setRowCount(len(pivot.index)); self.table.setColumnCount(len(pivot.columns))
        self.table.setVerticalHeaderLabels(pivot.index.astype(str)); self.table.setHorizontalHeaderLabels(pivot.columns.astype(str))
        for i in range(len(pivot.index)):
            for j in range(len(pivot.columns)):
                val = pivot.iloc[i, j]; txt = f"{val/1000000:.2f}" if pd.notna(val) and val != 0 else "-"
                item = QTableWidgetItem(txt); item.setTextAlignment(Qt.AlignCenter); self.table.setItem(i, j, item)
        self.table.resizeColumnsToContents()

class LaunchTrendWidget(BaseChartWidget):
    def __init__(self):
        super().__init__(); self.fig.subplots_adjust(right=0.75, left=0.08, top=0.9, bottom=0.15); self.full_df = None; self.current_brand = None; self.current_category = None; self.current_pivot = None; self.clear_plot()
    def clear_plot(self): super().clear_plot("Select Brand")
    def copy_current_data(self):
        if self.current_pivot is not None and not self.current_pivot.empty: df_mu = self.current_pivot / 1000000.0; df_mu.to_clipboard(); QMessageBox.information(self, "Info", "Copied!")
        else: QMessageBox.warning(self, "Warning", "No data.")
    def update_chart(self, full_df, brand, category, visible_models=None, x_limit=None, is_cumulative=False, time_unit="Month"):
        self.full_df = full_df; self.current_brand = brand; self.current_category = category
        if brand is None or brand == "Total": self.clear_plot(); return
        self.ax.clear(); self.lines_dict = {}
        self.annot = self.ax.annotate("", xy=(0,0), xytext=(15,15), textcoords="offset points", bbox=dict(boxstyle="round4,pad=0.5", fc=config.COLOR_23, ec="none", alpha=0.9), arrowprops=dict(arrowstyle="->", color=config.COLOR_23)); self.annot.set_visible(False)
        self.highlight_dot, = self.ax.plot([], [], 'o', markersize=8, color='white', markeredgecolor='black', visible=False)
        target_df = full_df[(full_df['Brand'] == brand) & (full_df['Category'] == category)]
        if visible_models is not None: target_df = target_df[target_df['Model'].isin(visible_models)]
        idx_col = 'QuartersSinceLaunch' if time_unit == "Quarter" else 'MonthsSinceLaunch'
        if idx_col not in target_df.columns: self.ax.text(0.5, 0.5, "Time column missing", ha='center', va='center'); self.canvas.draw(); return
        pivot = target_df.pivot_table(index=idx_col, columns='Model', values='Sales', aggfunc='sum')
        if pivot.empty: self.ax.text(0.5, 0.5, "No Data / Unchecked All", ha='center', va='center'); self.current_pivot = None; self.canvas.draw(); return
        if x_limit is not None: pivot = pivot[pivot.index <= x_limit]
        if is_cumulative: pivot = pivot.cumsum()
        self.current_pivot = pivot; max_idx = pivot.index.max(); max_idx = 0 if pd.isna(max_idx) else max_idx
        new_index = range(int(max_idx) + 1); pivot = pivot.reindex(new_index)
        all_models_in_cat = full_df[(full_df['Brand'] == brand) & (full_df['Category'] == category)]['Model'].unique(); all_models_in_cat.sort()
        colors = config.generate_gradient_colors(len(all_models_in_cat)); color_map = {m: c for m, c in zip(all_models_in_cat, colors)}
        for model in pivot.columns:
            valid_data = pivot[model].dropna(); color = color_map.get(model, 'black')
            line, = self.ax.plot(valid_data.index, valid_data.values / 1000000.0, label=model, color=color, linewidth=2.5, marker='o', markersize=6); self.lines_dict[line] = model
        self.ax.set_title("Model Launch Trend", fontsize=12, fontweight='bold', pad=10)
        xlabel = "Quarters Since Launch (Q+N)" if time_unit == "Quarter" else "Months Since Launch (T+N)"
        self.ax.set_xlabel(xlabel, fontsize=10); self.ax.set_ylabel("Sales Volume (Mn Units)", fontsize=10); self.ax.legend(frameon=False, bbox_to_anchor=(1.02, 1), loc='upper left')
        self.ax.grid(True, linestyle='--', alpha=0.5); self.fig.subplots_adjust(right=0.75, left=0.08, top=0.9, bottom=0.15); self.canvas.draw()
    def on_hover(self, event):
        if event.inaxes != self.ax or not self.annot: return
        found = False
        for line, model_name in self.lines_dict.items():
            cont, ind = line.contains(event)
            if cont:
                x_data, y_data = line.get_data(); idx = ind["ind"][0]; pos_x = x_data[idx]; pos_y = y_data[idx]
                self.annot.xy = (pos_x, pos_y); text = f"{model_name}\n+{int(pos_x)}\n{pos_y:.2f} Mu"
                self.annot.set_text(text); self.annot.set_visible(True); self.highlight_dot.set_data([pos_x], [pos_y]); self.highlight_dot.set_color(line.get_color()); self.highlight_dot.set_visible(True); self.canvas.draw_idle(); found = True; break 
        if not found and self.annot.get_visible(): self.annot.set_visible(False); self.highlight_dot.set_visible(False); self.canvas.draw_idle()

class LineChartWidget(BaseChartWidget):
    def __init__(self, time_col="Week"): super().__init__(); self.time_col = time_col; self.current_data = None; self.clear_plot()
    def clear_plot(self): super().clear_plot("Select a cell in Heatmap")
    def copy_current_data(self):
        if self.current_data is not None and not self.current_data.empty: df_mu = self.current_data / 1000000.0; df_mu.to_clipboard(); QMessageBox.information(self, "Info", "Copied!")
        else: QMessageBox.warning(self, "Warning", "No data.")
    def update_chart(self, full_df, brand, region, pivot_24, is_cumulative=False):
        self.ax.clear(); target_df = full_df.copy()
        if region != "Total": target_df = target_df[target_df['Region'] == region]
        if brand == "Total": pass 
        elif brand == "Others":
            if "Brand_Group" not in target_df.columns: target_df["Brand_Group"] = target_df["Brand"].apply(lambda x: HeatmapWidget.group_brand_static(x))
            grp_sums = target_df[target_df['Year'] == 2024].groupby("Brand_Group")['Sales'].sum()
            others_candidates = grp_sums[grp_sums < 1000000].index.tolist(); others_candidates.append("Others")
            target_df = target_df[target_df['Brand_Group'].isin(others_candidates)]
        else:
            if "Brand_Group" not in target_df.columns: target_df["Brand_Group"] = target_df["Brand"].apply(lambda x: HeatmapWidget.group_brand_static(x))
            target_df = target_df[target_df['Brand_Group'] == brand]
        target_df = target_df[target_df['Year'].isin([2023, 2024, 2025])]
        if self.time_col == "Week" and "Week" in target_df.columns: target_df.loc[target_df["Week"] == 53, "Week"] = 52
        weekly_trend = target_df.pivot_table(index=self.time_col, columns="Year", values="Sales", aggfunc="sum")
        if is_cumulative: weekly_trend = weekly_trend.cumsum()
        self.current_data = weekly_trend; years = [2023, 2024, 2025]; colors = {2023: config.COLOR_23, 2024: config.COLOR_24, 2025: config.COLOR_25}; self.lines_dict = {}
        for y in years:
            if y in weekly_trend.columns:
                data = weekly_trend[y].dropna()
                if not data.empty: mu_values = data.values / 1000000.0; line, = self.ax.plot(data.index, mu_values, label=str(y), color=colors[y], linewidth=2.5); self.lines_dict[y] = line
        trend_type = "Weekly" if self.time_col == "Week" else "Monthly"; title_suffix = "(Cumulative)" if is_cumulative else f"({trend_type} Trend)"
        self.ax.set_title(f"{brand} in {region} {title_suffix}", fontsize=12, fontweight='bold', pad=10)
        self.ax.legend(frameon=False); self.ax.grid(True, linestyle='--', alpha=0.5); self.ax.set_ylabel("(Mu)", fontsize=10, rotation=0, labelpad=20, y=1.02); self.ax.tick_params(axis='both', labelsize=10)
        self.highlight_dot, = self.ax.plot([], [], 'o', markersize=8, color='white', markeredgecolor='black', visible=False)
        self.annot = self.ax.annotate("", xy=(0,0), xytext=(15,15), textcoords="offset points", bbox=dict(boxstyle="round4,pad=0.5", fc=config.COLOR_23, ec="none", alpha=0.9), arrowprops=dict(arrowstyle="->", color=config.COLOR_23)); self.annot.set_visible(False); self.canvas.draw()
    def on_hover(self, event):
        if event.inaxes != self.ax or not self.annot: return
        found = False
        for year, line in self.lines_dict.items():
            cont, ind = line.contains(event)
            if cont:
                x_data, y_data = line.get_data(); idx = ind["ind"][0]; pos_x = x_data[idx]; pos_y = y_data[idx]
                self.annot.xy = (pos_x, pos_y); time_prefix = "Week" if self.time_col == "Week" else "Month"
                text = f"{time_prefix} {int(pos_x)}\n{pos_y:.2f} Mu"; self.annot.set_text(text); self.annot.set_visible(True); self.highlight_dot.set_data([pos_x], [pos_y]); self.highlight_dot.set_color(line.get_color()); self.highlight_dot.set_visible(True); self.canvas.draw_idle(); found = True; break 
        if not found and self.annot.get_visible(): self.annot.set_visible(False); self.highlight_dot.set_visible(False); self.canvas.draw_idle()

class TrendWidget(BaseChartWidget):
    def __init__(self, time_col="Week"):
        super().__init__(); self.time_col = time_col; self.fig.subplots_adjust(left=0.15, right=0.95, top=0.75, bottom=0.15); self.full_df = None; self.current_brand = None; self.current_region = None; self.is_vol_mode = False; self.current_data = None; self.pivot_vol = None; self.clear_plot()
    def clear_plot(self): super().clear_plot("Select Total Row/Col")
    def set_mode(self, is_checked):
        self.is_vol_mode = is_checked; 
        if self.full_df is not None and self.current_brand and self.current_region: self.update_chart(self.full_df, self.current_brand, self.current_region)
    def copy_current_data(self):
        if self.pivot_vol is not None and not self.pivot_vol.empty:
            df_export = self.pivot_vol.T 
            df_export = df_export / 1000000.0
            df_export.to_clipboard()
            QMessageBox.information(self, "Info", "Copied! (Year as Columns, Mu Unit)")
        else: QMessageBox.warning(self, "Warning", "No data.")
    def update_chart(self, full_df, brand, region):
        self.full_df = full_df; self.current_brand = brand; self.current_region = region
        if brand != "Total" and region != "Total": self.clear_plot(); self.ax.text(0.5, 0.5, "Select Total Row/Col for Trend", ha='center', va='center'); self.canvas.draw(); return
        self.fig.clear(); self.ax = self.fig.add_subplot(111); self.ax.set_facecolor('none'); self.fig.subplots_adjust(left=0.15, right=0.95, top=0.75, bottom=0.15)
        if 2025 in full_df['Year'].unique(): max_time = full_df[full_df['Year'] == 2025][self.time_col].max()
        else: max_time = 52 if self.time_col == "Week" else 12
        df = full_df[full_df['Year'].isin([2023, 2024, 2025])].copy(); df = df[df[self.time_col] <= max_time] 
        if "Brand_Group" not in df.columns: df["Brand_Group"] = df["Brand"].apply(lambda x: HeatmapWidget.group_brand_static(x))
        category_col = ""; title_prefix = ""; time_label = "W" if self.time_col == "Week" else "M"
        if brand == "Total" and region == "Total": category_col = "Brand_Group"; title_prefix = f"Global Market Breakdown (YTD {time_label}{int(max_time)})"
        elif brand != "Total": df = df[df['Brand_Group'] == brand]; category_col = "Region"; title_prefix = f"{brand}'s Regional Split (YTD {time_label}{int(max_time)})"
        elif region != "Total": df = df[df['Region'] == region]; category_col = "Brand_Group"; title_prefix = f"{region}'s Market Breakdown (YTD {time_label}{int(max_time)})"
        pivot = df.pivot_table(index="Year", columns=category_col, values="Sales", aggfunc="sum", fill_value=0)
        col_sum = pivot.sum(axis=0).sort_values(ascending=False); pivot = pivot[col_sum.index]
        years = [2023, 2024, 2025]; pivot = pivot.reindex(years)
        self.pivot_vol = pivot 
        if not self.is_vol_mode: pivot_pct = pivot.div(pivot.sum(axis=1), axis=0) * 100; plot_data = pivot_pct; ylabel = "Share (%)"
        else: plot_data = pivot / 1000000.0; ylabel = "Volume (Mu)"
        self.current_data = plot_data; categories = plot_data.columns; colors = config.generate_gradient_colors(len(categories)); bottom = np.zeros(len(years))
        for i, cat in enumerate(categories):
            vals = plot_data[cat].fillna(0).values; self.ax.bar(years, vals, bottom=bottom, label=cat, color=colors[i], width=0.6, edgecolor='white', linewidth=0.5)
            for j, val in enumerate(vals):
                threshold = 3 if not self.is_vol_mode else 0.5
                if val >= threshold:
                    y_pos = bottom[j] + val / 2; x_pos = years[j]; txt = f"{int(round(val))}%" if not self.is_vol_mode else f"{val:.1f}"
                    txt_color = 'white' if i < len(categories) * 0.5 else 'black'; self.ax.text(x_pos, y_pos, txt, ha='center', va='center', color=txt_color, fontsize=9, fontweight='bold')
            bottom += vals
        self.ax.set_xticks(years); self.ax.set_title(title_prefix, fontsize=12, fontweight='bold', pad=10); self.ax.set_ylabel(ylabel, fontsize=10)
        self.ax.legend(loc='lower center', bbox_to_anchor=(0.5, 1.18), ncol=min(len(categories), 4), frameon=False, fontsize=9)
        self.ax.tick_params(axis='both', labelsize=10); self.ax.grid(axis='y', linestyle='--', alpha=0.3); self.canvas.draw()

class AdvancedPivotWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.df = None
        main_layout = QHBoxLayout(self)
        field_layout = QVBoxLayout()
        field_layout.addWidget(QLabel("Available Fields:", font=QFont("나눔스퀘어 네오 ExtraBold", 10)))
        self.list_fields = QListWidget(); self.list_fields.setDragEnabled(True)
        field_layout.addWidget(self.list_fields); main_layout.addLayout(field_layout, 1)
        zone_layout = QVBoxLayout()
        zone_layout.addWidget(QLabel("Rows (Drag here):", font=QFont("나눔스퀘어 네오 ExtraBold", 10)))
        self.list_rows = QListWidget(); self.list_rows.setAcceptDrops(True); self.list_rows.setDragEnabled(True)
        zone_layout.addWidget(self.list_rows)
        zone_layout.addWidget(QLabel("Columns (Drag here):", font=QFont("나눔스퀘어 네오 ExtraBold", 10)))
        self.list_cols = QListWidget(); self.list_cols.setAcceptDrops(True); self.list_cols.setDragEnabled(True)
        zone_layout.addWidget(self.list_cols)
        zone_layout.addWidget(QLabel("Values (Drag here):", font=QFont("나눔스퀘어 네오 ExtraBold", 10)))
        self.list_vals = QListWidget(); self.list_vals.setAcceptDrops(True); self.list_vals.setDragEnabled(True)
        zone_layout.addWidget(self.list_vals)
        agg_layout = QHBoxLayout(); agg_layout.addWidget(QLabel("Agg:"))
        self.group_agg = QButtonGroup(self); self.rb_sum = QRadioButton("Sum"); self.rb_mean = QRadioButton("Mean"); self.rb_sum.setChecked(True)
        self.group_agg.addButton(self.rb_sum); self.group_agg.addButton(self.rb_mean)
        agg_layout.addWidget(self.rb_sum); agg_layout.addWidget(self.rb_mean); zone_layout.addLayout(agg_layout)
        self.btn_run = QPushButton("Update Pivot"); self.btn_run.setFont(QFont("나눔스퀘어 네오 ExtraBold", 10)); self.btn_run.clicked.connect(self.run_pivot)
        zone_layout.addWidget(self.btn_run)
        self.btn_clear = QPushButton("Clear Fields"); self.btn_clear.clicked.connect(self.reset_fields)
        zone_layout.addWidget(self.btn_clear); main_layout.addLayout(zone_layout, 1)
        self.table = QTableWidget(); self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setStyleSheet("""QTableWidget { background-color: transparent; gridline-color: #d0d0d0; font-family: 'Malgun Gothic'; font-size: 10pt; } QHeaderView::section { background-color: #f0f0f0; padding: 4px; border: 1px solid #d0d0d0; font-weight: bold; }""")
        main_layout.addWidget(self.table, 3)
    def set_data(self, df): self.df = df; self.reset_fields()
    def reset_fields(self):
        self.list_fields.clear(); self.list_rows.clear(); self.list_cols.clear(); self.list_vals.clear(); self.table.clear(); self.table.setRowCount(0); self.table.setColumnCount(0)
        if self.df is not None:
            for col in self.df.columns: self.list_fields.addItem(col)
    def run_pivot(self):
        if self.df is None: return
        rows = [self.list_rows.item(i).text() for i in range(self.list_rows.count())]
        cols = [self.list_cols.item(i).text() for i in range(self.list_cols.count())]
        vals = [self.list_vals.item(i).text() for i in range(self.list_vals.count())]
        agg = 'mean' if self.rb_mean.isChecked() else 'sum'
        if not rows and not cols: QMessageBox.warning(self, "Warning", "Please select at least one Row or Column."); return
        if not vals: QMessageBox.warning(self, "Warning", "Please select at least one Value."); return
        try:
            pivot_df = self.df.copy()
            for v in vals: pivot_df[v] = pd.to_numeric(pivot_df[v], errors='coerce').fillna(0)
            pivoted = pivot_df.pivot_table(index=rows if rows else None, columns=cols if cols else None, values=vals, aggfunc=agg, fill_value=0)
            display_df = pivoted.reset_index()
            self.table.clear(); self.table.setRowCount(len(display_df.index)); self.table.setColumnCount(len(display_df.columns))
            flat_cols = []
            for c in display_df.columns: flat_cols.append(" - ".join(map(str, c)) if isinstance(c, tuple) else str(c))
            self.table.setHorizontalHeaderLabels(flat_cols)
            for i in range(len(display_df.index)):
                for j in range(len(display_df.columns)):
                    v = display_df.iloc[i, j]
                    try: fv = float(v); txt = f"{fv:,.2f}" if fv != 0 else "-"
                    except: txt = str(v)
                    item = QTableWidgetItem(txt); item.setTextAlignment(Qt.AlignCenter); self.table.setItem(i, j, item)
            self.table.resizeColumnsToContents()
        except Exception as e: QMessageBox.critical(self, "Pivot Error", f"Failed to create pivot table.\n{e}")

class PivotWidget(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("Please use the 'Custom Pivot' tab for advanced features."))

class HistoryChartWidget(BaseChartWidget):
    def __init__(self):
        super().__init__()
        self.fig.subplots_adjust(right=0.9, left=0.15, top=0.9, bottom=0.2)
        self.clear_plot()

    def clear_plot(self, message="Select a cell to see History"):
        super().clear_plot(message)

    def update_chart(self, history_df, model, quarter):
        pass





