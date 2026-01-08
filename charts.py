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
                    try: fval = float(val); self.annot.set_text(f"{fval/1000000:.2f} Mu")
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
        self.ax.imshow(mapped_data, aspect='auto'); self.ax.xaxis.tick_top()
        self.ax.set_xticks(range(len(data_df.columns))); self.ax.set_yticks(range(len(data_df.index)))
        self.ax.set_xticklabels(data_df.columns, fontsize=11, rotation=45, ha='left')
        self.ax.set_yticklabels(data_df.index, fontsize=12)
        sm = plt.cm.ScalarMappable(cmap=custom_cmap, norm=norm); sm.set_array([]); self.cbar = self.fig.colorbar(sm, ax=self.ax, fraction=0.046, pad=0.04)
        for i in range(len(data_df.index)):
            for j in range(len(data_df.columns)):
                val = data_df.iloc[i, j]; is_dark = abs(val) > 40 if fmt_type == "pct" else (abs(val) > (vmax * 0.6) if fmt_type == "diff" else val > (vmax * 0.5))
                text_color = "white" if is_dark else "black"; text_alpha = 1.0 if not self.selected_idx or (i, j) == self.selected_idx else 0.3
                txt = f"{val:+.1f}%" if fmt_type == "pct" else f"{val/1000000:.2f}"
                self.ax.text(j, i, txt, ha="center", va="center", color=text_color, fontsize=11, alpha=text_alpha)

class LineChartWidget(BaseChartWidget):
    def __init__(self, time_col="Week"): super().__init__(); self.time_col = time_col; self.current_data = None; self.clear_plot()
    def update_chart(self, full_df, brand, region, pivot_24, is_cumulative=False):
        self.ax.clear(); target_df = full_df.copy()
        if region != "Total": target_df = target_df[target_df['Region'] == region]
        if brand != "Total": target_df = target_df[target_df['Brand_Group'] == brand]
        target_df = target_df[target_df['Year'].isin([2023, 2024, 2025])]
        weekly_trend = target_df.pivot_table(index=self.time_col, columns="Year", values="Sales", aggfunc="sum")
        if is_cumulative: weekly_trend = weekly_trend.cumsum()
        self.current_data = weekly_trend; colors = {2023: config.COLOR_23, 2024: config.COLOR_24, 2025: config.COLOR_25}
        for y in [2023, 2024, 2025]:
            if y in weekly_trend.columns:
                data = weekly_trend[y].dropna(); self.ax.plot(data.index, data.values / 1000000.0, label=str(y), color=colors[y], linewidth=2.5)
        self.ax.legend(); self.canvas.draw()
    def copy_current_data(self):
        if self.current_data is not None: (self.current_data / 1000000.0).to_clipboard()

class TrendWidget(BaseChartWidget):
    def __init__(self, time_col="Week"):
        super().__init__(); self.time_col = time_col; self.is_vol_mode = False; self.clear_plot()
    def set_mode(self, is_checked): self.is_vol_mode = is_checked
    def update_chart(self, full_df, brand, region):
        self.ax.clear()
        df = full_df[full_df['Year'].isin([2023, 2024, 2025])].copy()
        pivot = df.pivot_table(index="Year", columns="Brand_Group" if region != "Total" else "Region", values="Sales", aggfunc="sum", fill_value=0)
        if not self.is_vol_mode: pivot = pivot.div(pivot.sum(axis=1), axis=0) * 100
        else: pivot = pivot / 1000000.0
        pivot.plot(kind='bar', stacked=True, ax=self.ax); self.canvas.draw()
    def copy_current_data(self): pass

class LaunchTrendWidget(BaseChartWidget):
    def update_chart(self, full_df, brand, category, visible_models, x_limit, is_cumulative):
        self.ax.clear()
        target = full_df[(full_df['Brand'] == brand) & (full_df['Category'] == category)]
        if visible_models: target = target[target['Model'].isin(visible_models)]
        pivot = target.pivot_table(index='MonthsSinceLaunch', columns='Model', values='Sales', aggfunc='sum')
        if is_cumulative: pivot = pivot.cumsum()
        (pivot / 1000000.0).plot(ax=self.ax); self.canvas.draw()
    def copy_current_data(self): pass

class LaunchTableWidget(QWidget):
    def __init__(self):
        super().__init__(); layout = QVBoxLayout(self); self.table = QTableWidget(); layout.addWidget(self.table)
    def update_table(self, df, brand, category, models, mode="Release", target_years=None):
        if df is None: self.table.clear(); return
        # Logic for populating table...
        pass
    def copy_current_data(self): pass
