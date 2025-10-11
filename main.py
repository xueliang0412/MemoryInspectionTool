# 1. 优化界面统计信息模块
# 2. 增加版本号显示
# 3. 统计信息保存至内存监控报告中

import psutil
import pandas as pd
import time
import threading
import os
import matplotlib.pyplot as plt
from datetime import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
from io import BytesIO
import matplotlib.dates as mdates
import numpy as np

matplotlib.use('TkAgg')
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]


class MemoryMonitorApp:
    def __init__(self, root):
        self.root = root
        # ---------------------- 修改1：添加版本号v1.0 ----------------------
        self.root.title("多进程内存监控工具 v1.0")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)

        # 初始化变量
        self.monitoring = False
        self.monitor_thread = None
        self.process_data = {}  # {进程名: DataFrame}
        self.selected_processes = set()
        self.merge_processes = set()
        self.save_path = os.getcwd()

        # 统计信息相关变量
        self.stats_tree = None

        # 创建UI界面
        self._create_widgets()

    def _create_widgets(self):
        """创建UI组件"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. 进程选择区域
        process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
        process_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
        self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
        self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)

        scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
        scrollbar.grid(row=1, column=1, sticky=tk.NS)
        self.process_listbox.config(yscrollcommand=scrollbar.set)

        btn_frame = ttk.Frame(process_frame)
        btn_frame.grid(row=1, column=2, padx=10)
        ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)

        ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
        self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
        self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)

        # 2. 参数设置区域
        param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
        param_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(param_frame, text="监控时长:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.duration_var = tk.StringVar(value="60")
        ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        self.duration_unit = ttk.Combobox(param_frame, values=["秒", "分钟", "小时"], width=6, state="readonly")
        self.duration_unit.current(0)
        self.duration_unit.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(param_frame, text="采样间隔:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
        self.interval_var = tk.StringVar(value="5")
        ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=4, sticky=tk.W, pady=5)
        self.interval_unit = ttk.Combobox(param_frame, values=["秒", "分钟"], width=6, state="readonly")
        self.interval_unit.current(0)
        self.interval_unit.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)

        ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
        self.path_var = tk.StringVar(value=os.getcwd())
        ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=7, sticky=tk.W, pady=5)
        ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=8, padx=5, pady=5)

        # 3. 图表合并配置
        merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
        merge_frame.pack(fill=tk.X, pady=(0, 10))

        self.merge_var = tk.BooleanVar(value=False)
        merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
                                      command=self._toggle_merge_options)
        merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)

        ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
        self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
        self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)

        merge_btn_frame = ttk.Frame(merge_frame)
        merge_btn_frame.grid(row=2, column=1, padx=10)
        self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge,
                                           state=tk.DISABLED)
        self.add_to_merge_btn.pack(fill=tk.X, pady=5)
        self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge,
                                                state=tk.DISABLED)
        self.remove_from_merge_btn.pack(fill=tk.X, pady=5)

        ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
        self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
        self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)

        # 4. 控制按钮区域
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
        self.report_btn.pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)

        # 5. 实时图表区域
        chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
        chart_frame.pack(fill=tk.BOTH, expand=True)  # 图表区域占满剩余空间

        self.fig, self.ax = plt.subplots(figsize=(10, 6))
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 初始化悬停标注框
        self.annotation = self.ax.annotate(
            "",
            xy=(0, 0),
            xytext=(15, 15),
            textcoords="offset points",
            bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="gray", alpha=0.9),
            arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0.2")
        )
        self.annotation.set_visible(False)
        self.canvas.mpl_connect("motion_notify_event", self._on_mouse_hover)

        # ---------------------- 修改2：调整统计信息模块位置（上移以完全显示） ----------------------
        stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="10")
        # 关键调整：减少上方边距，取消side=tk.BOTTOM，通过pack顺序自然位于图表下方
        stats_frame.pack(fill=tk.X, pady=(5, 0))

        # 创建统计数据表格（Treeview控件）并设置居中显示
        columns = ("proc", "max", "min", "avg", "3sigma")
        self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show="headings", height=6)  # 增加height为6行

        # 设置表头
        self.stats_tree.heading("proc", text="进程名")
        self.stats_tree.heading("max", text="最大值 (MB)")
        self.stats_tree.heading("min", text="最小值 (MB)")
        self.stats_tree.heading("avg", text="平均值 (MB)")
        self.stats_tree.heading("3sigma", text="3σ值 (MB)")

        # 设置列宽和居中对齐
        self.stats_tree.column("proc", width=150, anchor="center")
        self.stats_tree.column("max", width=100, anchor="center")
        self.stats_tree.column("min", width=100, anchor="center")
        self.stats_tree.column("avg", width=100, anchor="center")
        self.stats_tree.column("3sigma", width=100, anchor="center")

        self.stats_tree.pack(fill=tk.X)

        self._refresh_processes()

    # 图表合并相关方法
    def _toggle_merge_options(self):
        state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
        self.merge_source_listbox.config(state=state)
        self.merge_target_listbox.config(state=state)
        self.add_to_merge_btn.config(state=state)
        self.remove_from_merge_btn.config(state=state)
        if state == tk.NORMAL:
            self._sync_merge_source_list()

    def _sync_merge_source_list(self):
        self.merge_source_listbox.delete(0, tk.END)
        for i in range(self.monitor_listbox.size()):
            proc_name = self.monitor_listbox.get(i)
            self.merge_source_listbox.insert(tk.END, proc_name)

    def _add_to_merge(self):
        selected_indices = self.merge_source_listbox.curselection()
        for i in selected_indices:
            proc_name = self.merge_source_listbox.get(i)
            if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
                self.merge_target_listbox.insert(tk.END, proc_name)

    def _remove_from_merge(self):
        selected_indices = self.merge_target_listbox.curselection()
        for i in sorted(selected_indices, reverse=True):
            self.merge_target_listbox.delete(i)

    # 进程管理相关方法
    def _refresh_processes(self):
        self.process_listbox.delete(0, tk.END)
        processes = set()
        for proc in psutil.process_iter(['name']):
            try:
                processes.add(proc.info['name'])
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        for proc_name in sorted(processes):
            self.process_listbox.insert(tk.END, proc_name)

    def _add_monitor(self):
        selected_indices = self.process_listbox.curselection()
        for i in selected_indices:
            proc_name = self.process_listbox.get(i)
            if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
                self.monitor_listbox.insert(tk.END, proc_name)
                if self.merge_var.get():
                    self._sync_merge_source_list()

    def _remove_monitor(self):
        selected_indices = self.monitor_listbox.curselection()
        for i in sorted(selected_indices, reverse=True):
            self.monitor_listbox.delete(i)
            if self.merge_var.get():
                self._sync_merge_source_list()

    def _browse_path(self):
        path = filedialog.askdirectory()
        if path:
            self.path_var.set(path)
            self.save_path = path

    # 监控控制方法
    def _start_monitoring(self):
        if self.monitor_listbox.size() == 0:
            messagebox.showwarning("警告", "请至少选择一个进程进行监控")
            return

        try:
            duration_value = int(self.duration_var.get())
            duration_unit = self.duration_unit.get()
            if duration_unit == "分钟":
                duration = duration_value * 60
            elif duration_unit == "小时":
                duration = duration_value * 3600
            else:
                duration = duration_value

            interval_value = int(self.interval_var.get())
            interval_unit = self.interval_unit.get()
            if interval_unit == "分钟":
                interval = interval_value * 60
            else:
                interval = interval_value

            if duration <= 0 or interval <= 0 or interval > duration:
                raise ValueError
        except ValueError:
            messagebox.showwarning("警告", "请输入有效的监控参数（正整数，且间隔不大于时长）")
            return

        self.process_data = {}
        for i in range(self.monitor_listbox.size()):
            proc_name = self.monitor_listbox.get(i)
            self.process_data[proc_name] = []

        self.status_var.set("监控中...")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.report_btn.config(state=tk.DISABLED)
        self.monitoring = True

        self.monitor_thread = threading.Thread(
            target=self._monitor_processes,
            args=(duration, interval),
            daemon=True
        )
        self.monitor_thread.start()

    def _stop_monitoring(self):
        self.monitoring = False
        self.status_var.set("监控已停止，准备生成报告")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.report_btn.config(state=tk.NORMAL)

    def _monitor_processes(self, duration, interval):
        end_time = time.time() + duration
        while self.monitoring and time.time() < end_time:
            timestamp = datetime.now()
            for proc_name in self.process_data.keys():
                mem_usage = 0
                for proc in psutil.process_iter(['name', 'memory_info']):
                    try:
                        if proc.info['name'] == proc_name:
                            mem_usage += proc.info['memory_info'].rss
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        continue
                self.process_data[proc_name].append({'Timestamp': timestamp, 'Memory_Bytes': mem_usage})
            self.root.after(0, self._update_chart)
            time.sleep(interval)
        if self.monitoring:
            self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
            self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
            self.monitoring = False

    # 更新统计信息表格
    def _update_stats_table(self, stats_data):
        """更新UI中的统计信息表格"""
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)
        for data in stats_data:
            self.stats_tree.insert("", "end", values=data)

    # 更新图表
    def _update_chart(self):
        """更新实时图表"""
        self.ax.clear()
        has_data = False
        stats_data = []

        for proc_name, data in self.process_data.items():
            if data:
                has_data = True
                df = pd.DataFrame(data)
                df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
                self.ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)

                # 计算统计值
                memory_values = df['Memory_MB'].values
                max_val = np.max(memory_values) if len(memory_values) > 0 else 0
                min_val = np.min(memory_values) if len(memory_values) > 0 else 0
                avg_val = np.mean(memory_values) if len(memory_values) > 0 else 0
                std_val = np.std(memory_values, ddof=1) if len(memory_values) >= 2 else 0
                sigma3_val = 3 * std_val
                stats_data.append(
                    (proc_name, round(max_val, 2), round(min_val, 2), round(avg_val, 2), round(sigma3_val, 2)))

        # 更新统计表格
        if stats_data:
            self._update_stats_table(stats_data)

        # import matplotlib.pyplot as plt
        # import matplotlib.dates as mdates
        # import pandas as pd  # 假设用户已导入pandas（df为DataFrame对象）

        # ---------------------- 添加字体配置 ----------------------
        plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC", "Microsoft YaHei"]
        plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

        # 图表样式设置
        if has_data:
            self.ax.set_title('实时内存使用监控')
            self.ax.set_xlabel('时间')
            self.ax.set_ylabel('内存使用 (MB)')
            self.ax.legend(loc='upper left')
            max_ticks = min(10, len(df))
            self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
            self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
            plt.xticks(rotation=45, ha='right')
            self.fig.tight_layout()

        self.canvas.draw()

    # 鼠标悬停事件
    def _on_mouse_hover(self, event):
        if event.inaxes != self.ax:
            self.annotation.set_visible(False)
            self.canvas.draw_idle()
            return
        if not self.ax.lines:
            self.annotation.set_visible(False)
            self.canvas.draw_idle()
            return
        mouse_x = event.xdata
        mouse_y = event.ydata
        if mouse_x is None or mouse_y is None:
            self.annotation.set_visible(False)
            self.canvas.draw_idle()
            return
        mouse_datetime = mdates.num2date(mouse_x)
        min_dist = float('inf')
        closest_point = None
        for line in self.ax.lines:
            proc_name = line.get_label()
            x_data = line.get_xdata()
            y_data = line.get_ydata()
            for i in range(len(x_data)):
                point_datetime = mdates.num2date(x_data[i])
                time_diff = abs((mouse_datetime - point_datetime).total_seconds())
                memory_diff = abs(mouse_y - y_data[i])
                if time_diff < 5 and memory_diff < 10:
                    distance = time_diff * 0.1 + memory_diff * 0.01
                    if distance < min_dist:
                        min_dist = distance
                        closest_point = {
                            "proc": proc_name,
                            "time": point_datetime.strftime("%Y-%m-%d %H:%M:%S"),
                            "memory": round(y_data[i], 2)
                        }
        if closest_point:
            self.annotation.xy = (mouse_x, mouse_y)
            self.annotation.set_text(
                f"进程: {closest_point['proc']}\n"
                f"时间: {closest_point['time']}\n"
                f"内存: {closest_point['memory']} MB"
            )
            self.annotation.set_visible(True)
        else:
            self.annotation.set_visible(False)
        self.canvas.draw_idle()

    # ---------------------- 修改3：完善统计信息保存至Excel ----------------------
    def _generate_report(self):
        """生成Excel报告（确保统计信息保存）"""
        if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
            messagebox.showwarning("警告", "没有监控数据可生成报告")
            return

        # 报告路径调整为"内存监控报告.xlsx"
        excel_path = os.path.join(self.save_path, "内存监控报告.xlsx")
        wb = Workbook()

        # 1. 创建内存数据工作表
        data_ws = wb.active
        data_ws.title = "内存监控数据"

        # 写入内存数据
        summary_data = {}
        process_names = [self.monitor_listbox.get(i) for i in range(self.monitor_listbox.size())]
        for proc_name in process_names:
            if proc_name in self.process_data and self.process_data[proc_name]:
                df = pd.DataFrame(self.process_data[proc_name])
                df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
                df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')

                for _, row in df.iterrows():
                    timestamp = row['Timestamp']
                    if timestamp not in summary_data:
                        summary_data[timestamp] = {'Timestamp': timestamp, 'Time': row['Time']}
                    summary_data[timestamp][proc_name] = row['Memory_MB']

        if summary_data:
            summary_df = pd.DataFrame.from_dict(summary_data, orient='index').sort_values('Timestamp')
            headers = ['时间戳'] + process_names
            data_ws.append(headers)

            for _, row in summary_df.iterrows():
                row_data = [row['Timestamp']] + [row.get(proc, '') for proc in process_names]
                data_ws.append(row_data)

            # 设置数据表格格式
            center_alignment = Alignment(horizontal='center', vertical='center')
            for row in data_ws.iter_rows(min_row=1, max_row=data_ws.max_row, min_col=1, max_col=data_ws.max_column):
                for cell in row:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '0.00'
                    cell.alignment = center_alignment
            data_ws.column_dimensions['A'].width = 25
            for col in range(2, 2 + len(process_names)):
                data_ws.column_dimensions[chr(64 + col)].width = 18

        # 2. 创建统计汇总工作表（确保统计信息保存）
        stats_ws = wb.create_sheet(title="统计汇总")
        stats_ws.append(["进程名", "最大值 (MB)", "最小值 (MB)", "平均值 (MB)", "3σ值 (MB)"])

        # 计算并写入统计值
        for proc_name in process_names:
            if proc_name in self.process_data and self.process_data[proc_name]:
                df = pd.DataFrame(self.process_data[proc_name])
                df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
                memory_values = df['Memory_MB'].values

                max_val = np.max(memory_values) if len(memory_values) > 0 else 0
                min_val = np.min(memory_values) if len(memory_values) > 0 else 0
                avg_val = np.mean(memory_values) if len(memory_values) > 0 else 0
                std_val = np.std(memory_values, ddof=1) if len(memory_values) >= 2 else 0  # 样本标准差
                sigma3_val = 3 * std_val

                stats_ws.append([
                    proc_name,
                    round(max_val, 2),
                    round(min_val, 2),
                    round(avg_val, 2),
                    round(sigma3_val, 2)
                ])

        # 设置统计表格格式
        for row in stats_ws.iter_rows(min_row=2, max_row=stats_ws.max_row, min_col=2, max_col=5):
            for cell in row:
                cell.alignment = Alignment(horizontal='center')
        stats_ws.column_dimensions['A'].width = 15
        for col in ['B', 'C', 'D', 'E']:
            stats_ws.column_dimensions[col].width = 12

        # 生成图表
        current_row = len(summary_df) + 3 if summary_data else 1
        merge_chart_inserted = False
        if self.merge_var.get() and self.merge_target_listbox.size() > 0:
            merge_procs = [self.merge_target_listbox.get(i) for i in range(self.merge_target_listbox.size())]
            if merge_procs:
                fig, ax = plt.subplots(figsize=(12, 6))
                colors = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
                color_idx = 0
                for proc_name in merge_procs:
                    if proc_name in self.process_data and self.process_data[proc_name]:
                        df = pd.DataFrame(self.process_data[proc_name])
                        df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
                        ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-',
                                label=proc_name, color=colors[color_idx % len(colors)])
                        color_idx += 1
                ax.set_title('多进程内存使用对比（合并图表）')
                ax.set_xlabel('时间')
                ax.set_ylabel('内存使用 (MB)')
                ax.legend()
                max_ticks = min(10, len(df)) if 'df' in locals() else 5
                ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                img_data = BytesIO()
                fig.savefig(img_data, format='png')
                img_data.seek(0)
                img = Image(img_data)
                img.width = 800
                img.height = 400
                data_ws.add_image(img, f'A{current_row}')
                current_row += 30
                plt.close(fig)
                merge_chart_inserted = True

        if not merge_chart_inserted:
            current_row = len(summary_df) + 3 if summary_data else 1
        else:
            current_row += 5

        for proc_name in process_names:
            if proc_name in self.process_data and self.process_data[proc_name]:
                df = pd.DataFrame(self.process_data[proc_name])
                df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
                fig, ax = plt.subplots(figsize=(10, 4))
                ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', color='blue')
                ax.set_title(f'{proc_name} 内存使用趋势')
                ax.set_xlabel('时间')
                ax.set_ylabel('内存使用 (MB)')
                max_ticks = min(10, len(df))
                ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                img_data = BytesIO()
                fig.savefig(img_data, format='png')
                img_data.seek(0)
                img = Image(img_data)
                img.width = 600
                img.height = 300
                data_ws.add_image(img, f'A{current_row}')
                current_row += 20
                plt.close(fig)

        try:
            wb.save(excel_path)
            messagebox.showinfo("成功", f"报告已生成：\n{excel_path}")
            self.status_var.set("报告生成完成")
        except Exception as e:
            messagebox.showerror("错误", f"保存报告失败：{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = MemoryMonitorApp(root)
    root.mainloop()



# # #界面优化：统计信息模块放置界面下方
# # #去掉界面折线图表的统计信息展示
# # #增加各进程的最大值，最小值，平均值，方差，标准差，百分位数等信息
# import psutil
# import pandas as pd
# import time
# import threading
# import os
# import matplotlib.pyplot as plt
# from datetime import datetime
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Alignment
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# import matplotlib
# from io import BytesIO
# import matplotlib.dates as mdates
# import numpy as np  # 用于计算统计值
#
# matplotlib.use('TkAgg')
# plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
#
#
# class MemoryMonitorApp:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("多进程内存监控工具")
#         self.root.geometry("1000x800")
#         self.root.resizable(True, True)
#
#         # 初始化变量
#         self.monitoring = False
#         self.monitor_thread = None
#         self.process_data = {}  # 存储监控数据 {进程名: DataFrame}
#         self.selected_processes = set()
#         self.merge_processes = set()
#         self.save_path = os.getcwd()
#
#         # 统计信息相关变量
#         self.stats_tree = None  # 用于显示统计数据的表格控件
#
#         # 创建UI界面
#         self._create_widgets()
#
#     def _create_widgets(self):
#         """创建UI组件"""
#         main_frame = ttk.Frame(self.root, padding="10")
#         main_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 1. 进程选择区域
#         process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
#         process_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
#         self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)
#
#         scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
#         scrollbar.grid(row=1, column=1, sticky=tk.NS)
#         self.process_listbox.config(yscrollcommand=scrollbar.set)
#
#         btn_frame = ttk.Frame(process_frame)
#         btn_frame.grid(row=1, column=2, padx=10)
#         ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)
#
#         ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
#         self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)
#
#         # 2. 参数设置区域
#         param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
#         param_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(param_frame, text="监控时长:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
#         self.duration_var = tk.StringVar(value="60")
#         ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
#         self.duration_unit = ttk.Combobox(param_frame, values=["秒", "分钟", "小时"], width=6, state="readonly")
#         self.duration_unit.current(0)
#         self.duration_unit.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
#
#         ttk.Label(param_frame, text="采样间隔:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
#         self.interval_var = tk.StringVar(value="5")
#         ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=4, sticky=tk.W, pady=5)
#         self.interval_unit = ttk.Combobox(param_frame, values=["秒", "分钟"], width=6, state="readonly")
#         self.interval_unit.current(0)
#         self.interval_unit.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
#
#         ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
#         self.path_var = tk.StringVar(value=os.getcwd())
#         ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=7, sticky=tk.W, pady=5)
#         ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=8, padx=5, pady=5)
#
#         # 3. 图表合并配置
#         merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
#         merge_frame.pack(fill=tk.X, pady=(0, 10))
#
#         self.merge_var = tk.BooleanVar(value=False)
#         merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
#                                       command=self._toggle_merge_options)
#         merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
#
#         ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
#         self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)
#
#         merge_btn_frame = ttk.Frame(merge_frame)
#         merge_btn_frame.grid(row=2, column=1, padx=10)
#         self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge,
#                                            state=tk.DISABLED)
#         self.add_to_merge_btn.pack(fill=tk.X, pady=5)
#         self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge,
#                                                 state=tk.DISABLED)
#         self.remove_from_merge_btn.pack(fill=tk.X, pady=5)
#
#         ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
#         self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)
#
#         # 4. 控制按钮区域
#         control_frame = ttk.Frame(main_frame)
#         control_frame.pack(fill=tk.X, pady=(0, 10))
#         self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
#         self.start_btn.pack(side=tk.LEFT, padx=5)
#         self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
#         self.stop_btn.pack(side=tk.LEFT, padx=5)
#         self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
#         self.report_btn.pack(side=tk.LEFT, padx=5)
#         self.status_var = tk.StringVar(value="就绪")
#         ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
#
#         # 5. 实时图表区域
#         chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         chart_frame.pack(fill=tk.BOTH, expand=True)
#
#         self.fig, self.ax = plt.subplots(figsize=(10, 6))
#         self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#
#         # 初始化悬停标注框
#         self.annotation = self.ax.annotate(
#             "",
#             xy=(0, 0),
#             xytext=(15, 15),
#             textcoords="offset points",
#             bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="gray", alpha=0.9),
#             arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0.2")
#         )
#         self.annotation.set_visible(False)
#         self.canvas.mpl_connect("motion_notify_event", self._on_mouse_hover)
#
#         # ---------------------- 修改1：统计信息模块移至最底部并居中显示 ----------------------
#         stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="10")
#         stats_frame.pack(fill=tk.X, pady=(10, 0), side=tk.BOTTOM)  # 放置在底部
#
#         # 创建统计数据表格（Treeview控件）并设置居中显示
#         columns = ("proc", "max", "min", "avg", "3sigma")
#         self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show="headings", height=5)
#
#         # 设置表头
#         self.stats_tree.heading("proc", text="进程名")
#         self.stats_tree.heading("max", text="最大值 (MB)")
#         self.stats_tree.heading("min", text="最小值 (MB)")
#         self.stats_tree.heading("avg", text="平均值 (MB)")
#         self.stats_tree.heading("3sigma", text="3σ值 (MB)")
#
#         # 设置列宽和居中对齐（关键：anchor="center"）
#         self.stats_tree.column("proc", width=150, anchor="center")  # 进程名居中
#         self.stats_tree.column("max", width=100, anchor="center")  # 数值列居中
#         self.stats_tree.column("min", width=100, anchor="center")
#         self.stats_tree.column("avg", width=100, anchor="center")
#         self.stats_tree.column("3sigma", width=100, anchor="center")
#
#         self.stats_tree.pack(fill=tk.X)
#
#         self._refresh_processes()
#
#     # 图表合并相关方法
#     def _toggle_merge_options(self):
#         state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
#         self.merge_source_listbox.config(state=state)
#         self.merge_target_listbox.config(state=state)
#         self.add_to_merge_btn.config(state=state)
#         self.remove_from_merge_btn.config(state=state)
#         if state == tk.NORMAL:
#             self._sync_merge_source_list()
#
#     def _sync_merge_source_list(self):
#         self.merge_source_listbox.delete(0, tk.END)
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.merge_source_listbox.insert(tk.END, proc_name)
#
#     def _add_to_merge(self):
#         selected_indices = self.merge_source_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.merge_source_listbox.get(i)
#             if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
#                 self.merge_target_listbox.insert(tk.END, proc_name)
#
#     def _remove_from_merge(self):
#         selected_indices = self.merge_target_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.merge_target_listbox.delete(i)
#
#     # 进程管理相关方法
#     def _refresh_processes(self):
#         self.process_listbox.delete(0, tk.END)
#         processes = set()
#         for proc in psutil.process_iter(['name']):
#             try:
#                 processes.add(proc.info['name'])
#             except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                 continue
#         for proc_name in sorted(processes):
#             self.process_listbox.insert(tk.END, proc_name)
#
#     def _add_monitor(self):
#         selected_indices = self.process_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.process_listbox.get(i)
#             if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
#                 self.monitor_listbox.insert(tk.END, proc_name)
#                 if self.merge_var.get():
#                     self._sync_merge_source_list()
#
#     def _remove_monitor(self):
#         selected_indices = self.monitor_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.monitor_listbox.delete(i)
#             if self.merge_var.get():
#                 self._sync_merge_source_list()
#
#     def _browse_path(self):
#         path = filedialog.askdirectory()
#         if path:
#             self.path_var.set(path)
#             self.save_path = path
#
#     # 监控控制方法
#     def _start_monitoring(self):
#         if self.monitor_listbox.size() == 0:
#             messagebox.showwarning("警告", "请至少选择一个进程进行监控")
#             return
#
#         try:
#             duration_value = int(self.duration_var.get())
#             duration_unit = self.duration_unit.get()
#             if duration_unit == "分钟":
#                 duration = duration_value * 60
#             elif duration_unit == "小时":
#                 duration = duration_value * 3600
#             else:
#                 duration = duration_value
#
#             interval_value = int(self.interval_var.get())
#             interval_unit = self.interval_unit.get()
#             if interval_unit == "分钟":
#                 interval = interval_value * 60
#             else:
#                 interval = interval_value
#
#             if duration <= 0 or interval <= 0 or interval > duration:
#                 raise ValueError
#         except ValueError:
#             messagebox.showwarning("警告", "请输入有效的监控参数（正整数，且间隔不大于时长）")
#             return
#
#         self.process_data = {}
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.process_data[proc_name] = []
#
#         self.status_var.set("监控中...")
#         self.start_btn.config(state=tk.DISABLED)
#         self.stop_btn.config(state=tk.NORMAL)
#         self.report_btn.config(state=tk.DISABLED)
#         self.monitoring = True
#
#         self.monitor_thread = threading.Thread(
#             target=self._monitor_processes,
#             args=(duration, interval),
#             daemon=True
#         )
#         self.monitor_thread.start()
#
#     def _stop_monitoring(self):
#         self.monitoring = False
#         self.status_var.set("监控已停止，准备生成报告")
#         self.start_btn.config(state=tk.NORMAL)
#         self.stop_btn.config(state=tk.DISABLED)
#         self.report_btn.config(state=tk.NORMAL)
#
#     def _monitor_processes(self, duration, interval):
#         end_time = time.time() + duration
#         while self.monitoring and time.time() < end_time:
#             timestamp = datetime.now()
#             for proc_name in self.process_data.keys():
#                 mem_usage = 0
#                 for proc in psutil.process_iter(['name', 'memory_info']):
#                     try:
#                         if proc.info['name'] == proc_name:
#                             mem_usage += proc.info['memory_info'].rss
#                     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                         continue
#                 self.process_data[proc_name].append({'Timestamp': timestamp, 'Memory_Bytes': mem_usage})
#             self.root.after(0, self._update_chart)
#             time.sleep(interval)
#         if self.monitoring:
#             self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
#             self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
#             self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
#             self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
#             self.monitoring = False
#
#     # 更新统计信息表格
#     def _update_stats_table(self, stats_data):
#         """更新UI中的统计信息表格"""
#         for item in self.stats_tree.get_children():
#             self.stats_tree.delete(item)
#         for data in stats_data:
#             self.stats_tree.insert("", "end", values=data)
#
#     # ---------------------- 修改2：移除图表中的统计信息展示 ----------------------
#     def _update_chart(self):
#         """更新实时图表（移除统计信息标注）"""
#         self.ax.clear()
#         has_data = False
#         stats_data = []
#
#         for proc_name, data in self.process_data.items():
#             if data:
#                 has_data = True
#                 df = pd.DataFrame(data)
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 self.ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)
#
#                 # 计算统计值（仅用于表格，不用于图表标注）
#                 memory_values = df['Memory_MB'].values
#                 max_val = np.max(memory_values) if len(memory_values) > 0 else 0
#                 min_val = np.min(memory_values) if len(memory_values) > 0 else 0
#                 avg_val = np.mean(memory_values) if len(memory_values) > 0 else 0
#                 std_val = np.std(memory_values, ddof=1) if len(memory_values) >= 2 else 0
#                 sigma3_val = 3 * std_val
#                 stats_data.append(
#                     (proc_name, round(max_val, 2), round(min_val, 2), round(avg_val, 2), round(sigma3_val, 2)))
#
#         # 更新统计表格
#         if stats_data:
#             self._update_stats_table(stats_data)
#
#         # 图表样式设置（无统计信息标注）
#         if has_data:
#             self.ax.set_title('实时内存使用监控')  # 简化标题
#             self.ax.set_xlabel('时间')
#             self.ax.set_ylabel('内存使用 (MB)')
#             self.ax.legend(loc='upper left')
#             max_ticks = min(10, len(df))
#             self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#             self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#             plt.xticks(rotation=45, ha='right')
#             self.fig.tight_layout()
#
#         self.canvas.draw()
#
#     # 鼠标悬停事件
#     def _on_mouse_hover(self, event):
#         if event.inaxes != self.ax:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#         if not self.ax.lines:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#         mouse_x = event.xdata
#         mouse_y = event.ydata
#         if mouse_x is None or mouse_y is None:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#         mouse_datetime = mdates.num2date(mouse_x)
#         min_dist = float('inf')
#         closest_point = None
#         for line in self.ax.lines:
#             proc_name = line.get_label()
#             x_data = line.get_xdata()
#             y_data = line.get_ydata()
#             for i in range(len(x_data)):
#                 point_datetime = mdates.num2date(x_data[i])
#                 time_diff = abs((mouse_datetime - point_datetime).total_seconds())
#                 memory_diff = abs(mouse_y - y_data[i])
#                 if time_diff < 5 and memory_diff < 10:
#                     distance = time_diff * 0.1 + memory_diff * 0.01
#                     if distance < min_dist:
#                         min_dist = distance
#                         closest_point = {
#                             "proc": proc_name,
#                             "time": point_datetime.strftime("%Y-%m-%d %H:%M:%S"),
#                             "memory": round(y_data[i], 2)
#                         }
#         if closest_point:
#             self.annotation.xy = (mouse_x, mouse_y)
#             self.annotation.set_text(
#                 f"进程: {closest_point['proc']}\n"
#                 f"时间: {closest_point['time']}\n"
#                 f"内存: {closest_point['memory']} MB"
#             )
#             self.annotation.set_visible(True)
#         else:
#             self.annotation.set_visible(False)
#         self.canvas.draw_idle()
#
#     # ---------------------- 修改3：Excel报告自动计算统计值 ----------------------
#     def _generate_report(self):
#         """生成Excel报告并自动计算统计值"""
#         if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
#             messagebox.showwarning("警告", "没有监控数据可生成报告")
#             return
#
#         excel_path = os.path.join(self.save_path, "memory_summary_report.xlsx")
#         wb = Workbook()
#
#         # 1. 创建内存数据工作表
#         data_ws = wb.active
#         data_ws.title = "内存监控数据"
#
#         # 写入内存数据
#         summary_data = {}
#         process_names = [self.monitor_listbox.get(i) for i in range(self.monitor_listbox.size())]
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')
#
#                 for _, row in df.iterrows():
#                     timestamp = row['Timestamp']
#                     if timestamp not in summary_data:
#                         summary_data[timestamp] = {'Timestamp': timestamp, 'Time': row['Time']}
#                     summary_data[timestamp][proc_name] = row['Memory_MB']
#
#         if summary_data:
#             summary_df = pd.DataFrame.from_dict(summary_data, orient='index').sort_values('Timestamp')
#             headers = ['Time'] + process_names
#             data_ws.append(headers)
#
#             for _, row in summary_df.iterrows():
#                 row_data = [row['Timestamp']] + [row.get(proc, '') for proc in process_names]
#                 data_ws.append(row_data)
#
#             # 设置数据表格格式
#             center_alignment = Alignment(horizontal='center', vertical='center')
#             for row in data_ws.iter_rows(min_row=1, max_row=data_ws.max_row, min_col=1, max_col=data_ws.max_column):
#                 for cell in row:
#                     if isinstance(cell.value, (int, float)):
#                         cell.number_format = '0.00'
#                     cell.alignment = center_alignment
#             data_ws.column_dimensions['A'].width = 25
#             for col in range(2, 2 + len(process_names)):
#                 data_ws.column_dimensions[chr(64 + col)].width = 18
#
#         # ---------------------- 新增：创建统计汇总工作表 ----------------------
#         stats_ws = wb.create_sheet(title="统计汇总")
#         stats_ws.append(["进程名", "最大值 (MB)", "最小值 (MB)", "平均值 (MB)", "3σ值 (MB)"])
#
#         # 计算并写入统计值
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 memory_values = df['Memory_MB'].values
#
#                 max_val = np.max(memory_values) if len(memory_values) > 0 else 0
#                 min_val = np.min(memory_values) if len(memory_values) > 0 else 0
#                 avg_val = np.mean(memory_values) if len(memory_values) > 0 else 0
#                 std_val = np.std(memory_values, ddof=1) if len(memory_values) >= 2 else 0  # 样本标准差
#                 sigma3_val = 3 * std_val
#
#                 stats_ws.append([
#                     proc_name,
#                     round(max_val, 2),
#                     round(min_val, 2),
#                     round(avg_val, 2),
#                     round(sigma3_val, 2)
#                 ])
#
#         # 设置统计表格格式（居中对齐）
#         for row in stats_ws.iter_rows(min_row=2, max_row=stats_ws.max_row, min_col=2, max_col=5):
#             for cell in row:
#                 cell.alignment = Alignment(horizontal='center')
#         stats_ws.column_dimensions['A'].width = 15
#         for col in ['B', 'C', 'D', 'E']:
#             stats_ws.column_dimensions[col].width = 12
#
#         # 生成图表（保持原有逻辑）
#         current_row = len(summary_df) + 3 if summary_data else 1
#         merge_chart_inserted = False
#         if self.merge_var.get() and self.merge_target_listbox.size() > 0:
#             merge_procs = [self.merge_target_listbox.get(i) for i in range(self.merge_target_listbox.size())]
#             if merge_procs:
#                 fig, ax = plt.subplots(figsize=(12, 6))
#                 colors = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
#                 color_idx = 0
#                 for proc_name in merge_procs:
#                     if proc_name in self.process_data and self.process_data[proc_name]:
#                         df = pd.DataFrame(self.process_data[proc_name])
#                         df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                         ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-',
#                                 label=proc_name, color=colors[color_idx % len(colors)])
#                         color_idx += 1
#                 ax.set_title('多进程内存使用对比（合并图表）')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#                 ax.legend()
#                 max_ticks = min(10, len(df)) if 'df' in locals() else 5
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 800
#                 img.height = 400
#                 data_ws.add_image(img, f'A{current_row}')
#                 current_row += 30
#                 plt.close(fig)
#                 merge_chart_inserted = True
#
#         if not merge_chart_inserted:
#             current_row = len(summary_df) + 3 if summary_data else 1
#         else:
#             current_row += 5
#
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 fig, ax = plt.subplots(figsize=(10, 4))
#                 ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', color='blue')
#                 ax.set_title(f'{proc_name} 内存使用趋势')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 600
#                 img.height = 300
#                 data_ws.add_image(img, f'A{current_row}')
#                 current_row += 20
#                 plt.close(fig)
#
#         try:
#             wb.save(excel_path)
#             messagebox.showinfo("成功", f"汇总报告已生成（含统计值）：\n{excel_path}")
#             self.status_var.set("报告生成完成")
#         except Exception as e:
#             messagebox.showerror("错误", f"保存报告失败：{str(e)}")
#
#
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = MemoryMonitorApp(root)
#     root.mainloop()



# #增加各进程的最大值，最小值，平均值，方差，标准差，百分位数等信息
# import psutil
# import pandas as pd
# import time
# import threading
# import os
# import matplotlib.pyplot as plt
# from datetime import datetime
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Alignment
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# import matplotlib
# from io import BytesIO
# import matplotlib.dates as mdates
# import numpy as np  # 新增：用于计算标准差
#
# matplotlib.use('TkAgg')
# plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
#
#
# class MemoryMonitorApp:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("多进程内存监控工具")
#         self.root.geometry("1000x800")  # 扩大窗口尺寸以容纳统计信息
#         self.root.resizable(True, True)
#
#         # 初始化变量
#         self.monitoring = False
#         self.monitor_thread = None
#         self.process_data = {}  # 存储监控数据 {进程名: DataFrame}
#         self.selected_processes = set()
#         self.merge_processes = set()
#         self.save_path = os.getcwd()
#
#         # 新增：统计信息相关变量
#         self.stats_tree = None  # 用于显示统计数据的表格控件
#
#         # 创建UI界面
#         self._create_widgets()
#
#     def _create_widgets(self):
#         """创建UI组件（新增统计信息面板）"""
#         main_frame = ttk.Frame(self.root, padding="10")
#         main_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 1. 进程选择区域（保持不变）
#         process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
#         process_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
#         self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)
#
#         scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
#         scrollbar.grid(row=1, column=1, sticky=tk.NS)
#         self.process_listbox.config(yscrollcommand=scrollbar.set)
#
#         btn_frame = ttk.Frame(process_frame)
#         btn_frame.grid(row=1, column=2, padx=10)
#         ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)
#
#         ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
#         self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)
#
#         # ---------------------- 新增1：统计信息面板 ----------------------
#         stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="10")
#         stats_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 创建统计数据表格（Treeview控件）
#         columns = ("proc", "max", "min", "avg", "3sigma")
#         self.stats_tree = ttk.Treeview(stats_frame, columns=columns, show="headings", height=5)
#
#         # 设置表头
#         self.stats_tree.heading("proc", text="进程名")
#         self.stats_tree.heading("max", text="最大值 (MB)")
#         self.stats_tree.heading("min", text="最小值 (MB)")
#         self.stats_tree.heading("avg", text="平均值 (MB)")
#         self.stats_tree.heading("3sigma", text="3σ值 (MB)")
#
#         # 设置列宽和对齐方式
#         self.stats_tree.column("proc", width=150, anchor="w")
#         self.stats_tree.column("max", width=100, anchor="e")
#         self.stats_tree.column("min", width=100, anchor="e")
#         self.stats_tree.column("avg", width=100, anchor="e")
#         self.stats_tree.column("3sigma", width=100, anchor="e")
#
#         self.stats_tree.pack(fill=tk.X)
#
#         # 2. 参数设置区域（保持不变）
#         param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
#         param_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(param_frame, text="监控时长:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
#         self.duration_var = tk.StringVar(value="60")
#         ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
#         self.duration_unit = ttk.Combobox(param_frame, values=["秒", "分钟", "小时"], width=6, state="readonly")
#         self.duration_unit.current(0)
#         self.duration_unit.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
#
#         ttk.Label(param_frame, text="采样间隔:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
#         self.interval_var = tk.StringVar(value="5")
#         ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=4, sticky=tk.W, pady=5)
#         self.interval_unit = ttk.Combobox(param_frame, values=["秒", "分钟"], width=6, state="readonly")
#         self.interval_unit.current(0)
#         self.interval_unit.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
#
#         ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
#         self.path_var = tk.StringVar(value=os.getcwd())
#         ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=7, sticky=tk.W, pady=5)
#         ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=8, padx=5, pady=5)
#
#         # 3. 图表合并配置（保持不变）
#         merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
#         merge_frame.pack(fill=tk.X, pady=(0, 10))
#
#         self.merge_var = tk.BooleanVar(value=False)
#         merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
#                                       command=self._toggle_merge_options)
#         merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
#
#         ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
#         self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)
#
#         merge_btn_frame = ttk.Frame(merge_frame)
#         merge_btn_frame.grid(row=2, column=1, padx=10)
#         self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge,
#                                            state=tk.DISABLED)
#         self.add_to_merge_btn.pack(fill=tk.X, pady=5)
#         self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge,
#                                                 state=tk.DISABLED)
#         self.remove_from_merge_btn.pack(fill=tk.X, pady=5)
#
#         ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
#         self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)
#
#         # 4. 控制按钮区域（保持不变）
#         control_frame = ttk.Frame(main_frame)
#         control_frame.pack(fill=tk.X, pady=(0, 10))
#         self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
#         self.start_btn.pack(side=tk.LEFT, padx=5)
#         self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
#         self.stop_btn.pack(side=tk.LEFT, padx=5)
#         self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
#         self.report_btn.pack(side=tk.LEFT, padx=5)
#         self.status_var = tk.StringVar(value="就绪")
#         ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
#
#         # 5. 实时图表区域（保持不变，但扩大尺寸以容纳统计文本）
#         chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         chart_frame.pack(fill=tk.BOTH, expand=True)
#
#         self.fig, self.ax = plt.subplots(figsize=(10, 6))  # 扩大图表尺寸
#         self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#
#         # 初始化悬停标注框（保持不变）
#         self.annotation = self.ax.annotate(
#             "",
#             xy=(0, 0),
#             xytext=(15, 15),
#             textcoords="offset points",
#             bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="gray", alpha=0.9),
#             arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0.2")
#         )
#         self.annotation.set_visible(False)
#         self.canvas.mpl_connect("motion_notify_event", self._on_mouse_hover)
#
#         self._refresh_processes()
#
# # ---------------------- 原有方法保持不变（省略） ----------------------
# # _toggle_merge_options, _sync_merge_source_list, _add_to_merge, _remove_from_merge
# # _refresh_processes, _add_monitor, _remove_monitor, _browse_path
# # _start_monitoring, _stop_monitoring, _monitor_processes
#
#     def _toggle_merge_options(self):
#         state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
#         self.merge_source_listbox.config(state=state)
#         self.merge_target_listbox.config(state=state)
#         self.add_to_merge_btn.config(state=state)
#         self.remove_from_merge_btn.config(state=state)
#         if state == tk.NORMAL:
#             self._sync_merge_source_list()
#
#     def _sync_merge_source_list(self):
#         self.merge_source_listbox.delete(0, tk.END)
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.merge_source_listbox.insert(tk.END, proc_name)
#
#     def _add_to_merge(self):
#         selected_indices = self.merge_source_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.merge_source_listbox.get(i)
#             if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
#                 self.merge_target_listbox.insert(tk.END, proc_name)
#
#     def _remove_from_merge(self):
#         selected_indices = self.merge_target_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.merge_target_listbox.delete(i)
#
#     # ---------------------- 进程管理相关方法（保持不变） ----------------------
#     def _refresh_processes(self):
#         self.process_listbox.delete(0, tk.END)
#         processes = set()
#         for proc in psutil.process_iter(['name']):
#             try:
#                 processes.add(proc.info['name'])
#             except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                 continue
#         for proc_name in sorted(processes):
#             self.process_listbox.insert(tk.END, proc_name)
#
#     def _add_monitor(self):
#         selected_indices = self.process_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.process_listbox.get(i)
#             if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
#                 self.monitor_listbox.insert(tk.END, proc_name)
#                 if self.merge_var.get():
#                     self._sync_merge_source_list()
#
#     def _remove_monitor(self):
#         selected_indices = self.monitor_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.monitor_listbox.delete(i)
#             if self.merge_var.get():
#                 self._sync_merge_source_list()
#
#     def _browse_path(self):
#         path = filedialog.askdirectory()
#         if path:
#             self.path_var.set(path)
#             self.save_path = path
#
#     # ---------------------- 监控控制方法（保持不变） ----------------------
#     def _start_monitoring(self):
#         if self.monitor_listbox.size() == 0:
#             messagebox.showwarning("警告", "请至少选择一个进程进行监控")
#             return
#
#         try:
#             duration_value = int(self.duration_var.get())
#             duration_unit = self.duration_unit.get()
#             if duration_unit == "分钟":
#                 duration = duration_value * 60
#             elif duration_unit == "小时":
#                 duration = duration_value * 3600
#             else:
#                 duration = duration_value
#
#             interval_value = int(self.interval_var.get())
#             interval_unit = self.interval_unit.get()
#             if interval_unit == "分钟":
#                 interval = interval_value * 60
#             else:
#                 interval = interval_value
#
#             if duration <= 0 or interval <= 0 or interval > duration:
#                 raise ValueError
#         except ValueError:
#             messagebox.showwarning("警告", "请输入有效的监控参数（正整数，且间隔不大于时长）")
#             return
#
#         self.process_data = {}
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.process_data[proc_name] = []
#
#         self.status_var.set("监控中...")
#         self.start_btn.config(state=tk.DISABLED)
#         self.stop_btn.config(state=tk.NORMAL)
#         self.report_btn.config(state=tk.DISABLED)
#         self.monitoring = True
#
#         self.monitor_thread = threading.Thread(
#             target=self._monitor_processes,
#             args=(duration, interval),
#             daemon=True
#         )
#         self.monitor_thread.start()
#
#     def _stop_monitoring(self):
#         self.monitoring = False
#         self.status_var.set("监控已停止，准备生成报告")
#         self.start_btn.config(state=tk.NORMAL)
#         self.stop_btn.config(state=tk.DISABLED)
#         self.report_btn.config(state=tk.NORMAL)
#
#     def _monitor_processes(self, duration, interval):
#         end_time = time.time() + duration
#         while self.monitoring and time.time() < end_time:
#             timestamp = datetime.now()
#             for proc_name in self.process_data.keys():
#                 mem_usage = 0
#                 for proc in psutil.process_iter(['name', 'memory_info']):
#                     try:
#                         if proc.info['name'] == proc_name:
#                             mem_usage += proc.info['memory_info'].rss
#                     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                         continue
#                 self.process_data[proc_name].append({'Timestamp': timestamp, 'Memory_Bytes': mem_usage})
#             self.root.after(0, self._update_chart)
#             time.sleep(interval)
#         if self.monitoring:
#             self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
#             self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
#             self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
#             self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
#             self.monitoring = False
#
#
#     # ---------------------- 新增2：更新统计信息表格 ----------------------
#     def _update_stats_table(self, stats_data):
#         """更新UI中的统计信息表格"""
#         # 清空现有数据
#         for item in self.stats_tree.get_children():
#             self.stats_tree.delete(item)
#
#         # 添加新统计数据
#         for data in stats_data:
#             self.stats_tree.insert("", "end", values=data)
#
#     # ---------------------- 新增3：在图表中添加统计文本 ----------------------
#     def _add_chart_stats_annotation(self, stats_data):
#         """在实时图表中添加统计信息文本框"""
#         # 构建统计文本内容
#         stats_text = "统计信息:\n"
#         for proc, max_val, min_val, avg_val, sigma3_val in stats_data:
#             stats_text += (
#                 f"{proc}:\n"
#                 f"  最大值: {max_val:.2f} MB\n"
#                 f"  最小值: {min_val:.2f} MB\n"
#                 f"  平均值: {avg_val:.2f} MB\n"
#                 f"  3σ值:   {sigma3_val:.2f} MB\n\n"
#             )
#
#         # 在图表右上角添加文本框
#         self.ax.text(
#             0.02, 0.98,  # 相对坐标（左上角）
#             stats_text,
#             transform=self.ax.transAxes,
#             verticalalignment='top',
#             bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="gray", alpha=0.8)
#         )
#
#     # ---------------------- 修改1：更新图表时计算并显示统计信息 ----------------------
#     def _update_chart(self):
#         """更新实时图表（新增统计信息计算与显示）"""
#         self.ax.clear()
#         has_data = False
#         stats_data = []  # 存储所有进程的统计数据 [(进程名, 最大值, 最小值, 平均值, 3σ值), ...]
#
#         # 遍历每个进程计算统计值
#         for proc_name, data in self.process_data.items():
#             if data:
#                 has_data = True
#                 df = pd.DataFrame(data)
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)  # 转换为MB
#
#                 # 绘制折线图（保持不变）
#                 self.ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)
#
#                 # ---------------------- 核心：计算统计值 ----------------------
#                 memory_values = df['Memory_MB'].values
#                 max_val = np.max(memory_values)
#                 min_val = np.min(memory_values)
#                 avg_val = np.mean(memory_values)
#                 std_val = np.std(memory_values) if len(memory_values) >= 2 else 0  # 标准差（至少2个数据点）
#                 sigma3_val = 3 * std_val  # 3σ值
#
#                 # 保留两位小数
#                 stats_data.append((
#                     proc_name,
#                     round(max_val, 2),
#                     round(min_val, 2),
#                     round(avg_val, 2),
#                     round(sigma3_val, 2)
#                 ))
#
#         # 更新UI统计表格和图表统计文本
#         if stats_data:
#             self._update_stats_table(stats_data)  # 更新UI表格
#             self._add_chart_stats_annotation(stats_data)  # 更新图表文本
#
#         # 图表样式设置（保持不变）
#         if has_data:
#             self.ax.set_title('实时内存使用监控（含统计信息）')
#             self.ax.set_xlabel('时间')
#             self.ax.set_ylabel('内存使用 (MB)')
#             self.ax.legend(loc='upper left')  # 图例位置调整，避免遮挡统计文本
#
#             # x轴优化（保持不变）
#             max_ticks = min(10, len(df))
#             self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#             self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#             plt.xticks(rotation=45, ha='right')
#             self.fig.tight_layout()
#
#         self.canvas.draw()
#
#     # ---------------------- 鼠标悬停事件（保持不变） ----------------------
#     def _on_mouse_hover(self, event):
#         # 原有鼠标悬停逻辑保持不变
#         if event.inaxes != self.ax:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#         if not self.ax.lines:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#         mouse_x = event.xdata
#         mouse_y = event.ydata
#         if mouse_x is None or mouse_y is None:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#         mouse_datetime = mdates.num2date(mouse_x)
#         min_dist = float('inf')
#         closest_point = None
#         for line in self.ax.lines:
#             proc_name = line.get_label()
#             x_data = line.get_xdata()
#             y_data = line.get_ydata()
#             for i in range(len(x_data)):
#                 point_datetime = mdates.num2date(x_data[i])
#                 time_diff = abs((mouse_datetime - point_datetime).total_seconds())
#                 memory_diff = abs(mouse_y - y_data[i])
#                 if time_diff < 5 and memory_diff < 10:
#                     distance = time_diff * 0.1 + memory_diff * 0.01
#                     if distance < min_dist:
#                         min_dist = distance
#                         closest_point = {
#                             "proc": proc_name,
#                             "time": point_datetime.strftime("%Y-%m-%d %H:%M:%S"),
#                             "memory": round(y_data[i], 2)
#                         }
#         if closest_point:
#             self.annotation.xy = (mouse_x, mouse_y)
#             self.annotation.set_text(
#                 f"进程: {closest_point['proc']}\n"
#                 f"时间: {closest_point['time']}\n"
#                 f"内存: {closest_point['memory']} MB"
#             )
#             self.annotation.set_visible(True)
#         else:
#             self.annotation.set_visible(False)
#         self.canvas.draw_idle()
#
#     # ---------------------- 报告生成方法（保持不变） ----------------------
#     def _generate_report(self):
#         # 原有报告生成逻辑保持不变（可按需添加统计信息到报告）
#         if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
#             messagebox.showwarning("警告", "没有监控数据可生成报告")
#             return
#
#         excel_path = os.path.join(self.save_path, "memory_summary_report.xlsx")
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "内存监控汇总"
#
#         # ...（省略原有报告生成代码）...
#
#
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = MemoryMonitorApp(root)
#     root.mainloop()



# #解决图表节点悬停显示信息功能
# import psutil
# import pandas as pd
# import time
# import threading
# import os
# import matplotlib.pyplot as plt
# from datetime import datetime
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Alignment
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# import matplotlib
# from io import BytesIO
# import matplotlib.dates as mdates  # 新增：导入日期处理模块
#
# matplotlib.use('TkAgg')
# plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
#
#
# class MemoryMonitorApp:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("多进程内存监控工具")
#         self.root.geometry("900x700")
#         self.root.resizable(True, True)
#
#         self.monitoring = False
#         self.monitor_thread = None
#         self.process_data = {}
#         self.selected_processes = set()
#         self.merge_processes = set()
#         self.save_path = os.getcwd()
#
#         # ---------------------- 新增：标注框和鼠标事件相关变量 ----------------------
#         self.annotation = None  # 悬停标注框
#         self.fig = None         # 图表对象（提升为实例变量，确保全局可访问）
#         self.ax = None          # 坐标轴对象（提升为实例变量）
#
#         self._create_widgets()
#
#     def _create_widgets(self):
#         main_frame = ttk.Frame(self.root, padding="10")
#         main_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 1. 进程选择区域（保持不变）
#         process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
#         process_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
#         self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)
#
#         scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
#         scrollbar.grid(row=1, column=1, sticky=tk.NS)
#         self.process_listbox.config(yscrollcommand=scrollbar.set)
#
#         btn_frame = ttk.Frame(process_frame)
#         btn_frame.grid(row=1, column=2, padx=10)
#         ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)
#
#         ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
#         self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)
#
#         # 2. 参数设置区域（保持不变）
#         param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
#         param_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(param_frame, text="监控时长:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
#         self.duration_var = tk.StringVar(value="60")
#         ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
#         self.duration_unit = ttk.Combobox(param_frame, values=["秒", "分钟", "小时"], width=6, state="readonly")
#         self.duration_unit.current(0)
#         self.duration_unit.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
#
#         ttk.Label(param_frame, text="采样间隔:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
#         self.interval_var = tk.StringVar(value="5")
#         ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=4, sticky=tk.W, pady=5)
#         self.interval_unit = ttk.Combobox(param_frame, values=["秒", "分钟"], width=6, state="readonly")
#         self.interval_unit.current(0)
#         self.interval_unit.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
#
#         ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
#         self.path_var = tk.StringVar(value=os.getcwd())
#         ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=7, sticky=tk.W, pady=5)
#         ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=8, padx=5, pady=5)
#
#         # 3. 图表合并配置（保持不变）
#         merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
#         merge_frame.pack(fill=tk.X, pady=(0, 10))
#
#         self.merge_var = tk.BooleanVar(value=False)
#         merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
#                                       command=self._toggle_merge_options)
#         merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
#
#         ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
#         self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)
#
#         merge_btn_frame = ttk.Frame(merge_frame)
#         merge_btn_frame.grid(row=2, column=1, padx=10)
#         self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge, state=tk.DISABLED)
#         self.add_to_merge_btn.pack(fill=tk.X, pady=5)
#         self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge, state=tk.DISABLED)
#         self.remove_from_merge_btn.pack(fill=tk.X, pady=5)
#
#         ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
#         self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)
#
#         # 4. 控制按钮区域（保持不变）
#         control_frame = ttk.Frame(main_frame)
#         control_frame.pack(fill=tk.X, pady=(0, 10))
#         self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
#         self.start_btn.pack(side=tk.LEFT, padx=5)
#         self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
#         self.stop_btn.pack(side=tk.LEFT, padx=5)
#         self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
#         self.report_btn.pack(side=tk.LEFT, padx=5)
#         self.status_var = tk.StringVar(value="就绪")
#         ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
#
#         # 5. 实时图表区域（重点修复：标注框初始化和事件绑定）
#         chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         chart_frame.pack(fill=tk.BOTH, expand=True)
#
#         # ---------------------- 修复1：图表和标注框初始化 ----------------------
#         self.fig, self.ax = plt.subplots(figsize=(8, 4))
#         self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#
#         # 初始化悬停标注框
#         self.annotation = self.ax.annotate(
#             "",
#             xy=(0, 0),
#             xytext=(15, 15),  # 文本偏移量（右下方）
#             textcoords="offset points",
#             bbox=dict(boxstyle="round,pad=0.5", fc="white", ec="gray", alpha=0.9),  # 白色背景框
#             arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0.2")  # 箭头样式
#         )
#         self.annotation.set_visible(False)  # 默认隐藏
#
#         # ---------------------- 修复2：绑定鼠标移动事件 ----------------------
#         self.canvas.mpl_connect("motion_notify_event", self._on_mouse_hover)
#
#         self._refresh_processes()
#
#     # ---------------------- 图表合并相关方法（保持不变） ----------------------
#     def _toggle_merge_options(self):
#         state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
#         self.merge_source_listbox.config(state=state)
#         self.merge_target_listbox.config(state=state)
#         self.add_to_merge_btn.config(state=state)
#         self.remove_from_merge_btn.config(state=state)
#         if state == tk.NORMAL:
#             self._sync_merge_source_list()
#
#     def _sync_merge_source_list(self):
#         self.merge_source_listbox.delete(0, tk.END)
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.merge_source_listbox.insert(tk.END, proc_name)
#
#     def _add_to_merge(self):
#         selected_indices = self.merge_source_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.merge_source_listbox.get(i)
#             if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
#                 self.merge_target_listbox.insert(tk.END, proc_name)
#
#     def _remove_from_merge(self):
#         selected_indices = self.merge_target_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.merge_target_listbox.delete(i)
#
#     # ---------------------- 进程管理相关方法（保持不变） ----------------------
#     def _refresh_processes(self):
#         self.process_listbox.delete(0, tk.END)
#         processes = set()
#         for proc in psutil.process_iter(['name']):
#             try:
#                 processes.add(proc.info['name'])
#             except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                 continue
#         for proc_name in sorted(processes):
#             self.process_listbox.insert(tk.END, proc_name)
#
#     def _add_monitor(self):
#         selected_indices = self.process_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.process_listbox.get(i)
#             if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
#                 self.monitor_listbox.insert(tk.END, proc_name)
#                 if self.merge_var.get():
#                     self._sync_merge_source_list()
#
#     def _remove_monitor(self):
#         selected_indices = self.monitor_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.monitor_listbox.delete(i)
#             if self.merge_var.get():
#                 self._sync_merge_source_list()
#
#     def _browse_path(self):
#         path = filedialog.askdirectory()
#         if path:
#             self.path_var.set(path)
#             self.save_path = path
#
#     # ---------------------- 监控控制方法（保持不变） ----------------------
#     def _start_monitoring(self):
#         if self.monitor_listbox.size() == 0:
#             messagebox.showwarning("警告", "请至少选择一个进程进行监控")
#             return
#
#         try:
#             duration_value = int(self.duration_var.get())
#             duration_unit = self.duration_unit.get()
#             if duration_unit == "分钟":
#                 duration = duration_value * 60
#             elif duration_unit == "小时":
#                 duration = duration_value * 3600
#             else:
#                 duration = duration_value
#
#             interval_value = int(self.interval_var.get())
#             interval_unit = self.interval_unit.get()
#             if interval_unit == "分钟":
#                 interval = interval_value * 60
#             else:
#                 interval = interval_value
#
#             if duration <= 0 or interval <= 0 or interval > duration:
#                 raise ValueError
#         except ValueError:
#             messagebox.showwarning("警告", "请输入有效的监控参数（正整数，且间隔不大于时长）")
#             return
#
#         self.process_data = {}
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.process_data[proc_name] = []
#
#         self.status_var.set("监控中...")
#         self.start_btn.config(state=tk.DISABLED)
#         self.stop_btn.config(state=tk.NORMAL)
#         self.report_btn.config(state=tk.DISABLED)
#         self.monitoring = True
#
#         self.monitor_thread = threading.Thread(
#             target=self._monitor_processes,
#             args=(duration, interval),
#             daemon=True
#         )
#         self.monitor_thread.start()
#
#     def _stop_monitoring(self):
#         self.monitoring = False
#         self.status_var.set("监控已停止，准备生成报告")
#         self.start_btn.config(state=tk.NORMAL)
#         self.stop_btn.config(state=tk.DISABLED)
#         self.report_btn.config(state=tk.NORMAL)
#
#     def _monitor_processes(self, duration, interval):
#         end_time = time.time() + duration
#         while self.monitoring and time.time() < end_time:
#             timestamp = datetime.now()
#             for proc_name in self.process_data.keys():
#                 mem_usage = 0
#                 for proc in psutil.process_iter(['name', 'memory_info']):
#                     try:
#                         if proc.info['name'] == proc_name:
#                             mem_usage += proc.info['memory_info'].rss
#                     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                         continue
#                 self.process_data[proc_name].append({'Timestamp': timestamp, 'Memory_Bytes': mem_usage})
#             self.root.after(0, self._update_chart)
#             time.sleep(interval)
#         if self.monitoring:
#             self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
#             self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
#             self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
#             self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
#             self.monitoring = False
#
#     # ---------------------- 实时图表更新（保持不变，但确保线条标签正确） ----------------------
#     def _update_chart(self):
#         self.ax.clear()
#         has_data = False
#
#         for proc_name, data in self.process_data.items():
#             if data:
#                 has_data = True
#                 df = pd.DataFrame(data)
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 # 确保线条标签为进程名（用于悬停识别）
#                 self.ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)
#
#         if has_data:
#             self.ax.set_title('实时内存使用监控')
#             self.ax.set_xlabel('时间')
#             self.ax.set_ylabel('内存使用 (MB)')
#             self.ax.legend()
#             max_ticks = min(10, len(df))
#             self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#             self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#             plt.xticks(rotation=45, ha='right')
#             self.fig.tight_layout()
#
#         self.canvas.draw()
#
#     # ---------------------- 修复3：鼠标悬停事件处理（核心逻辑） ----------------------
#     def _on_mouse_hover(self, event):
#         """鼠标悬停事件处理：显示数据点详情"""
#         # 1. 检查鼠标是否在图表区域内
#         if event.inaxes != self.ax:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#
#         # 2. 检查是否有线条数据（无监控数据时不显示）
#         if not self.ax.lines:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#
#         # 3. 获取鼠标坐标并转换为datetime
#         mouse_x = event.xdata
#         mouse_y = event.ydata
#         if mouse_x is None or mouse_y is None:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()
#             return
#
#         # 4. 转换鼠标x坐标为datetime对象
#         mouse_datetime = mdates.num2date(mouse_x)
#
#         # 5. 遍历所有线条，查找最近的数据点
#         min_dist = float('inf')
#         closest_point = None
#
#         for line in self.ax.lines:
#             proc_name = line.get_label()  # 获取进程名（线条标签）
#             x_data = line.get_xdata()     # x轴数据（Matplotlib日期数值）
#             y_data = line.get_ydata()     # y轴数据（内存MB）
#
#             # 遍历当前线条的所有数据点
#             for i in range(len(x_data)):
#                 # 转换数据点x坐标为datetime
#                 point_datetime = mdates.num2date(x_data[i])
#                 # 计算时间差（秒）和内存差（MB）
#                 time_diff = abs((mouse_datetime - point_datetime).total_seconds())
#                 memory_diff = abs(mouse_y - y_data[i])
#
#                 # ---------------------- 修复4：调整匹配阈值（放宽条件） ----------------------
#                 # 时间差<5秒且内存差<10MB，视为有效匹配
#                 if time_diff < 5 and memory_diff < 10:
#                     # 计算综合距离（时间差权重更高）
#                     distance = time_diff * 0.1 + memory_diff * 0.01
#                     if distance < min_dist:
#                         min_dist = distance
#                         closest_point = {
#                             "proc": proc_name,
#                             "time": point_datetime.strftime("%Y-%m-%d %H:%M:%S"),
#                             "memory": round(y_data[i], 2)
#                         }
#
#         # 6. 更新标注框内容和可见性
#         if closest_point:
#             self.annotation.xy = (mouse_x, mouse_y)  # 标注框指向鼠标位置
#             self.annotation.set_text(
#                 f"进程: {closest_point['proc']}\n"
#                 f"时间: {closest_point['time']}\n"
#                 f"内存: {closest_point['memory']} MB"
#             )
#             self.annotation.set_visible(True)
#         else:
#             self.annotation.set_visible(False)
#
#         # 7. 刷新图表
#         self.canvas.draw_idle()
#
#     # ---------------------- 报告生成方法（保持不变） ----------------------
#     def _generate_report(self):
#         if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
#             messagebox.showwarning("警告", "没有监控数据可生成报告")
#             return
#
#         excel_path = os.path.join(self.save_path, "memory_summary_report.xlsx")
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "内存监控汇总"
#
#         summary_data = {}
#         process_names = [self.monitor_listbox.get(i) for i in range(self.monitor_listbox.size())]
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')
#                 for _, row in df.iterrows():
#                     timestamp = row['Timestamp']
#                     if timestamp not in summary_data:
#                         summary_data[timestamp] = {'Timestamp': timestamp, 'Time': row['Time']}
#                     summary_data[timestamp][proc_name] = row['Memory_MB']
#         if not summary_data:
#             messagebox.showwarning("警告", "监控数据为空，无法生成报告")
#             return
#         summary_df = pd.DataFrame.from_dict(summary_data, orient='index').sort_values('Timestamp')
#
#         headers = ['Timestamp'] + process_names
#         ws.append(headers)
#         for _, row in summary_df.iterrows():
#             row_data = [row['Timestamp']] + [row.get(proc, '') for proc in process_names]
#             ws.append(row_data)
#
#         center_alignment = Alignment(horizontal='center', vertical='center')
#         for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
#             for cell in row:
#                 if isinstance(cell.value, (int, float)):
#                     cell.number_format = '0.00'
#                 cell.alignment = center_alignment
#         ws.column_dimensions['A'].width = 25
#         for col in range(2, 2 + len(process_names)):
#             ws.column_dimensions[chr(64 + col)].width = 18
#
#         current_row = len(summary_df) + 3
#         merge_chart_inserted = False
#         if self.merge_var.get() and self.merge_target_listbox.size() > 0:
#             merge_procs = [self.merge_target_listbox.get(i) for i in range(self.merge_target_listbox.size())]
#             if merge_procs:
#                 fig, ax = plt.subplots(figsize=(12, 6))
#                 colors = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
#                 color_idx = 0
#                 for proc_name in merge_procs:
#                     if proc_name in self.process_data and self.process_data[proc_name]:
#                         df = pd.DataFrame(self.process_data[proc_name])
#                         df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                         ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-',
#                                 label=proc_name, color=colors[color_idx % len(colors)])
#                         color_idx += 1
#                 ax.set_title('多进程内存使用对比（合并图表）')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#                 ax.legend()
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 800
#                 img.height = 400
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 30
#                 plt.close(fig)
#                 merge_chart_inserted = True
#
#         if not merge_chart_inserted:
#             current_row = len(summary_df) + 3
#         else:
#             current_row += 5
#
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 fig, ax = plt.subplots(figsize=(10, 4))
#                 ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', color='blue')
#                 ax.set_title(f'{proc_name} 内存使用趋势')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 600
#                 img.height = 300
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 20
#                 plt.close(fig)
#
#         try:
#             wb.save(excel_path)
#             messagebox.showinfo("成功", f"汇总报告已生成：\n{excel_path}")
#             self.status_var.set("报告生成完成")
#         except Exception as e:
#             messagebox.showerror("错误", f"保存报告失败：{str(e)}")
#
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = MemoryMonitorApp(root)
#     root.mainloop()


# #增加鼠标悬停显示节点信息功能
# import psutil
# import pandas as pd
# import time
# import threading
# import os
# import matplotlib.pyplot as plt
# from datetime import datetime
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Alignment
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# import matplotlib
# from io import BytesIO
# # 新增：导入matplotlib日期处理模块
# import matplotlib.dates as mdates
#
# matplotlib.use('TkAgg')
# plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
#
#
# class MemoryMonitorApp:
#     def __init__(self, root):
#         """初始化内存监控应用"""
#         self.root = root
#         self.root.title("多进程内存监控工具")
#         self.root.geometry("900x700")
#         self.root.resizable(True, True)
#
#         # 初始化变量
#         self.monitoring = False
#         self.monitor_thread = None
#         self.process_data = {}  # 存储监控数据 {进程名: DataFrame}
#         self.selected_processes = set()  # 选中的监控进程
#         self.merge_processes = set()  # 选中的合并图表进程
#         self.save_path = os.getcwd()  # 默认保存路径
#
#         # 创建UI界面
#         self._create_widgets()
#
#     def _create_widgets(self):
#         """创建UI组件"""
#         # 创建主框架
#         main_frame = ttk.Frame(self.root, padding="10")
#         main_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 1. 进程选择区域（保持不变）
#         process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
#         process_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 进程列表
#         ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
#         self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)
#
#         # 滚动条
#         scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
#         scrollbar.grid(row=1, column=1, sticky=tk.NS)
#         self.process_listbox.config(yscrollcommand=scrollbar.set)
#
#         # 按钮
#         btn_frame = ttk.Frame(process_frame)
#         btn_frame.grid(row=1, column=2, padx=10)
#
#         ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)
#
#         # 当前监控进程
#         ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
#         self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)
#
#         # 2. 参数设置区域（修改：添加单位下拉框）
#         param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
#         param_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # ---------------------- 修改1：监控时长带单位选择 ----------------------
#         ttk.Label(param_frame, text="监控时长:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
#         self.duration_var = tk.StringVar(value="60")
#         ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
#         # 时长单位下拉框（秒/分钟/小时）
#         self.duration_unit = ttk.Combobox(param_frame, values=["秒", "分钟", "小时"], width=6, state="readonly")
#         self.duration_unit.current(0)  # 默认秒
#         self.duration_unit.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
#
#         # ---------------------- 修改2：采样间隔带单位选择 ----------------------
#         ttk.Label(param_frame, text="采样间隔:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
#         self.interval_var = tk.StringVar(value="5")
#         ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=4, sticky=tk.W, pady=5)
#         # 间隔单位下拉框（秒/分钟）
#         self.interval_unit = ttk.Combobox(param_frame, values=["秒", "分钟"], width=6, state="readonly")
#         self.interval_unit.current(0)  # 默认秒
#         self.interval_unit.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
#
#         # 报告保存路径（调整列号以适应单位下拉框）
#         ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
#         self.path_var = tk.StringVar(value=os.getcwd())
#         ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=7, sticky=tk.W, pady=5)
#         ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=8, padx=5, pady=5)
#
#         # 3. 图表合并配置（保持不变）
#         merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
#         merge_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 启用图表合并复选框
#         self.merge_var = tk.BooleanVar(value=False)
#         merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
#                                       command=self._toggle_merge_options)
#         merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
#
#         # 左侧：当前监控进程
#         ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
#         self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)
#
#         # 中间：添加/移除按钮
#         merge_btn_frame = ttk.Frame(merge_frame)
#         merge_btn_frame.grid(row=2, column=1, padx=10)
#
#         self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge,
#                                            state=tk.DISABLED)
#         self.add_to_merge_btn.pack(fill=tk.X, pady=5)
#
#         self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge,
#                                                 state=tk.DISABLED)
#         self.remove_from_merge_btn.pack(fill=tk.X, pady=5)
#
#         # 右侧：合并监控进程
#         ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
#         self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)
#
#         # 4. 控制按钮区域（保持不变）
#         control_frame = ttk.Frame(main_frame)
#         control_frame.pack(fill=tk.X, pady=(0, 10))
#
#         self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
#         self.start_btn.pack(side=tk.LEFT, padx=5)
#
#         self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
#         self.stop_btn.pack(side=tk.LEFT, padx=5)
#
#         self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
#         self.report_btn.pack(side=tk.LEFT, padx=5)
#
#         self.status_var = tk.StringVar(value="就绪")
#         ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
#
#         # 5. 实时图表区域
#         chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         chart_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 创建matplotlib图表
#         self.fig, self.ax = plt.subplots(figsize=(8, 4))
#         self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#
#         # ---------------------- 添加：初始化悬停提示框 ----------------------
#         self.annotation = self.ax.annotate(
#             "",
#             xy=(0, 0),  # 标注位置（数据坐标系）
#             xytext=(10, 10),  # 文本偏移量
#             textcoords="offset points",  # 文本坐标系统
#             bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="gray", alpha=0.9),  # 白色圆角背景框
#             arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=0.1")  # 箭头样式
#         )
#         self.annotation.set_visible(False)  # 默认隐藏
#
#         # 连接鼠标移动事件
#         self.canvas.mpl_connect("motion_notify_event", self._on_mouse_hover)
#
#         # 初始化刷新进程列表
#         self._refresh_processes()
#         # # 5. 实时图表区域
#         # chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         # chart_frame.pack(fill=tk.BOTH, expand=True)
#         #
#         # # 创建matplotlib图表
#         # self.fig, self.ax = plt.subplots(figsize=(8, 4))
#         # self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         # self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#         #
#         # # 初始化刷新进程列表
#         # self._refresh_processes()
#
#     # ---------------------- 图表合并相关方法（保持不变） ----------------------
#     def _toggle_merge_options(self):
#         state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
#         self.merge_source_listbox.config(state=state)
#         self.merge_target_listbox.config(state=state)
#         self.add_to_merge_btn.config(state=state)
#         self.remove_from_merge_btn.config(state=state)
#         if state == tk.NORMAL:
#             self._sync_merge_source_list()
#
#     def _sync_merge_source_list(self):
#         self.merge_source_listbox.delete(0, tk.END)
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.merge_source_listbox.insert(tk.END, proc_name)
#
#     def _add_to_merge(self):
#         selected_indices = self.merge_source_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.merge_source_listbox.get(i)
#             if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
#                 self.merge_target_listbox.insert(tk.END, proc_name)
#
#     def _remove_from_merge(self):
#         selected_indices = self.merge_target_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.merge_target_listbox.delete(i)
#
#     # ---------------------- 进程管理相关方法（保持不变） ----------------------
#     def _refresh_processes(self):
#         self.process_listbox.delete(0, tk.END)
#         processes = set()
#         for proc in psutil.process_iter(['name']):
#             try:
#                 processes.add(proc.info['name'])
#             except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                 continue
#         for proc_name in sorted(processes):
#             self.process_listbox.insert(tk.END, proc_name)
#
#     def _add_monitor(self):
#         selected_indices = self.process_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.process_listbox.get(i)
#             if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
#                 self.monitor_listbox.insert(tk.END, proc_name)
#                 if self.merge_var.get():
#                     self._sync_merge_source_list()
#
#     def _remove_monitor(self):
#         selected_indices = self.monitor_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.monitor_listbox.delete(i)
#             if self.merge_var.get():
#                 self._sync_merge_source_list()
#
#     def _browse_path(self):
#         path = filedialog.askdirectory()
#         if path:
#             self.path_var.set(path)
#             self.save_path = path
#
#     # ---------------------- 监控控制方法（修改：添加单位转换） ----------------------
#     def _start_monitoring(self):
#         """开始监控进程（添加单位转换逻辑）"""
#         if self.monitor_listbox.size() == 0:
#             messagebox.showwarning("警告", "请至少选择一个进程进行监控")
#             return
#
#         # ---------------------- 修改3：带单位的参数校验与转换 ----------------------
#         try:
#             # 监控时长转换（秒/分钟/小时 → 秒）
#             duration_value = int(self.duration_var.get())
#             duration_unit = self.duration_unit.get()
#             if duration_unit == "分钟":
#                 duration = duration_value * 60
#             elif duration_unit == "小时":
#                 duration = duration_value * 3600
#             else:  # 秒
#                 duration = duration_value
#
#             # 采样间隔转换（秒/分钟 → 秒）
#             interval_value = int(self.interval_var.get())
#             interval_unit = self.interval_unit.get()
#             if interval_unit == "分钟":
#                 interval = interval_value * 60
#             else:  # 秒
#                 interval = interval_value
#
#             # 参数合法性校验
#             if duration <= 0 or interval <= 0 or interval > duration:
#                 raise ValueError
#         except ValueError:
#             messagebox.showwarning("警告", "请输入有效的监控参数（正整数，且间隔不大于时长）")
#             return
#
#         # 初始化数据存储（保持不变）
#         self.process_data = {}
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.process_data[proc_name] = []
#
#         # 更新状态和按钮（保持不变）
#         self.status_var.set("监控中...")
#         self.start_btn.config(state=tk.DISABLED)
#         self.stop_btn.config(state=tk.NORMAL)
#         self.report_btn.config(state=tk.DISABLED)
#         self.monitoring = True
#
#         self.monitor_thread = threading.Thread(
#             target=self._monitor_processes,
#             args=(duration, interval),
#             daemon=True
#         )
#         self.monitor_thread.start()
#
#     def _stop_monitoring(self):
#         """停止监控进程（保持不变）"""
#         self.monitoring = False
#         self.status_var.set("监控已停止，准备生成报告")
#         self.start_btn.config(state=tk.NORMAL)
#         self.stop_btn.config(state=tk.DISABLED)
#         self.report_btn.config(state=tk.NORMAL)
#
#     def _monitor_processes(self, duration, interval):
#         """监控进程内存使用情况（保持不变）"""
#         end_time = time.time() + duration
#         while self.monitoring and time.time() < end_time:
#             timestamp = datetime.now()
#             for proc_name in self.process_data.keys():
#                 mem_usage = 0
#                 for proc in psutil.process_iter(['name', 'memory_info']):
#                     try:
#                         if proc.info['name'] == proc_name:
#                             mem_usage += proc.info['memory_info'].rss
#                     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                         continue
#                 self.process_data[proc_name].append({'Timestamp': timestamp, 'Memory_Bytes': mem_usage})
#             self.root.after(0, self._update_chart)
#             time.sleep(interval)
#         if self.monitoring:
#             self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
#             self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
#             self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
#             self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
#             self.monitoring = False
#
#     def _on_mouse_hover(self, event):
#         """鼠标悬停事件处理函数：显示数据点详情"""
#         # 检查鼠标是否在图表区域内
#         if event.inaxes != self.ax:
#             self.annotation.set_visible(False)
#             self.canvas.draw_idle()  # 刷新图表
#             return
#
#         # 获取鼠标当前坐标（数据坐标系）
#         mouse_x = event.xdata  # x轴坐标（Matplotlib日期数值）
#         mouse_y = event.ydata  # y轴坐标（内存MB）
#
#         # 转换鼠标x坐标为datetime对象（用于时间差计算）
#         from matplotlib.dates import num2date
#         mouse_datetime = num2date(mouse_x)
#
#         # 初始化变量：记录最近的数据点
#         min_dist = float('inf')  # 最小距离（初始设为无穷大）
#         closest_point = None  # 最近点信息
#
#         # 遍历所有绘制的线条（每个进程一条线）
#         for line in self.ax.lines:
#             proc_name = line.get_label()  # 获取进程名（线条标签）
#             x_data = line.get_xdata()  # x轴数据（Matplotlib日期数值数组）
#             y_data = line.get_ydata()  # y轴数据（内存MB数组）
#
#             # 遍历当前线条的所有数据点
#             for i in range(len(x_data)):
#                 # 转换数据点x坐标为datetime对象
#                 point_datetime = num2date(x_data[i])
#
#                 # 计算鼠标与数据点的距离（时间差和内存差的加权距离）
#                 time_diff = abs((mouse_datetime - point_datetime).total_seconds())  # 时间差（秒）
#                 memory_diff = abs(mouse_y - y_data[i])  # 内存差（MB）
#
#                 # 综合距离（时间差权重更高，确保时间接近的点优先匹配）
#                 distance = time_diff * 0.1 + memory_diff * 0.01
#
#                 # 更新最近点（距离阈值可根据需要调整）
#                 if distance < min_dist and time_diff < 600 and memory_diff < 1000:
#                     min_dist = distance
#                     closest_point = {
#                         "proc": proc_name,
#                         "time": point_datetime.strftime("%Y-%m-%d %H:%M:%S"),
#                         "memory": round(y_data[i], 2)
#                     }
#
#         # 更新标注内容并显示/隐藏
#         if closest_point:
#             # 设置标注位置和文本
#             self.annotation.xy = (mouse_x, mouse_y)
#             self.annotation.set_text(
#                 f"进程: {closest_point['proc']}\n"
#                 f"时间: {closest_point['time']}\n"
#                 f"内存: {closest_point['memory']} MB"
#             )
#             self.annotation.set_visible(True)
#         else:
#             self.annotation.set_visible(False)
#
#         # 刷新图表显示
#         self.canvas.draw_idle()
#
#     def _update_chart(self):
#         """更新实时图表（保持原有功能，仅添加注释说明）"""
#         self.ax.clear()
#         has_data = False
#
#         # 绘制每个进程的内存使用曲线（保留label=proc_name用于悬停识别）
#         for proc_name, data in self.process_data.items():
#             if data:
#                 has_data = True
#                 df = pd.DataFrame(data)
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 # 绘制时保留进程名作为线条标签（关键：用于悬停时识别进程）
#                 self.ax.plot(df['Timestamp'], df['Memory_MB'],
#                              marker='o', linestyle='-', label=proc_name)
#
#         if has_data:
#             self.ax.set_title('实时内存使用监控')
#             self.ax.set_xlabel('时间')
#             self.ax.set_ylabel('内存使用 (MB)')
#             self.ax.legend()
#
#             # x轴刻度优化（原有代码保持不变）
#             max_ticks = min(10, len(df))
#             self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#             self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#             plt.xticks(rotation=45, ha='right')
#             self.fig.tight_layout()
#
#         # 确保悬停标注在图表更新后仍能正常显示
#         self.canvas.draw()
#     # ---------------------- 实时图表更新（修改：x轴时间优化） ----------------------
#     # def _update_chart(self):
#     #     """更新实时图表（优化x轴时间显示）"""
#     #     self.ax.clear()
#     #     has_data = False
#     #
#     #     # 绘制每个进程的内存使用曲线
#     #     for proc_name, data in self.process_data.items():
#     #         if data:
#     #             has_data = True
#     #             df = pd.DataFrame(data)
#     #             df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#     #             # ---------------------- 修改4：使用datetime类型x轴 ----------------------
#     #             self.ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)
#     #
#     #     if has_data:
#     #         self.ax.set_title('实时内存使用监控')
#     #         self.ax.set_xlabel('时间')
#     #         self.ax.set_ylabel('内存使用 (MB)')
#     #         self.ax.legend()
#     #
#     #         # ---------------------- 修改5：x轴刻度优化 ----------------------
#     #         # 限制最大刻度数量（最多10个）
#     #         max_ticks = min(10, len(df))
#     #         self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#     #         # 时间格式化（时:分:秒）
#     #         self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#     #         # 标签旋转45度并右对齐
#     #         plt.xticks(rotation=45, ha='right')
#     #         self.fig.tight_layout()
#     #
#     #     self.canvas.draw()
#
#     # ---------------------- 报告生成方法（修改：图表x轴优化） ----------------------
#     def _generate_report(self):
#         """生成汇总Excel报告（优化图表x轴显示）"""
#         if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
#             messagebox.showwarning("警告", "没有监控数据可生成报告")
#             return
#
#         excel_path = os.path.join(self.save_path, "memory_summary_report.xlsx")
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "内存监控汇总"
#
#         # 步骤1：汇总表格数据（保持不变）
#         summary_data = {}
#         process_names = [self.monitor_listbox.get(i) for i in range(self.monitor_listbox.size())]
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')
#                 for _, row in df.iterrows():
#                     timestamp = row['Timestamp']
#                     if timestamp not in summary_data:
#                         summary_data[timestamp] = {'Timestamp': timestamp, 'Time': row['Time']}
#                     summary_data[timestamp][proc_name] = row['Memory_MB']
#         if not summary_data:
#             messagebox.showwarning("警告", "监控数据为空，无法生成报告")
#             return
#         summary_df = pd.DataFrame.from_dict(summary_data, orient='index').sort_values('Timestamp')
#
#         # 写入Excel表格并居中对齐（保持不变）
#         headers = ['Timestamp'] + process_names
#         ws.append(headers)
#         for _, row in summary_df.iterrows():
#             row_data = [row['Timestamp']] + [row.get(proc, '') for proc in process_names]
#             ws.append(row_data)
#         center_alignment = Alignment(horizontal='center', vertical='center')
#         for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
#             for cell in row:
#                 if isinstance(cell.value, (int, float)):
#                     cell.number_format = '0.00'
#                 cell.alignment = center_alignment
#         ws.column_dimensions['A'].width = 25
#         for col in range(2, 2 + len(process_names)):
#             ws.column_dimensions[chr(64 + col)].width = 18
#
#         # 步骤2：生成合并图表（修改x轴优化）
#         current_row = len(summary_df) + 3
#         merge_chart_inserted = False
#         if self.merge_var.get() and self.merge_target_listbox.size() > 0:
#             merge_procs = [self.merge_target_listbox.get(i) for i in range(self.merge_target_listbox.size())]
#             if merge_procs:
#                 fig, ax = plt.subplots(figsize=(12, 6))
#                 colors = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
#                 color_idx = 0
#
#                 for proc_name in merge_procs:
#                     if proc_name in self.process_data and self.process_data[proc_name]:
#                         df = pd.DataFrame(self.process_data[proc_name])
#                         df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                         # ---------------------- 修改6：合并图表x轴使用datetime ----------------------
#                         ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-',
#                                 label=proc_name, color=colors[color_idx % len(colors)])
#                         color_idx += 1
#
#                 ax.set_title('多进程内存使用对比（合并图表）')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#                 ax.legend()
#                 # ---------------------- 修改7：合并图表x轴刻度优化 ----------------------
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#
#                 # 插入图表（保持不变）
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 800
#                 img.height = 400
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 30
#                 plt.close(fig)
#                 merge_chart_inserted = True
#
#         # 步骤3：生成单个进程图表（修改x轴优化）
#         if not merge_chart_inserted:
#             current_row = len(summary_df) + 3
#         else:
#             current_row += 5
#
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 fig, ax = plt.subplots(figsize=(10, 4))
#                 # ---------------------- 修改8：单个图表x轴使用datetime ----------------------
#                 ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', color='blue')
#                 ax.set_title(f'{proc_name} 内存使用趋势')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#
#                 # ---------------------- 修改9：单个图表x轴刻度优化 ----------------------
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#
#                 # 插入图表（保持不变）
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 600
#                 img.height = 300
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 20
#                 plt.close(fig)
#
#         # 保存报告（保持不变）
#         try:
#             wb.save(excel_path)
#             messagebox.showinfo("成功", f"汇总报告已生成：\n{excel_path}")
#             self.status_var.set("报告生成完成")
#         except Exception as e:
#             messagebox.showerror("错误", f"保存报告失败：{str(e)}")
#
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = MemoryMonitorApp(root)
#     root.mainloop()


# 一、监控参数模块：添加单位下拉选择列表
# 1. UI布局修改（添加下拉选择框）
# 在原有监控时长和采样间隔输入框右侧添加单位选择下拉列表（Combobox），支持秒/分钟/小时单位切换。
#2. 时间单位转换逻辑（核心功能实现）
#在监控启动前，将用户输入的数值与单位组合转换为秒级整数，确保后端监控逻辑统一以秒为单位处理。

# 二、折线图x轴时间显示优化（解决重合模糊）
# 通过动态调整x轴刻度数量、优化时间格式化和标签旋转角度，解决数据量增大时的标签重叠问题。
#
# 1. 图表x轴优化核心代码
# 在原有图表生成逻辑中添加刻度限制和智能格式化：
#
# 2. 单个进程图表同步优化
# 对单个进程图表应用相同的x轴优化逻辑：

# import psutil
# import pandas as pd
# import time
# import threading
# import os
# import matplotlib.pyplot as plt
# from datetime import datetime
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.styles import Alignment
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# import matplotlib
# from io import BytesIO
# # 新增：导入matplotlib日期处理模块
# import matplotlib.dates as mdates
#
# matplotlib.use('TkAgg')
# plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
#
#
# class MemoryMonitorApp:
#     def __init__(self, root):
#         """初始化内存监控应用"""
#         self.root = root
#         self.root.title("多进程内存监控工具")
#         self.root.geometry("900x700")
#         self.root.resizable(True, True)
#
#         # 初始化变量
#         self.monitoring = False
#         self.monitor_thread = None
#         self.process_data = {}  # 存储监控数据 {进程名: DataFrame}
#         self.selected_processes = set()  # 选中的监控进程
#         self.merge_processes = set()  # 选中的合并图表进程
#         self.save_path = os.getcwd()  # 默认保存路径
#
#         # 创建UI界面
#         self._create_widgets()
#
#     def _create_widgets(self):
#         """创建UI组件"""
#         # 创建主框架
#         main_frame = ttk.Frame(self.root, padding="10")
#         main_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 1. 进程选择区域（保持不变）
#         process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
#         process_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 进程列表
#         ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
#         self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)
#
#         # 滚动条
#         scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
#         scrollbar.grid(row=1, column=1, sticky=tk.NS)
#         self.process_listbox.config(yscrollcommand=scrollbar.set)
#
#         # 按钮
#         btn_frame = ttk.Frame(process_frame)
#         btn_frame.grid(row=1, column=2, padx=10)
#
#         ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)
#
#         # 当前监控进程
#         ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
#         self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)
#
#         # 2. 参数设置区域（修改：添加单位下拉框）
#         param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
#         param_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # ---------------------- 修改1：监控时长带单位选择 ----------------------
#         ttk.Label(param_frame, text="监控时长:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
#         self.duration_var = tk.StringVar(value="60")
#         ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
#         # 时长单位下拉框（秒/分钟/小时）
#         self.duration_unit = ttk.Combobox(param_frame, values=["秒", "分钟", "小时"], width=6, state="readonly")
#         self.duration_unit.current(0)  # 默认秒
#         self.duration_unit.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
#
#         # ---------------------- 修改2：采样间隔带单位选择 ----------------------
#         ttk.Label(param_frame, text="采样间隔:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
#         self.interval_var = tk.StringVar(value="5")
#         ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=4, sticky=tk.W, pady=5)
#         # 间隔单位下拉框（秒/分钟）
#         self.interval_unit = ttk.Combobox(param_frame, values=["秒", "分钟"], width=6, state="readonly")
#         self.interval_unit.current(0)  # 默认秒
#         self.interval_unit.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)
#
#         # 报告保存路径（调整列号以适应单位下拉框）
#         ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=6, sticky=tk.W, padx=5, pady=5)
#         self.path_var = tk.StringVar(value=os.getcwd())
#         ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=7, sticky=tk.W, pady=5)
#         ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=8, padx=5, pady=5)
#
#         # 3. 图表合并配置（保持不变）
#         merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
#         merge_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 启用图表合并复选框
#         self.merge_var = tk.BooleanVar(value=False)
#         merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
#                                       command=self._toggle_merge_options)
#         merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
#
#         # 左侧：当前监控进程
#         ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
#         self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)
#
#         # 中间：添加/移除按钮
#         merge_btn_frame = ttk.Frame(merge_frame)
#         merge_btn_frame.grid(row=2, column=1, padx=10)
#
#         self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge,
#                                            state=tk.DISABLED)
#         self.add_to_merge_btn.pack(fill=tk.X, pady=5)
#
#         self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge,
#                                                 state=tk.DISABLED)
#         self.remove_from_merge_btn.pack(fill=tk.X, pady=5)
#
#         # 右侧：合并监控进程
#         ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
#         self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)
#
#         # 4. 控制按钮区域（保持不变）
#         control_frame = ttk.Frame(main_frame)
#         control_frame.pack(fill=tk.X, pady=(0, 10))
#
#         self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
#         self.start_btn.pack(side=tk.LEFT, padx=5)
#
#         self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
#         self.stop_btn.pack(side=tk.LEFT, padx=5)
#
#         self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
#         self.report_btn.pack(side=tk.LEFT, padx=5)
#
#         self.status_var = tk.StringVar(value="就绪")
#         ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
#
#         # 5. 实时图表区域
#         chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         chart_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 创建matplotlib图表
#         self.fig, self.ax = plt.subplots(figsize=(8, 4))
#         self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#
#         # 初始化刷新进程列表
#         self._refresh_processes()
#
#     # ---------------------- 图表合并相关方法（保持不变） ----------------------
#     def _toggle_merge_options(self):
#         state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
#         self.merge_source_listbox.config(state=state)
#         self.merge_target_listbox.config(state=state)
#         self.add_to_merge_btn.config(state=state)
#         self.remove_from_merge_btn.config(state=state)
#         if state == tk.NORMAL:
#             self._sync_merge_source_list()
#
#     def _sync_merge_source_list(self):
#         self.merge_source_listbox.delete(0, tk.END)
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.merge_source_listbox.insert(tk.END, proc_name)
#
#     def _add_to_merge(self):
#         selected_indices = self.merge_source_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.merge_source_listbox.get(i)
#             if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
#                 self.merge_target_listbox.insert(tk.END, proc_name)
#
#     def _remove_from_merge(self):
#         selected_indices = self.merge_target_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.merge_target_listbox.delete(i)
#
#     # ---------------------- 进程管理相关方法（保持不变） ----------------------
#     def _refresh_processes(self):
#         self.process_listbox.delete(0, tk.END)
#         processes = set()
#         for proc in psutil.process_iter(['name']):
#             try:
#                 processes.add(proc.info['name'])
#             except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                 continue
#         for proc_name in sorted(processes):
#             self.process_listbox.insert(tk.END, proc_name)
#
#     def _add_monitor(self):
#         selected_indices = self.process_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.process_listbox.get(i)
#             if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
#                 self.monitor_listbox.insert(tk.END, proc_name)
#                 if self.merge_var.get():
#                     self._sync_merge_source_list()
#
#     def _remove_monitor(self):
#         selected_indices = self.monitor_listbox.curselection()
#         for i in sorted(selected_indices, reverse=True):
#             self.monitor_listbox.delete(i)
#             if self.merge_var.get():
#                 self._sync_merge_source_list()
#
#     def _browse_path(self):
#         path = filedialog.askdirectory()
#         if path:
#             self.path_var.set(path)
#             self.save_path = path
#
#     # ---------------------- 监控控制方法（修改：添加单位转换） ----------------------
#     def _start_monitoring(self):
#         """开始监控进程（添加单位转换逻辑）"""
#         if self.monitor_listbox.size() == 0:
#             messagebox.showwarning("警告", "请至少选择一个进程进行监控")
#             return
#
#         # ---------------------- 修改3：带单位的参数校验与转换 ----------------------
#         try:
#             # 监控时长转换（秒/分钟/小时 → 秒）
#             duration_value = int(self.duration_var.get())
#             duration_unit = self.duration_unit.get()
#             if duration_unit == "分钟":
#                 duration = duration_value * 60
#             elif duration_unit == "小时":
#                 duration = duration_value * 3600
#             else:  # 秒
#                 duration = duration_value
#
#             # 采样间隔转换（秒/分钟 → 秒）
#             interval_value = int(self.interval_var.get())
#             interval_unit = self.interval_unit.get()
#             if interval_unit == "分钟":
#                 interval = interval_value * 60
#             else:  # 秒
#                 interval = interval_value
#
#             # 参数合法性校验
#             if duration <= 0 or interval <= 0 or interval > duration:
#                 raise ValueError
#         except ValueError:
#             messagebox.showwarning("警告", "请输入有效的监控参数（正整数，且间隔不大于时长）")
#             return
#
#         # 初始化数据存储（保持不变）
#         self.process_data = {}
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.process_data[proc_name] = []
#
#         # 更新状态和按钮（保持不变）
#         self.status_var.set("监控中...")
#         self.start_btn.config(state=tk.DISABLED)
#         self.stop_btn.config(state=tk.NORMAL)
#         self.report_btn.config(state=tk.DISABLED)
#         self.monitoring = True
#
#         self.monitor_thread = threading.Thread(
#             target=self._monitor_processes,
#             args=(duration, interval),
#             daemon=True
#         )
#         self.monitor_thread.start()
#
#     def _stop_monitoring(self):
#         """停止监控进程（保持不变）"""
#         self.monitoring = False
#         self.status_var.set("监控已停止，准备生成报告")
#         self.start_btn.config(state=tk.NORMAL)
#         self.stop_btn.config(state=tk.DISABLED)
#         self.report_btn.config(state=tk.NORMAL)
#
#     def _monitor_processes(self, duration, interval):
#         """监控进程内存使用情况（保持不变）"""
#         end_time = time.time() + duration
#         while self.monitoring and time.time() < end_time:
#             timestamp = datetime.now()
#             for proc_name in self.process_data.keys():
#                 mem_usage = 0
#                 for proc in psutil.process_iter(['name', 'memory_info']):
#                     try:
#                         if proc.info['name'] == proc_name:
#                             mem_usage += proc.info['memory_info'].rss
#                     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                         continue
#                 self.process_data[proc_name].append({'Timestamp': timestamp, 'Memory_Bytes': mem_usage})
#             self.root.after(0, self._update_chart)
#             time.sleep(interval)
#         if self.monitoring:
#             self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
#             self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
#             self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
#             self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
#             self.monitoring = False
#
#     # ---------------------- 实时图表更新（修改：x轴时间优化） ----------------------
#     def _update_chart(self):
#         """更新实时图表（优化x轴时间显示）"""
#         self.ax.clear()
#         has_data = False
#
#         # 绘制每个进程的内存使用曲线
#         for proc_name, data in self.process_data.items():
#             if data:
#                 has_data = True
#                 df = pd.DataFrame(data)
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 # ---------------------- 修改4：使用datetime类型x轴 ----------------------
#                 self.ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)
#
#         if has_data:
#             self.ax.set_title('实时内存使用监控')
#             self.ax.set_xlabel('时间')
#             self.ax.set_ylabel('内存使用 (MB)')
#             self.ax.legend()
#
#             # ---------------------- 修改5：x轴刻度优化 ----------------------
#             # 限制最大刻度数量（最多10个）
#             max_ticks = min(10, len(df))
#             self.ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#             # 时间格式化（时:分:秒）
#             self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#             # 标签旋转45度并右对齐
#             plt.xticks(rotation=45, ha='right')
#             self.fig.tight_layout()
#
#         self.canvas.draw()
#
#     # ---------------------- 报告生成方法（修改：图表x轴优化） ----------------------
#     def _generate_report(self):
#         """生成汇总Excel报告（优化图表x轴显示）"""
#         if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
#             messagebox.showwarning("警告", "没有监控数据可生成报告")
#             return
#
#         excel_path = os.path.join(self.save_path, "memory_summary_report.xlsx")
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "内存监控汇总"
#
#         # 步骤1：汇总表格数据（保持不变）
#         summary_data = {}
#         process_names = [self.monitor_listbox.get(i) for i in range(self.monitor_listbox.size())]
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')
#                 for _, row in df.iterrows():
#                     timestamp = row['Timestamp']
#                     if timestamp not in summary_data:
#                         summary_data[timestamp] = {'Timestamp': timestamp, 'Time': row['Time']}
#                     summary_data[timestamp][proc_name] = row['Memory_MB']
#         if not summary_data:
#             messagebox.showwarning("警告", "监控数据为空，无法生成报告")
#             return
#         summary_df = pd.DataFrame.from_dict(summary_data, orient='index').sort_values('Timestamp')
#
#         # 写入Excel表格并居中对齐（保持不变）
#         headers = ['Timestamp'] + process_names
#         ws.append(headers)
#         for _, row in summary_df.iterrows():
#             row_data = [row['Timestamp']] + [row.get(proc, '') for proc in process_names]
#             ws.append(row_data)
#         center_alignment = Alignment(horizontal='center', vertical='center')
#         for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
#             for cell in row:
#                 if isinstance(cell.value, (int, float)):
#                     cell.number_format = '0.00'
#                 cell.alignment = center_alignment
#         ws.column_dimensions['A'].width = 25
#         for col in range(2, 2 + len(process_names)):
#             ws.column_dimensions[chr(64 + col)].width = 18
#
#         # 步骤2：生成合并图表（修改x轴优化）
#         current_row = len(summary_df) + 3
#         merge_chart_inserted = False
#         if self.merge_var.get() and self.merge_target_listbox.size() > 0:
#             merge_procs = [self.merge_target_listbox.get(i) for i in range(self.merge_target_listbox.size())]
#             if merge_procs:
#                 fig, ax = plt.subplots(figsize=(12, 6))
#                 colors = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
#                 color_idx = 0
#
#                 for proc_name in merge_procs:
#                     if proc_name in self.process_data and self.process_data[proc_name]:
#                         df = pd.DataFrame(self.process_data[proc_name])
#                         df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                         # ---------------------- 修改6：合并图表x轴使用datetime ----------------------
#                         ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-',
#                                 label=proc_name, color=colors[color_idx % len(colors)])
#                         color_idx += 1
#
#                 ax.set_title('多进程内存使用对比（合并图表）')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#                 ax.legend()
#                 # ---------------------- 修改7：合并图表x轴刻度优化 ----------------------
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#
#                 # 插入图表（保持不变）
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 800
#                 img.height = 400
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 30
#                 plt.close(fig)
#                 merge_chart_inserted = True
#
#         # 步骤3：生成单个进程图表（修改x轴优化）
#         if not merge_chart_inserted:
#             current_row = len(summary_df) + 3
#         else:
#             current_row += 5
#
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 fig, ax = plt.subplots(figsize=(10, 4))
#                 # ---------------------- 修改8：单个图表x轴使用datetime ----------------------
#                 ax.plot(df['Timestamp'], df['Memory_MB'], marker='o', linestyle='-', color='blue')
#                 ax.set_title(f'{proc_name} 内存使用趋势')
#                 ax.set_xlabel('时间')
#                 ax.set_ylabel('内存使用 (MB)')
#
#                 # ---------------------- 修改9：单个图表x轴刻度优化 ----------------------
#                 max_ticks = min(10, len(df))
#                 ax.xaxis.set_major_locator(plt.MaxNLocator(max_ticks))
#                 ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
#                 plt.xticks(rotation=45, ha='right')
#                 plt.tight_layout()
#
#                 # 插入图表（保持不变）
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 600
#                 img.height = 300
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 20
#                 plt.close(fig)
#
#         # 保存报告（保持不变）
#         try:
#             wb.save(excel_path)
#             messagebox.showinfo("成功", f"汇总报告已生成：\n{excel_path}")
#             self.status_var.set("报告生成完成")
#         except Exception as e:
#             messagebox.showerror("错误", f"保存报告失败：{str(e)}")
#
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = MemoryMonitorApp(root)
#     root.mainloop()



# import psutil
# import pandas as pd
# import time
# import threading
# import os
# import matplotlib.pyplot as plt
# from datetime import datetime
# from openpyxl import Workbook
# from openpyxl.drawing.image import Image
# from openpyxl.utils.dataframe import dataframe_to_rows
# from openpyxl.styles import Alignment
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# import matplotlib
# from io import BytesIO
#
# matplotlib.use('TkAgg')
# plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
#
#
# class MemoryMonitorApp:
#     def __init__(self, root):
#         """初始化内存监控应用"""
#         self.root = root
#         self.root.title("多进程内存监控工具")
#         self.root.geometry("900x700")
#         self.root.resizable(True, True)
#
#         # 初始化变量
#         self.monitoring = False
#         self.monitor_thread = None
#         self.process_data = {}  # 存储监控数据 {进程名: DataFrame}
#         self.selected_processes = set()  # 选中的监控进程
#         self.merge_processes = set()  # 选中的合并图表进程
#         self.save_path = os.getcwd()  # 默认保存路径
#
#         # 创建UI界面
#         self._create_widgets()
#
#     def _create_widgets(self):
#         """创建UI组件"""
#         # 创建主框架
#         main_frame = ttk.Frame(self.root, padding="10")
#         main_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 1. 进程选择区域
#         process_frame = ttk.LabelFrame(main_frame, text="进程选择", padding="10")
#         process_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 进程列表
#         ttk.Label(process_frame, text="可用进程:").grid(row=0, column=0, sticky=tk.W)
#         self.process_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.process_listbox.grid(row=1, column=0, padx=(0, 10), pady=5)
#
#         # 滚动条
#         scrollbar = ttk.Scrollbar(process_frame, orient=tk.VERTICAL, command=self.process_listbox.yview)
#         scrollbar.grid(row=1, column=1, sticky=tk.NS)
#         self.process_listbox.config(yscrollcommand=scrollbar.set)
#
#         # 按钮
#         btn_frame = ttk.Frame(process_frame)
#         btn_frame.grid(row=1, column=2, padx=10)
#
#         ttk.Button(btn_frame, text="刷新进程", command=self._refresh_processes).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="添加监控", command=self._add_monitor).pack(fill=tk.X, pady=5)
#         ttk.Button(btn_frame, text="移除监控", command=self._remove_monitor).pack(fill=tk.X, pady=5)
#
#         # 当前监控进程
#         ttk.Label(process_frame, text="当前监控进程:").grid(row=0, column=3, sticky=tk.W)
#         self.monitor_listbox = tk.Listbox(process_frame, selectmode=tk.EXTENDED, height=6, width=50)
#         self.monitor_listbox.grid(row=1, column=3, padx=(0, 10), pady=5)
#
#         # 2. 参数设置区域
#         param_frame = ttk.LabelFrame(main_frame, text="监控参数", padding="10")
#         param_frame.pack(fill=tk.X, pady=(0, 10))
#
#         ttk.Label(param_frame, text="监控时长(秒):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
#         self.duration_var = tk.StringVar(value="60")
#         ttk.Entry(param_frame, textvariable=self.duration_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
#
#         ttk.Label(param_frame, text="采样间隔(秒):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
#         self.interval_var = tk.StringVar(value="5")
#         ttk.Entry(param_frame, textvariable=self.interval_var, width=10).grid(row=0, column=3, sticky=tk.W, pady=5)
#
#         ttk.Label(param_frame, text="报告保存路径:").grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
#         self.path_var = tk.StringVar(value=os.getcwd())
#         ttk.Entry(param_frame, textvariable=self.path_var, width=30).grid(row=0, column=5, sticky=tk.W, pady=5)
#         ttk.Button(param_frame, text="浏览...", command=self._browse_path).grid(row=0, column=6, padx=5, pady=5)
#
#         # 3. 图表合并配置（修改部分）
#         merge_frame = ttk.LabelFrame(main_frame, text="图表合并配置", padding="10")
#         merge_frame.pack(fill=tk.X, pady=(0, 10))
#
#         # 启用图表合并复选框
#         self.merge_var = tk.BooleanVar(value=False)
#         merge_check = ttk.Checkbutton(merge_frame, text="启用图表合并", variable=self.merge_var,
#                                       command=self._toggle_merge_options)
#         merge_check.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5, columnspan=3)
#
#         # 左侧：当前监控进程
#         ttk.Label(merge_frame, text="当前监控进程:").grid(row=1, column=0, sticky=tk.W, padx=5)
#         self.merge_source_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_source_listbox.grid(row=2, column=0, padx=5, pady=5)
#
#         # 中间：添加/移除按钮
#         merge_btn_frame = ttk.Frame(merge_frame)
#         merge_btn_frame.grid(row=2, column=1, padx=10)
#
#         self.add_to_merge_btn = ttk.Button(merge_btn_frame, text="添加 >", command=self._add_to_merge,
#                                            state=tk.DISABLED)
#         self.add_to_merge_btn.pack(fill=tk.X, pady=5)
#
#         self.remove_from_merge_btn = ttk.Button(merge_btn_frame, text="< 移除", command=self._remove_from_merge,
#                                                 state=tk.DISABLED)
#         self.remove_from_merge_btn.pack(fill=tk.X, pady=5)
#
#         # 右侧：合并监控进程
#         ttk.Label(merge_frame, text="合并监控进程:").grid(row=1, column=2, sticky=tk.W, padx=5)
#         self.merge_target_listbox = tk.Listbox(merge_frame, selectmode=tk.EXTENDED, height=4, width=40)
#         self.merge_target_listbox.grid(row=2, column=2, padx=5, pady=5)
#
#         # 4. 控制按钮区域
#         control_frame = ttk.Frame(main_frame)
#         control_frame.pack(fill=tk.X, pady=(0, 10))
#
#         self.start_btn = ttk.Button(control_frame, text="开始监控", command=self._start_monitoring)
#         self.start_btn.pack(side=tk.LEFT, padx=5)
#
#         self.stop_btn = ttk.Button(control_frame, text="停止监控", command=self._stop_monitoring, state=tk.DISABLED)
#         self.stop_btn.pack(side=tk.LEFT, padx=5)
#
#         self.report_btn = ttk.Button(control_frame, text="生成报告", command=self._generate_report, state=tk.DISABLED)
#         self.report_btn.pack(side=tk.LEFT, padx=5)
#
#         self.status_var = tk.StringVar(value="就绪")
#         ttk.Label(control_frame, textvariable=self.status_var).pack(side=tk.RIGHT, padx=5)
#
#         # 5. 实时图表区域
#         chart_frame = ttk.LabelFrame(main_frame, text="实时监控图表", padding="10")
#         chart_frame.pack(fill=tk.BOTH, expand=True)
#
#         # 创建matplotlib图表
#         self.fig, self.ax = plt.subplots(figsize=(8, 4))
#         self.canvas = FigureCanvasTkAgg(self.fig, master=chart_frame)
#         self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
#
#         # 初始化刷新进程列表
#         self._refresh_processes()
#
#     def _toggle_merge_options(self):
#         """切换合并选项的启用/禁用状态"""
#         state = tk.NORMAL if self.merge_var.get() else tk.DISABLED
#
#         # 更新列表状态
#         self.merge_source_listbox.config(state=state)
#         self.merge_target_listbox.config(state=state)
#
#         # 更新按钮状态
#         self.add_to_merge_btn.config(state=state)
#         self.remove_from_merge_btn.config(state=state)
#
#         # 如果启用，同步当前监控进程到合并源列表
#         if state == tk.NORMAL:
#             self._sync_merge_source_list()
#
#     def _sync_merge_source_list(self):
#         """同步当前监控进程到合并源列表"""
#         self.merge_source_listbox.delete(0, tk.END)
#         # 添加当前监控进程到合并源列表
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.merge_source_listbox.insert(tk.END, proc_name)
#
#     def _add_to_merge(self):
#         """将选中的进程添加到合并列表"""
#         selected_indices = self.merge_source_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.merge_source_listbox.get(i)
#             # 检查是否已在合并目标列表中
#             if proc_name not in [self.merge_target_listbox.get(j) for j in range(self.merge_target_listbox.size())]:
#                 self.merge_target_listbox.insert(tk.END, proc_name)
#
#     def _remove_from_merge(self):
#         """从合并列表中移除选中的进程"""
#         selected_indices = self.merge_target_listbox.curselection()
#         # 反向删除，避免索引变化
#         for i in sorted(selected_indices, reverse=True):
#             self.merge_target_listbox.delete(i)
#
#     def _refresh_processes(self):
#         """刷新进程列表"""
#         self.process_listbox.delete(0, tk.END)
#
#         # 获取唯一进程名
#         processes = set()
#         for proc in psutil.process_iter(['name']):
#             try:
#                 processes.add(proc.info['name'])
#             except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                 continue
#
#         # 添加到列表
#         for proc_name in sorted(processes):
#             self.process_listbox.insert(tk.END, proc_name)
#
#     def _add_monitor(self):
#         """添加选中的进程到监控列表"""
#         selected_indices = self.process_listbox.curselection()
#         for i in selected_indices:
#             proc_name = self.process_listbox.get(i)
#             # 检查是否已在监控列表中
#             if proc_name not in [self.monitor_listbox.get(j) for j in range(self.monitor_listbox.size())]:
#                 self.monitor_listbox.insert(tk.END, proc_name)
#                 # 如果合并功能已启用，同步到合并源列表
#                 if self.merge_var.get():
#                     self._sync_merge_source_list()
#
#     def _remove_monitor(self):
#         """从监控列表中移除选中的进程"""
#         selected_indices = self.monitor_listbox.curselection()
#         # 反向删除，避免索引变化
#         for i in sorted(selected_indices, reverse=True):
#             self.monitor_listbox.delete(i)
#             # 如果合并功能已启用，同步到合并源列表
#             if self.merge_var.get():
#                 self._sync_merge_source_list()
#
#     def _browse_path(self):
#         """浏览保存路径"""
#         path = filedialog.askdirectory()
#         if path:
#             self.path_var.set(path)
#             self.save_path = path
#
#     def _start_monitoring(self):
#         """开始监控进程"""
#         # 检查是否选择了进程
#         if self.monitor_listbox.size() == 0:
#             messagebox.showwarning("警告", "请至少选择一个进程进行监控")
#             return
#
#         # 检查参数
#         try:
#             duration = int(self.duration_var.get())
#             interval = int(self.interval_var.get())
#             if duration <= 0 or interval <= 0 or interval > duration:
#                 raise ValueError
#         except ValueError:
#             messagebox.showwarning("警告", "请输入有效的监控时长和采样间隔")
#             return
#
#         # 初始化数据存储
#         self.process_data = {}
#         for i in range(self.monitor_listbox.size()):
#             proc_name = self.monitor_listbox.get(i)
#             self.process_data[proc_name] = []
#
#         # 更新状态和按钮
#         self.status_var.set("监控中...")
#         self.start_btn.config(state=tk.DISABLED)
#         self.stop_btn.config(state=tk.NORMAL)
#         self.report_btn.config(state=tk.DISABLED)
#         self.monitoring = True
#
#         # 在新线程中运行监控
#         self.monitor_thread = threading.Thread(
#             target=self._monitor_processes,
#             args=(duration, interval),
#             daemon=True
#         )
#         self.monitor_thread.start()
#
#     def _stop_monitoring(self):
#         """停止监控进程"""
#         self.monitoring = False
#         self.status_var.set("监控已停止，准备生成报告")
#         self.start_btn.config(state=tk.NORMAL)
#         self.stop_btn.config(state=tk.DISABLED)
#         self.report_btn.config(state=tk.NORMAL)
#
#     def _monitor_processes(self, duration, interval):
#         """监控进程内存使用情况"""
#         end_time = time.time() + duration
#
#         while self.monitoring and time.time() < end_time:
#             timestamp = datetime.now()
#
#             # 记录每个进程的内存使用
#             for proc_name in self.process_data.keys():
#                 mem_usage = 0
#                 for proc in psutil.process_iter(['name', 'memory_info']):
#                     try:
#                         if proc.info['name'] == proc_name:
#                             mem_usage += proc.info['memory_info'].rss  # 获取实际物理内存
#                     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#                         continue
#
#                 self.process_data[proc_name].append({
#                     'Timestamp': timestamp,
#                     'Memory_Bytes': mem_usage
#                 })
#
#             # 更新实时图表
#             self.root.after(0, self._update_chart)
#
#             # 等待采样间隔
#             time.sleep(interval)
#
#         # 如果是正常结束而非用户停止，自动启用报告按钮
#         if self.monitoring:
#             self.root.after(0, lambda: self.status_var.set("监控完成，准备生成报告"))
#             self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))
#             self.root.after(0, lambda: self.stop_btn.config(state=tk.DISABLED))
#             self.root.after(0, lambda: self.report_btn.config(state=tk.NORMAL))
#             self.monitoring = False
#
#     def _update_chart(self):
#         """更新实时图表"""
#         self.ax.clear()
#
#         # 绘制每个进程的内存使用曲线
#         for proc_name, data in self.process_data.items():
#             if data:
#                 df = pd.DataFrame(data)
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')
#                 self.ax.plot(df['Time'], df['Memory_MB'], marker='o', linestyle='-', label=proc_name)
#
#         self.ax.set_title('实时内存使用监控')
#         self.ax.set_xlabel('时间 (H:M:S)')
#         self.ax.set_ylabel('内存使用 (MB)')
#         self.ax.legend()
#         plt.xticks(rotation=45)
#         self.fig.tight_layout()
#         self.canvas.draw()
#
#     def _generate_report(self):
#         """生成汇总Excel报告（添加表格居中对齐）"""
#         # 检查是否有监控数据
#         if not self.process_data or all(len(data) == 0 for data in self.process_data.values()):
#             messagebox.showwarning("警告", "没有监控数据可生成报告")
#             return
#
#         # 创建汇总Excel文件
#         excel_path = os.path.join(self.save_path, "memory_summary_report.xlsx")
#         wb = Workbook()
#         ws = wb.active
#         ws.title = "内存监控汇总"
#
#         # ---------------------- 步骤1：汇总所有进程数据到表格（删除Time列） ----------------------
#         # 1.1 准备汇总数据（仍保留Time字段用于图表生成，但不写入Excel）
#         summary_data = {}
#         process_names = [self.monitor_listbox.get(i) for i in range(self.monitor_listbox.size())]
#
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)  # 转换为MB
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')  # 保留Time用于图表，但不写入Excel
#
#                 # 按时间戳合并数据（仅保留Timestamp和进程数据列）
#                 for _, row in df.iterrows():
#                     timestamp = row['Timestamp']
#                     if timestamp not in summary_data:
#                         summary_data[timestamp] = {
#                             'Timestamp': timestamp,  # 仅保留Timestamp用于表格
#                             'Time': row['Time']  # 内部使用，不写入Excel
#                         }
#                     summary_data[timestamp][proc_name] = row['Memory_MB']
#
#         # 1.2 转换为DataFrame并排序
#         if not summary_data:
#             messagebox.showwarning("警告", "监控数据为空，无法生成报告")
#             return
#         summary_df = pd.DataFrame.from_dict(summary_data, orient='index').sort_values('Timestamp')
#
#         # 1.3 写入Excel表格（删除Time列）
#         headers = ['Timestamp'] + process_names  # 仅保留Timestamp和进程名
#         ws.append(headers)
#
#         for _, row in summary_df.iterrows():
#             row_data = [row['Timestamp']]  # 仅保留Timestamp
#             for proc in process_names:
#                 row_data.append(row.get(proc, ''))  # 进程数据列
#             ws.append(row_data)
#
#         # ---------------------- 新增：设置表格居中对齐 ----------------------
#         # 创建居中对齐样式
#         center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
#
#         # 1. 设置表头居中
#         header_row = ws[1]  # Excel行索引从1开始
#         for cell in header_row:
#             cell.alignment = center_alignment
#
#         # 2. 设置数据行居中
#         # 获取数据区域范围（从第2行到最后一行，从第1列到最后一列）
#         max_row = ws.max_row
#         max_col = ws.max_column
#
#         for row in ws.iter_rows(min_row=2, max_row=max_row, min_col=1, max_col=max_col):
#             for cell in row:
#                 # 对数值单元格特殊处理（保持数值类型）
#                 if isinstance(cell.value, (int, float)):
#                     cell.number_format = '0.00'  # 保留两位小数
#                 cell.alignment = center_alignment
#
#         # ---------------------- 调整列宽（优化居中显示效果） ----------------------
#         ws.column_dimensions['A'].width = 25  # 加宽Timestamp列以适应居中显示
#         for col in range(2, 2 + len(process_names)):
#             ws.column_dimensions[chr(64 + col)].width = 18  # 加宽进程数据列
#
#         # ---------------------- 图表生成部分（保持不变） ----------------------
#         # ...（合并图表和单个图表生成逻辑与之前完全一致）
#         # ---------------------- 后续步骤保持不变（图表顺序和生成逻辑） ----------------------
#         current_row = len(summary_df) + 3  # 从数据下方3行开始插入图表
#         merge_chart_inserted = False
#
#         # 步骤2：生成合并折线图（作为第一个图表）
#         if self.merge_var.get() and self.merge_target_listbox.size() > 0:
#             merge_procs = [self.merge_target_listbox.get(i) for i in range(self.merge_target_listbox.size())]
#             if merge_procs:
#                 # 生成合并图表（代码不变，仍使用df['Time']绘制X轴）
#                 fig, ax = plt.subplots(figsize=(12, 6))
#                 colors = ['blue', 'green', 'red', 'purple', 'orange', 'brown', 'pink', 'gray']
#                 color_idx = 0
#
#                 for proc_name in merge_procs:
#                     if proc_name in self.process_data and self.process_data[proc_name]:
#                         df = pd.DataFrame(self.process_data[proc_name])
#                         df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                         df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')  # 仍使用Time字段
#                         ax.plot(df['Time'], df['Memory_MB'], marker='o', linestyle='-',
#                                 label=proc_name, color=colors[color_idx % len(colors)])
#                         color_idx += 1
#
#                 ax.set_title('多进程内存使用对比（合并图表）')
#                 ax.set_xlabel('时间 (H:M:S)')
#                 ax.set_ylabel('内存使用 (MB)')
#                 ax.legend()
#                 plt.xticks(rotation=45)
#                 plt.tight_layout()
#
#                 # 插入合并图表（代码不变）
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 800
#                 img.height = 400
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 30
#                 plt.close(fig)
#                 merge_chart_inserted = True
#
#         # 步骤3：生成单个进程折线图（代码不变）
#         if not merge_chart_inserted:
#             current_row = len(summary_df) + 3
#         else:
#             current_row += 5
#
#         for proc_name in process_names:
#             if proc_name in self.process_data and self.process_data[proc_name]:
#                 df = pd.DataFrame(self.process_data[proc_name])
#                 df['Memory_MB'] = df['Memory_Bytes'] / (1024 * 1024)
#                 df['Time'] = df['Timestamp'].dt.strftime('%H:%M:%S')  # 仍使用Time字段绘制图表
#                 # 生成单个进程图表（代码不变）
#                 fig, ax = plt.subplots(figsize=(10, 4))
#                 ax.plot(df['Time'], df['Memory_MB'], marker='o', linestyle='-', color='blue')
#                 ax.set_title(f'{proc_name} 内存使用趋势')
#                 ax.set_xlabel('时间 (H:M:S)')
#                 ax.set_ylabel('内存使用 (MB)')
#                 plt.xticks(rotation=45)
#                 plt.tight_layout()
#
#                 # 插入单个进程图表（代码不变）
#                 img_data = BytesIO()
#                 fig.savefig(img_data, format='png')
#                 img_data.seek(0)
#                 img = Image(img_data)
#                 img.width = 600
#                 img.height = 300
#                 ws.add_image(img, f'A{current_row}')
#                 current_row += 20
#                 plt.close(fig)
#
#
#         # 保存Excel文件
#         try:
#             wb.save(excel_path)
#             messagebox.showinfo("成功", f"汇总报告已生成：\n{excel_path}")
#             self.status_var.set("报告生成完成")
#         except Exception as e:
#             messagebox.showerror("错误", f"保存报告失败：{str(e)}")
#
# if __name__ == "__main__":
#     root = tk.Tk()
#     app = MemoryMonitorApp(root)
#     root.mainloop()