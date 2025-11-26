#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
屏幕数值OCR监控器 - 兼容版本
功能：读取屏幕特定区域，OCR识别数值，绘制图表并保存
作者：MiniMax Agent
版本：1.0.1
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import re
import csv
import os
from datetime import datetime
from PIL import ImageGrab, Image
import queue

def check_dependencies():
    """检查并报告依赖状态"""
    deps = {}
    
    # 检查必要依赖
    try:
        import pytesseract
        deps['pytesseract'] = 'OK'
    except ImportError:
        deps['pytesseract'] = 'MISSING'
        
    try:
        import matplotlib
        deps['matplotlib'] = 'OK'
    except ImportError:
        deps['matplotlib'] = 'MISSING'
        
    try:
        import numpy
        deps['numpy'] = f'OK (v{numpy.__version__})'
    except ImportError:
        deps['numpy'] = 'MISSING'
    
    # 检查可选依赖
    try:
        import cv2
        deps['opencv-python'] = 'AVAILABLE (not needed)'
    except ImportError:
        deps['opencv-python'] = 'NOT INSTALLED (not needed)'
        
    return deps

class ScreenOCROrMonitor:
    def __init__(self):
        # 检查依赖
        self.deps = check_dependencies()
        missing = [dep for dep, status in self.deps.items() if 'MISSING' in status]
        if missing:
            messagebox.showerror("依赖缺失", f"缺少必要的依赖库：\n{', '.join(missing)}\n\n请运行：pip install -r requirements.txt")
            return
            
        # 导入必要依赖
        try:
            import pytesseract
            self.pytesseract = pytesseract
        except ImportError:
            return
            
        try:
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            self.plt = plt
            self.FigureCanvasTkAgg = FigureCanvasTkAgg
        except ImportError:
            return
            
        try:
            import numpy as np
            self.np = np
        except ImportError:
            return
        
        # 设置matplotlib中文字体
        self.setup_matplotlib_font()
        
        self.root = tk.Tk()
        self.root.title("屏幕数值OCR监控器 v1.0.1")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        
        # 初始化变量
        self.monitoring = False
        self.capture_rect = None
        self.data_queue = queue.Queue()
        self.times = []
        self.values = []
        self.interval = 2.0
        self.max_points = 1000
        
        # 数据同步验证
        self.last_processed_count = 0
        
        # 图表渲染锁
        self._chart_update_lock = threading.Lock()
        
        # Tesseract路径设置
        self.setup_tesseract()
        
        # 创建界面
        self.create_widgets()
        
    def setup_tesseract(self):
        """设置Tesseract路径"""
        try:
            import pytesseract
            # 尝试常见路径
            tesseract_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            ]
            
            for path in tesseract_paths:
                if os.path.exists(path):
                    pytesseract.pytesseract.tesseract_cmd = path
                    break
        except Exception:
            pass
    
    def setup_matplotlib_font(self):
        """设置matplotlib中文字体"""
        try:
            import matplotlib.pyplot as plt
            import matplotlib.font_manager as fm
            
            # 尝试设置中文字体
            chinese_fonts = [
                'Microsoft YaHei',      # 微软雅黑
                'SimHei',               # 黑体
                'SimSun',               # 宋体
                'KaiTi',                # 楷体
                'FangSong',             # 仿宋
                'Microsoft JhengHei',   # 微软正黑体（繁体）
                'Arial Unicode MS',     # Arial Unicode
                'DejaVu Sans',          # DejaVu Sans
            ]
            
            # 检查可用的中文字体
            available_fonts = [f.name for f in fm.fontManager.ttflist]
            
            selected_font = None
            for font in chinese_fonts:
                if font in available_fonts:
                    selected_font = font
                    break
            
            if selected_font:
                plt.rcParams['font.sans-serif'] = [selected_font]
                plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
                print(f"[DEBUG] 使用中文字体: {selected_font}")
            else:
                # 如果没有找到中文字体，尝试设置通用字体
                plt.rcParams['font.sans-serif'] = ['DejaVu Sans', 'Arial Unicode MS', 'Arial']
                plt.rcParams['axes.unicode_minus'] = False
                print("[DEBUG] 未找到中文字体，使用默认字体")
                
        except Exception as e:
            print(f"[DEBUG] 字体设置异常: {e}")
    
    def create_widgets(self):
        """创建界面控件"""
        # 依赖状态显示
        status_frame = ttk.LabelFrame(self.root, text="依赖状态", padding=10)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        status_text = " | ".join([f"{dep}: {status}" for dep, status in self.deps.items()])
        status_label = ttk.Label(status_frame, text=status_text)
        status_label.pack(anchor=tk.W)
        
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding=10)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 按钮
        self.select_btn = ttk.Button(control_frame, text="选择监控区域", command=self.select_region)
        self.select_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.start_btn = ttk.Button(control_frame, text="开始监控", command=self.toggle_monitoring, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 间隔设置
        ttk.Label(control_frame, text="监控间隔(秒):").pack(side=tk.LEFT, padx=(20, 5))
        self.interval_var = tk.DoubleVar(value=2.0)
        interval_spinbox = ttk.Spinbox(control_frame, from_=0.5, to=10.0, increment=0.5, 
                                     textvariable=self.interval_var, width=8)
        interval_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空按钮
        self.clear_btn = ttk.Button(control_frame, text="清空数据", command=self.clear_data)
        self.clear_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 状态显示
        self.status_label = ttk.Label(control_frame, text="状态: 准备就绪")
        self.status_label.pack(side=tk.RIGHT)
        
        # 区域信息
        region_frame = ttk.LabelFrame(main_frame, text="监控区域信息", padding=10)
        region_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.region_info = ttk.Label(region_frame, text="未选择区域")
        self.region_info.pack(anchor=tk.W)
        
        # 图表框架
        if 'matplotlib' in self.deps:
            chart_frame = ttk.LabelFrame(main_frame, text="数据图表", padding=10)
            chart_frame.pack(fill=tk.BOTH, expand=True)
            
            # 创建matplotlib图形
            self.fig, self.ax = self.plt.subplots(figsize=(8, 4), dpi=100)
            self.ax.set_title("实时数值监控")
            self.ax.set_xlabel("时间")
            self.ax.set_ylabel("数值")
            self.ax.grid(True, alpha=0.3)
            
            # 嵌入画布
            self.canvas = self.FigureCanvasTkAgg(self.fig, chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 数据显示
        data_frame = ttk.LabelFrame(main_frame, text="实时数据", padding=10)
        data_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.current_value = ttk.Label(data_frame, text="当前值: --", font=("Arial", 12, "bold"))
        self.current_value.pack(side=tk.LEFT)
        
        # 保存按钮
        self.save_btn = ttk.Button(data_frame, text="保存图表", command=self.save_chart)
        self.save_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.save_csv_btn = ttk.Button(data_frame, text="导出CSV", command=self.save_csv)
        self.save_csv_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 启动队列处理
        self.process_queue()
    
    def select_region(self):
        """选择监控区域"""
        region_window = tk.Toplevel(self.root)
        region_window.attributes("-fullscreen", True)
        region_window.attributes("-alpha", 0.3)
        region_window.attributes("-topmost", True)
        region_window.configure(bg='black')
        
        canvas = tk.Canvas(region_window, bg='black', highlightthickness=0)
        canvas.pack(fill=tk.BOTH, expand=True)
        
        start_x = start_y = end_x = end_y = None
        rect_id = None
        
        def on_mouse_down(event):
            nonlocal start_x, start_y, rect_id
            start_x = event.x
            start_y = event.y
            rect_id = canvas.create_rectangle(start_x, start_y, start_x, start_y, 
                                            outline='red', width=2, fill='blue')
        
        def on_mouse_move(event):
            nonlocal rect_id
            if rect_id:
                canvas.delete(rect_id)
                rect_id = canvas.create_rectangle(start_x, start_y, event.x, event.y, 
                                                outline='red', width=2, fill='blue', stipple='gray50')
        
        def on_mouse_up(event):
            nonlocal end_x, end_y
            end_x = event.x
            end_y = event.y
            region_window.destroy()
            
            x1, y1 = min(start_x, end_x), min(start_y, end_y)
            x2, y2 = max(start_x, end_x), max(start_y, end_y)
            
            if x2 - x1 > 10 and y2 - y1 > 10:
                self.capture_rect = (x1, y1, x2, y2)
                self.region_info.config(text=f"监控区域: ({x1}, {y1}) - ({x2}, {y2})")
                self.start_btn.config(state=tk.NORMAL)
                self.status_label.config(text="状态: 区域已选择，可以开始监控")
            else:
                messagebox.showerror("错误", "选择区域太小，请重新选择")
        
        canvas.bind("<Button-1>", on_mouse_down)
        canvas.bind("<B1-Motion>", on_mouse_move)
        canvas.bind("<ButtonRelease-1>", on_mouse_up)
        
        canvas.create_text(400, 50, text="拖拽鼠标选择监控区域", fill='white', 
                          font=("Arial", 16), tag="tip")
    
    def toggle_monitoring(self):
        """切换监控状态"""
        if not self.monitoring:
            self.start_monitoring()
        else:
            self.stop_monitoring()
    
    def start_monitoring(self):
        """开始监控"""
        if not self.capture_rect:
            messagebox.showerror("错误", "请先选择监控区域")
            return
            
        self.monitoring = True
        self.start_btn.config(text="停止监控")
        self.select_btn.config(state=tk.DISABLED)
        self.interval = self.interval_var.get()
        self.status_label.config(text="状态: 监控中...")
        
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()
    
    def stop_monitoring(self):
        """停止监控"""
        self.monitoring = False
        self.start_btn.config(text="开始监控")
        self.select_btn.config(state=tk.NORMAL)
        self.status_label.config(text="状态: 已停止")
    
    def monitor_loop(self):
        """监控主循环 - 线程安全版本"""
        while self.monitoring:
            try:
                if self.capture_rect:
                    screenshot = ImageGrab.grab(bbox=self.capture_rect)
                    text = self.pytesseract.image_to_string(screenshot, lang='chi_sim+eng')
                    value = self.parse_value(text)
                    
                    # 只有成功解析到有效数值才放入队列
                    if value is not None and value >= 0:
                        timestamp = datetime.now()
                        # 线程安全地放入队列
                        try:
                            self.data_queue.put((timestamp, value), block=False)
                            print(f"[DEBUG] 解析成功: {value:.2f} at {timestamp.strftime('%H:%M:%S')}")
                        except queue.Full:
                            print(f"[DEBUG] 队列已满，丢弃数据: {value:.2f}")
                    else:
                        print(f"[DEBUG] 解析失败或无效值: {text.strip()}")
                
                time.sleep(self.interval)
            except Exception as e:
                print(f"[DEBUG] 监控错误: {e}")
                time.sleep(1)
    
    def parse_value(self, text):
        """解析数值"""
        try:
            # 清理文本，移除多余字符
            cleaned_text = text.strip()
            print(f"[DEBUG] OCR原始文本: '{cleaned_text}'")
            
            patterns = [
                r'计数率[：:]\s*(\d+\.?\d*)\s*cps',
                r'(\d+\.?\d*)\s*cps',
                r'计数率[：:]\s*(\d+\.?\d*)',
                r'Rate[：:]\s*(\d+\.?\d*)',
                r'数值[：:]\s*(\d+\.?\d*)',
                # 添加更多模式
                r'(\d{1,6}\.?\d{0,2})',  # 匹配1-6位整数加可选小数部分
                r'\b(\d+\.?\d*)\b',  # 匹配独立的数字
            ]
            
            for i, pattern in enumerate(patterns):
                match = re.search(pattern, cleaned_text, re.IGNORECASE)
                if match:
                    value_str = match.group(1)
                    value = float(value_str)
                    print(f"[DEBUG] 模式{i+1}匹配成功: {value_str} -> {value}")
                    
                    # 检查数值合理性（0-100000的范围比较合理）
                    if 0 <= value <= 100000:
                        return value
                    else:
                        print(f"[DEBUG] 数值超出合理范围: {value}")
                        continue
            
            print(f"[DEBUG] 所有模式都未匹配到有效数值")
            return None
        except Exception as e:
            print(f"[DEBUG] 数值解析异常: {e}")
            return None
    
    def process_queue(self):
        """处理数据队列 - 线程安全版本"""
        # 确保在主线程中执行UI更新
        if threading.current_thread() != threading.main_thread():
            return
            
        try:
            new_data_added = False
            batch_data = []
            
            # 批量处理队列中的所有数据
            while True:
                try:
                    timestamp, value = self.data_queue.get_nowait()
                    
                    # 只添加有效的数据点
                    if value is not None and value >= 0:
                        batch_data.append((timestamp, value))
                        print(f"[DEBUG] 添加数据点: {value:.2f} at {timestamp.strftime('%H:%M:%S')}")
                    
                except queue.Empty:
                    break
            
            # 批量更新数据（减少锁竞争）
            if batch_data:
                for timestamp, value in batch_data:
                    self.times.append(timestamp)
                    self.values.append(value)
                
                new_data_added = True
                
                # 限制数据点数量
                if len(self.times) > self.max_points:
                    excess = len(self.times) - self.max_points
                    self.times = self.times[excess:]
                    self.values = self.values[excess:]
            
            # 如果有新数据添加，更新显示
            if new_data_added and self.times and self.values:
                # 更新当前值显示（确保使用同一个数据源）
                latest_value = self.values[-1]
                if hasattr(self, 'current_value'):
                    self.current_value.config(text=f"当前值: {latest_value:.2f}")
                
                print(f"[DEBUG] 更新显示: 当前值 = {latest_value:.2f}, 数据点总数 = {len(self.values)}")
                
                # 数据同步验证
                processed_count = len(self.values)
                if processed_count > self.last_processed_count:
                    self.last_processed_count = processed_count
                    print(f"[DEBUG] 数据同步验证: 已处理 {processed_count} 个有效数据点")
                
                # 安全地更新图表
                self.update_chart()
                
        except Exception as e:
            print(f"[DEBUG] 队列处理异常: {e}")
            import traceback
            traceback.print_exc()
        
        # 安全地重新调度
        try:
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.after(100, self.process_queue)
        except:
            pass
    
    def update_chart(self):
        """更新图表 - 线程安全版本"""
        # 在主线程中执行图表更新
        if not hasattr(self, 'root') or not self.root.winfo_exists():
            return
            
        # 确保在主线程中执行
        if threading.current_thread() != threading.main_thread():
            self.root.after(0, self._update_chart_safe)
            return
        
        self._update_chart_safe()
    
    def _update_chart_safe(self):
        """安全地更新图表（在主线程中执行）"""
        if not hasattr(self, 'ax'):
            return
            
        try:
            # 获取数据副本，避免并发访问问题
            times_copy = list(self.times) if hasattr(self, 'times') else []
            values_copy = list(self.values) if hasattr(self, 'values') else []
            
            if len(times_copy) == 0 or len(values_copy) == 0:
                return
            
            # 确保数据长度一致
            current_data_length = min(len(times_copy), len(values_copy))
            if current_data_length == 0:
                return
                
            # 获取最新的同步数据副本
            latest_times = times_copy[-current_data_length:]
            latest_values = values_copy[-current_data_length:]
            
            # 验证数据有效性
            valid_data = [(t, v) for t, v in zip(latest_times, latest_values) 
                         if v is not None and v >= 0]
            
            if not valid_data:
                print("[DEBUG] 没有有效数据用于绘图")
                return
                
            valid_times, valid_values = zip(*valid_data)
            
            # 计算时间差（基于第一个数据点的时间）
            if len(valid_times) > 1:
                base_time = valid_times[0]
                time_deltas = [(t - base_time).total_seconds() for t in valid_times]
            else:
                time_deltas = [0]
            
            # 清空图表并重新绘制
            self.ax.clear()
            
            # 绘制曲线
            self.ax.plot(time_deltas, valid_values, 'b-', linewidth=1.5, alpha=0.8, label='数值')
            self.ax.scatter(time_deltas, valid_values, c='red', s=20, alpha=0.6, zorder=5)
            
            # 获取最新的值用于显示和调试
            current_display_value = valid_values[-1] if valid_values else None
            
            print(f"[DEBUG] 图表安全更新: {len(valid_values)}个数据点")
            print(f"[DEBUG] 显示当前值: {current_display_value:.2f}, 图表最新点: {valid_values[-1]:.2f}")
            print(f"[DEBUG] 数据范围: {min(valid_values):.2f}-{max(valid_values):.2f}")
            
            # 设置图表样式
            self.ax.set_title("实时数值监控")
            self.ax.set_xlabel("时间 (秒)")
            self.ax.set_ylabel("数值")
            self.ax.grid(True, alpha=0.3)
            
            # 如果有多个数据点，添加统计信息
            if len(valid_values) > 1:
                avg_val = self.np.mean(valid_values)
                max_val = self.np.max(valid_values)
                min_val = self.np.min(valid_values)
                
                # 添加参考线
                self.ax.axhline(y=avg_val, color='g', linestyle='--', alpha=0.7, label=f'平均值: {avg_val:.2f}')
                self.ax.axhline(y=max_val, color='r', linestyle='--', alpha=0.7, label=f'最大值: {max_val:.2f}')
                self.ax.axhline(y=min_val, color='orange', linestyle='--', alpha=0.7, label=f'最小值: {min_val:.2f}')
                
                self.ax.legend(loc='upper right', fontsize=8)
            
            # 自动调整Y轴范围，避免显示0
            y_min, y_max = min(valid_values), max(valid_values)
            margin = (y_max - y_min) * 0.1 if y_max > y_min else 1
            self.ax.set_ylim(max(0, y_min - margin), y_max + margin)
            
            # 安全地刷新画布
            if hasattr(self, 'canvas'):
                self.canvas.draw()
            
        except Exception as e:
            print(f"[DEBUG] 图表更新异常: {e}")
            import traceback
            traceback.print_exc()
    
    def clear_data(self):
        """清空数据 - 线程安全版本"""
        # 清空数据队列，防止残留数据
        while True:
            try:
                self.data_queue.get_nowait()
            except queue.Empty:
                break
        
        # 使用锁保护数据清空操作
        with self._chart_update_lock:
            # 清空数据列表
            self.times.clear()
            self.values.clear()
            
            # 重置计数器
            self.last_processed_count = 0
            
            # 安全地更新显示
            if hasattr(self, 'current_value'):
                self.current_value.config(text="当前值: --")
            
            # 清空图表
            if hasattr(self, 'ax'):
                self.ax.clear()
                self.ax.set_title("实时数值监控")
                self.ax.set_xlabel("时间")
                self.ax.set_ylabel("数值")
                self.ax.grid(True, alpha=0.3)
                
                # 安全地刷新画布
                if hasattr(self, 'canvas'):
                    try:
                        self.canvas.draw()
                    except:
                        pass
        
        print("[DEBUG] 数据已清空，所有数据结构已重置")
    
    def save_chart(self):
        """保存图表"""
        if not hasattr(self, 'fig') or len(self.times) == 0:
            messagebox.showwarning("警告", "没有数据可保存")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
            title="保存图表"
        )
        
        if filename:
            try:
                self.fig.savefig(filename, dpi=300, bbox_inches='tight')
                messagebox.showinfo("成功", f"图表已保存到: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"保存失败: {e}")
    
    def save_csv(self):
        """导出CSV"""
        if len(self.times) == 0:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="导出CSV数据"
        )
        
        if filename:
            try:
                with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(['时间', '数值'])
                    for time, value in zip(self.times, self.values):
                        writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), value])
                messagebox.showinfo("成功", f"数据已导出到: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}")
    
    def run(self):
        """启动程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ScreenOCROrMonitor()
    app.run()
