#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
屏幕数值OCR监控器 - 多通道自适应版本
功能：多通道监控屏幕多个区域，OCR识别数值，自适应分辨率，绘制图表并保存
版本：2.1.1
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
            from matplotlib.patches import Rectangle
            import matplotlib.dates as mdates
            self.plt = plt
            self.FigureCanvasTkAgg = FigureCanvasTkAgg
            self.Rectangle = Rectangle
            self.mdates = mdates
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
        self.root.title("核医屏幕数值OCR监控器 v2.1.1 - 多通道自适应")
        self.root.geometry("1400x900")
        self.root.resizable(True, True)
        
        # 初始化变量
        self.monitoring = False
        self.channels = []  # 多通道数据
        self.active_channel_index = 0  # 当前活动通道
        self.data_queue = queue.Queue()
        self.interval = 2.0
        self.max_points = 1000
        
        # 图表交互相关变量
        self.selected_points = []  # 选中的数据点
        self.drag_start = None  # 拖动起始位置
        self.is_dragging = False  # 是否正在拖动
        self.original_xlim = None  # 原始X轴范围
        self.original_ylim = None  # 原始Y轴范围
        
        # 区域选择相关变量
        self.region_windows = []  # 存储区域选择窗口引用
        self.region_rectangles = {}  # 存储区域矩形引用 {channel_index: rect_id}
        
        # 获取屏幕缩放比例
        self.scale_factor = self.get_screen_scale_factor()
        
        # 数据同步验证
        self.last_processed_count = 0
        
        # 图表渲染锁
        self._chart_update_lock = threading.Lock()
        
        # Tesseract路径设置
        self.setup_tesseract()
        
        # 创建界面
        self.create_widgets()
        
    def get_screen_scale_factor(self):
        """获取屏幕缩放比例"""
        try:
            # 尝试获取Windows缩放比例
            import ctypes
            try:
                # Windows 8.1及以上
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
                scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100
            except:
                # 老版本Windows或其他系统
                scale_factor = 1.0
        except:
            scale_factor = 1.0
            
        print(f"[DEBUG] 屏幕缩放比例: {scale_factor}")
        return scale_factor
    
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
        
        # 缩放比例显示
        scale_label = ttk.Label(status_frame, text=f"屏幕缩放比例: {self.scale_factor}")
        scale_label.pack(anchor=tk.W)
        
        # 主框架
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 控制面板
        control_frame = ttk.LabelFrame(main_frame, text="控制面板", padding=10)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 通道管理
        channel_frame = ttk.Frame(control_frame)
        channel_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(channel_frame, text="通道管理:").pack(side=tk.LEFT)
        self.add_channel_btn = ttk.Button(channel_frame, text="添加监控区域", command=self.add_channel)
        self.add_channel_btn.pack(side=tk.LEFT, padx=(10, 5))
        
        self.remove_channel_btn = ttk.Button(channel_frame, text="删除当前通道", command=self.remove_channel, state=tk.DISABLED)
        self.remove_channel_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 通道选择
        ttk.Label(channel_frame, text="当前通道:").pack(side=tk.LEFT, padx=(20, 5))
        self.channel_var = tk.StringVar(value="无通道")
        self.channel_combo = ttk.Combobox(channel_frame, textvariable=self.channel_var, state="readonly", width=15)
        self.channel_combo.pack(side=tk.LEFT, padx=(0, 10))
        self.channel_combo.bind('<<ComboboxSelected>>', self.on_channel_change)
        
        # 显示区域按钮
        self.show_regions_btn = ttk.Button(channel_frame, text="显示所有区域", command=self.show_all_regions)
        self.show_regions_btn.pack(side=tk.LEFT, padx=(20, 5))
        
        self.hide_regions_btn = ttk.Button(channel_frame, text="隐藏所有区域", command=self.hide_all_regions)
        self.hide_regions_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 监控控制
        monitor_frame = ttk.Frame(control_frame)
        monitor_frame.pack(fill=tk.X)
        
        self.start_btn = ttk.Button(monitor_frame, text="开始监控", command=self.toggle_monitoring, state=tk.DISABLED)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 间隔设置
        ttk.Label(monitor_frame, text="监控间隔(秒):").pack(side=tk.LEFT, padx=(20, 5))
        self.interval_var = tk.DoubleVar(value=2.0)
        interval_spinbox = ttk.Spinbox(monitor_frame, from_=0.5, to=10.0, increment=0.5, 
                                     textvariable=self.interval_var, width=8)
        interval_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空按钮
        self.clear_btn = ttk.Button(monitor_frame, text="清空所有数据", command=self.clear_all_data)
        self.clear_btn.pack(side=tk.LEFT, padx=(10, 0))
        
        # 图表控制按钮
        self.zoom_in_btn = ttk.Button(monitor_frame, text="放大", command=self.zoom_in)
        self.zoom_in_btn.pack(side=tk.LEFT, padx=(20, 5))
        
        self.zoom_out_btn = ttk.Button(monitor_frame, text="缩小", command=self.zoom_out)
        self.zoom_out_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.reset_view_btn = ttk.Button(monitor_frame, text="重置视图", command=self.reset_view)
        self.reset_view_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.clear_selection_btn = ttk.Button(monitor_frame, text="清除选择", command=self.clear_selection)
        self.clear_selection_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 状态显示
        self.status_label = ttk.Label(monitor_frame, text="状态: 准备就绪")
        self.status_label.pack(side=tk.RIGHT)
        
        # 通道信息框架
        self.channels_frame = ttk.LabelFrame(main_frame, text="监控通道信息", padding=10)
        self.channels_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 初始显示无通道信息
        self.no_channels_label = ttk.Label(self.channels_frame, text="暂无监控通道，请点击'添加监控区域'创建通道")
        self.no_channels_label.pack(anchor=tk.W)
        
        # 图表框架
        if 'matplotlib' in self.deps:
            chart_frame = ttk.LabelFrame(main_frame, text="数据图表", padding=10)
            chart_frame.pack(fill=tk.BOTH, expand=True)
            
            # 创建matplotlib图形
            self.fig, self.ax = self.plt.subplots(figsize=(10, 5), dpi=100)
            self.ax.set_title("多通道实时数值监控 (支持拖动和点选)")
            self.ax.set_xlabel("时间")
            self.ax.set_ylabel("数值")
            self.ax.grid(True, alpha=0.3)
            
            # 设置时间格式化
            self.ax.xaxis.set_major_formatter(self.mdates.DateFormatter('%H:%M:%S'))
            self.ax.xaxis.set_major_locator(self.mdates.AutoDateLocator())
            
            # 嵌入画布
            self.canvas = self.FigureCanvasTkAgg(self.fig, chart_frame)
            self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # 连接事件
            self.canvas.mpl_connect('button_press_event', self.on_click)
            self.canvas.mpl_connect('motion_notify_event', self.on_motion)
            self.canvas.mpl_connect('button_release_event', self.on_release)
            self.canvas.mpl_connect('scroll_event', self.on_scroll)
        
        # 数据显示和保存框架
        data_frame = ttk.LabelFrame(main_frame, text="实时数据与导出", padding=10)
        data_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 当前值显示
        self.current_values_frame = ttk.Frame(data_frame)
        self.current_values_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.no_data_label = ttk.Label(self.current_values_frame, text="暂无数据")
        self.no_data_label.pack(anchor=tk.W)
        
        # 选中点信息显示
        self.selection_info_frame = ttk.Frame(data_frame)
        self.selection_info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.no_selection_label = ttk.Label(self.selection_info_frame, text="未选中任何数据点")
        self.no_selection_label.pack(anchor=tk.W)
        
        # 保存按钮
        save_frame = ttk.Frame(data_frame)
        save_frame.pack(fill=tk.X)
        
        self.save_btn = ttk.Button(save_frame, text="保存图表", command=self.save_chart)
        self.save_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.save_csv_btn = ttk.Button(save_frame, text="导出CSV", command=self.save_csv)
        self.save_csv_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        # 启动队列处理
        self.process_queue()
    
    def add_channel(self):
        """添加新的监控通道"""
        # 创建区域选择窗口
        region_window = tk.Toplevel(self.root)
        region_window.title("选择监控区域 - 通道 {}".format(len(self.channels) + 1))
        region_window.attributes("-fullscreen", True)
        region_window.attributes("-alpha", 0.7)  # 调整为更低的透明度，便于看到内容
        region_window.attributes("-topmost", True)
        region_window.configure(bg='gray')
        
        # 存储窗口引用
        self.region_windows.append(region_window)
        
        canvas = tk.Canvas(region_window, bg='gray', highlightthickness=0, cursor="crosshair")
        canvas.pack(fill=tk.BOTH, expand=True)
        
        start_x = start_y = end_x = end_y = None
        rect_id = None
        selection_border = None
        
        def on_mouse_down(event):
            nonlocal start_x, start_y, rect_id, selection_border
            start_x = event.x
            start_y = event.y
            
            # 创建选择矩形（半透明蓝色）
            rect_id = canvas.create_rectangle(start_x, start_y, start_x, start_y, 
                                            outline='red', width=3, fill='blue', stipple='gray50')
            
            # 创建选中效果边框（黄色实线）
            selection_border = canvas.create_rectangle(start_x-2, start_y-2, start_x+2, start_y+2,
                                                     outline='yellow', width=3, fill='')
        
        def on_mouse_move(event):
            nonlocal rect_id, selection_border
            if rect_id:
                canvas.delete(rect_id)
                canvas.delete(selection_border)
                
                # 重新绘制选择矩形
                rect_id = canvas.create_rectangle(start_x, start_y, event.x, event.y, 
                                                outline='red', width=3, fill='blue', stipple='gray50')
                
                # 更新选中效果边框
                x1, y1 = min(start_x, event.x), min(start_y, event.y)
                x2, y2 = max(start_x, event.x), max(start_y, event.y)
                selection_border = canvas.create_rectangle(x1-3, y1-3, x2+3, y2+3,
                                                         outline='yellow', width=3, fill='')
        
        def on_mouse_up(event):
            nonlocal end_x, end_y
            end_x = event.x
            end_y = event.y
            
            # 调整坐标以适应屏幕缩放
            x1, y1 = min(start_x, end_x), min(start_y, end_y)
            x2, y2 = max(start_x, end_x), max(start_y, end_y)
            
            if x2 - x1 > 10 and y2 - y1 > 10:
                # 创建新通道
                channel_index = len(self.channels)
                channel_name = f"通道 {channel_index + 1}"
                channel_data = {
                    'name': channel_name,
                    'rect': (x1, y1, x2, y2),
                    'times': [],
                    'values': [],
                    'color': self.get_channel_color(channel_index),
                    'visible': True,
                    'scatter': None,  # 用于存储散点对象引用
                    'region_window': region_window,  # 存储窗口引用
                    'selection_border': selection_border  # 存储边框引用
                }
                self.channels.append(channel_data)
                
                # 存储矩形引用
                self.region_rectangles[channel_index] = selection_border
                
                # 不关闭窗口，保持显示
                # 移除全屏属性，调整为正常窗口显示选中区域
                region_window.attributes("-fullscreen", False)
                region_window.geometry(f"{x2-x1+60}x{y2-y1+60}+{max(0, x1-30)}+{max(0, y1-30)}")
                region_window.title(f"监控区域 - {channel_name}")
                region_window.attributes("-alpha", 0.8)  # 保持一定透明度但可见
                region_window.attributes("-topmost", True)
                
                # 添加关闭按钮
                close_frame = ttk.Frame(region_window)
                close_frame.pack(fill=tk.X, padx=10, pady=5)
                
                close_btn = ttk.Button(close_frame, text="关闭区域显示", 
                                     command=lambda: self.close_region_window(channel_index))
                close_btn.pack(side=tk.RIGHT)
                
                # 更新通道信息显示
                info_label = ttk.Label(close_frame, text=channel_name, 
                                     foreground=channel_data['color'])
                info_label.pack(side=tk.LEFT)
                
                # 更新界面
                self.update_channels_display()
                self.update_channel_combo()
                self.start_btn.config(state=tk.NORMAL)
                self.status_label.config(text=f"状态: 已添加{channel_name}，可以开始监控")
            else:
                messagebox.showerror("错误", "选择区域太小，请重新选择")
                region_window.destroy()
                if region_window in self.region_windows:
                    self.region_windows.remove(region_window)
        
        # 添加操作提示
        tip_text = "拖拽鼠标选择监控区域 - 通道 {}".format(len(self.channels) + 1)
        canvas.create_text(region_window.winfo_screenwidth() // 2, 50, 
                          text=tip_text, fill='white', font=("Arial", 16, "bold"))
        canvas.create_text(region_window.winfo_screenwidth() // 2, 80, 
                          text="释放鼠标完成选择", fill='yellow', font=("Arial", 12))
        canvas.create_text(region_window.winfo_screenwidth() // 2, 110, 
                          text="黄色边框表示选中区域", fill='yellow', font=("Arial", 10))
        
        canvas.bind("<Button-1>", on_mouse_down)
        canvas.bind("<B1-Motion>", on_mouse_move)
        canvas.bind("<ButtonRelease-1>", on_mouse_up)
        
        # ESC键取消选择
        def cancel_selection(event):
            region_window.destroy()
            if region_window in self.region_windows:
                self.region_windows.remove(region_window)
        
        region_window.bind("<Escape>", cancel_selection)
        region_window.focus_set()
    
    def close_region_window(self, channel_index):
        """关闭区域显示窗口"""
        if 0 <= channel_index < len(self.channels):
            channel = self.channels[channel_index]
            if 'region_window' in channel and channel['region_window']:
                channel['region_window'].destroy()
                if channel['region_window'] in self.region_windows:
                    self.region_windows.remove(channel['region_window'])
                channel['region_window'] = None
    
    def show_all_regions(self):
        """显示所有监控区域"""
        for i, channel in enumerate(self.channels):
            if not channel.get('region_window'):
                # 重新创建区域显示窗口
                self.create_region_display_window(i)
    
    def hide_all_regions(self):
        """隐藏所有监控区域"""
        for channel in self.channels:
            if channel.get('region_window'):
                self.close_region_window(self.channels.index(channel))
    
    def create_region_display_window(self, channel_index):
        """为指定通道创建区域显示窗口"""
        if 0 <= channel_index < len(self.channels):
            channel = self.channels[channel_index]
            rect = channel['rect']
            x1, y1, x2, y2 = rect
            
            # 创建窗口
            region_window = tk.Toplevel(self.root)
            region_window.title(f"监控区域 - {channel['name']}")
            region_window.geometry(f"{x2-x1+60}x{y2-y1+60}+{max(0, x1-30)}+{max(0, y1-30)}")
            region_window.attributes("-alpha", 0.8)
            region_window.attributes("-topmost", True)
            region_window.configure(bg='gray')
            
            # 存储窗口引用
            channel['region_window'] = region_window
            self.region_windows.append(region_window)
            
            canvas = tk.Canvas(region_window, bg='gray', highlightthickness=0)
            canvas.pack(fill=tk.BOTH, expand=True)
            
            # 绘制选中区域
            canvas.create_rectangle(30, 30, x2-x1+30, y2-y1+30, 
                                  outline='red', width=3, fill='blue', stipple='gray50')
            
            # 绘制选中边框
            selection_border = canvas.create_rectangle(27, 27, x2-x1+33, y2-y1+33,
                                                     outline='yellow', width=3, fill='')
            channel['selection_border'] = selection_border
            self.region_rectangles[channel_index] = selection_border
            
            # 添加关闭按钮
            close_frame = ttk.Frame(region_window)
            close_frame.pack(fill=tk.X, padx=10, pady=5)
            
            close_btn = ttk.Button(close_frame, text="关闭区域显示", 
                                 command=lambda: self.close_region_window(channel_index))
            close_btn.pack(side=tk.RIGHT)
            
            # 显示通道信息
            info_label = ttk.Label(close_frame, text=channel['name'], 
                                 foreground=channel['color'])
            info_label.pack(side=tk.LEFT)
    
    def remove_channel(self):
        """删除当前通道"""
        if not self.channels:
            return
            
        current_index = self.active_channel_index
        channel_name = self.channels[current_index]['name']
        
        if messagebox.askyesno("确认删除", f"确定要删除{channel_name}吗？"):
            # 关闭对应的区域窗口
            self.close_region_window(current_index)
            
            # 从列表中移除
            self.channels.pop(current_index)
            
            # 更新region_rectangles字典
            new_rectangles = {}
            for i, channel in enumerate(self.channels):
                if i >= current_index:
                    new_rectangles[i] = self.region_rectangles.get(i+1)
            self.region_rectangles = new_rectangles
            
            # 更新活动通道索引
            if self.channels:
                self.active_channel_index = min(self.active_channel_index, len(self.channels) - 1)
            else:
                self.active_channel_index = 0
                self.start_btn.config(state=tk.DISABLED)
            
            self.update_channels_display()
            self.update_channel_combo()
            self.update_chart()
    
    def on_channel_change(self, event):
        """通道选择改变事件"""
        if self.channels:
            selected_text = self.channel_var.get()
            for i, channel in enumerate(self.channels):
                if channel['name'] == selected_text:
                    self.active_channel_index = i
                    break
            self.update_channels_display()
            self.update_chart()
    
    def update_channels_display(self):
        """更新通道信息显示"""
        # 清空现有显示
        for widget in self.channels_frame.winfo_children():
            widget.destroy()
        
        if not self.channels:
            self.no_channels_label = ttk.Label(self.channels_frame, text="暂无监控通道，请点击'添加监控区域'创建通道")
            self.no_channels_label.pack(anchor=tk.W)
            self.remove_channel_btn.config(state=tk.DISABLED)
            self.show_regions_btn.config(state=tk.DISABLED)
            self.hide_regions_btn.config(state=tk.DISABLED)
            return
        
        self.remove_channel_btn.config(state=tk.NORMAL)
        self.show_regions_btn.config(state=tk.NORMAL)
        self.hide_regions_btn.config(state=tk.NORMAL)
        
        # 显示所有通道信息
        for i, channel in enumerate(self.channels):
            channel_frame = ttk.Frame(self.channels_frame)
            channel_frame.pack(fill=tk.X, pady=2)
            
            # 通道名称和区域信息
            rect = channel['rect']
            data_count = len(channel['values'])
            latest_value = channel['values'][-1] if channel['values'] else '无数据'
            region_status = "显示中" if channel.get('region_window') else "隐藏"
            info_text = f"{channel['name']}: 区域({rect[0]}, {rect[1]}) - ({rect[2]}, {rect[3]}) | 数据点: {data_count} | 最新值: {latest_value} | 区域: {region_status}"
            
            channel_label = ttk.Label(channel_frame, text=info_text, 
                                    foreground=channel['color'] if i == self.active_channel_index else 'black')
            channel_label.pack(side=tk.LEFT)
            
            # 区域显示控制
            region_frame = ttk.Frame(channel_frame)
            region_frame.pack(side=tk.RIGHT)
            
            if channel.get('region_window'):
                hide_btn = ttk.Button(region_frame, text="隐藏区域", 
                                    command=lambda idx=i: self.close_region_window(idx))
                hide_btn.pack(side=tk.RIGHT, padx=(5, 0))
            else:
                show_btn = ttk.Button(region_frame, text="显示区域", 
                                    command=lambda idx=i: self.create_region_display_window(idx))
                show_btn.pack(side=tk.RIGHT, padx=(5, 0))
            
            # 可见性控制
            visible_var = tk.BooleanVar(value=channel['visible'])
            visible_cb = ttk.Checkbutton(region_frame, text="显示曲线", variable=visible_var,
                                       command=lambda idx=i, var=visible_var: self.toggle_channel_visibility(idx, var))
            visible_cb.pack(side=tk.RIGHT, padx=(10, 5))
    
    def update_channel_combo(self):
        """更新通道选择下拉框"""
        if self.channels:
            channel_names = [channel['name'] for channel in self.channels]
            self.channel_combo['values'] = channel_names
            self.channel_var.set(self.channels[self.active_channel_index]['name'])
        else:
            self.channel_combo['values'] = []
            self.channel_var.set("无通道")
    
    def toggle_channel_visibility(self, channel_index, var):
        """切换通道可见性"""
        self.channels[channel_index]['visible'] = var.get()
        self.update_chart()
    
    def get_channel_color(self, index):
        """获取通道颜色"""
        colors = ['blue', 'red', 'green', 'orange', 'purple', 'brown', 'pink', 'gray']
        return colors[index % len(colors)]
    
    def toggle_monitoring(self):
        """切换监控状态"""
        if not self.monitoring:
            self.start_monitoring()
        else:
            self.stop_monitoring()
    
    def start_monitoring(self):
        """开始监控"""
        if not self.channels:
            messagebox.showerror("错误", "请先添加监控通道")
            return
            
        self.monitoring = True
        self.start_btn.config(text="停止监控")
        self.add_channel_btn.config(state=tk.DISABLED)
        self.remove_channel_btn.config(state=tk.DISABLED)
        self.show_regions_btn.config(state=tk.DISABLED)
        self.hide_regions_btn.config(state=tk.DISABLED)
        self.interval = self.interval_var.get()
        self.status_label.config(text="状态: 监控中...")
        
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()
    
    def stop_monitoring(self):
        """停止监控"""
        self.monitoring = False
        self.start_btn.config(text="开始监控")
        self.add_channel_btn.config(state=tk.NORMAL)
        if self.channels:
            self.remove_channel_btn.config(state=tk.NORMAL)
            self.show_regions_btn.config(state=tk.NORMAL)
            self.hide_regions_btn.config(state=tk.NORMAL)
        self.status_label.config(text="状态: 已停止")
    
    def monitor_loop(self):
        """监控主循环 - 多通道版本"""
        while self.monitoring:
            try:
                if self.channels:
                    timestamp = datetime.now()
                    
                    # 遍历所有通道进行截图和OCR
                    for i, channel in enumerate(self.channels):
                        rect = channel['rect']
                        screenshot = ImageGrab.grab(bbox=rect)
                        text = self.pytesseract.image_to_string(screenshot, lang='chi_sim+eng')
                        value = self.parse_value(text)
                        
                        # 只有成功解析到有效数值才放入队列
                        if value is not None and value >= 0:
                            # 线程安全地放入队列
                            try:
                                self.data_queue.put((i, timestamp, value), block=False)
                                print(f"[DEBUG] 通道{i}解析成功: {value:.2f} at {timestamp.strftime('%H:%M:%S')}")
                            except queue.Full:
                                print(f"[DEBUG] 队列已满，丢弃数据: 通道{i} - {value:.2f}")
                        else:
                            print(f"[DEBUG] 通道{i}解析失败或无效值: {text.strip()}")
                
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
        """处理数据队列 - 多通道版本"""
        # 确保在主线程中执行UI更新
        if threading.current_thread() != threading.main_thread():
            return
            
        try:
            new_data_added = False
            
            # 批量处理队列中的所有数据
            while True:
                try:
                    channel_index, timestamp, value = self.data_queue.get_nowait()
                    
                    # 确保通道索引有效
                    if 0 <= channel_index < len(self.channels):
                        # 只添加有效的数据点
                        if value is not None and value >= 0:
                            self.channels[channel_index]['times'].append(timestamp)
                            self.channels[channel_index]['values'].append(value)
                            new_data_added = True
                            
                            # 限制数据点数量
                            if len(self.channels[channel_index]['times']) > self.max_points:
                                excess = len(self.channels[channel_index]['times']) - self.max_points
                                self.channels[channel_index]['times'] = self.channels[channel_index]['times'][excess:]
                                self.channels[channel_index]['values'] = self.channels[channel_index]['values'][excess:]
                    
                except queue.Empty:
                    break
            
            # 如果有新数据添加，更新显示
            if new_data_added:
                self.update_current_values_display()
                self.update_channels_display()
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
    
    def update_current_values_display(self):
        """更新当前值显示"""
        # 清空现有显示
        for widget in self.current_values_frame.winfo_children():
            widget.destroy()
        
        if not any(len(channel['values']) > 0 for channel in self.channels):
            self.no_data_label = ttk.Label(self.current_values_frame, text="暂无数据")
            self.no_data_label.pack(anchor=tk.W)
            return
        
        # 显示所有有数据的通道的当前值
        for i, channel in enumerate(self.channels):
            if channel['values']:
                current_value = channel['values'][-1]
                value_label = ttk.Label(self.current_values_frame, 
                                      text=f"{channel['name']}: {current_value:.2f}",
                                      font=("Arial", 10, "bold"),
                                      foreground=channel['color'])
                value_label.pack(side=tk.LEFT, padx=(0, 20))
    
    def update_selection_info(self):
        """更新选中点信息显示"""
        # 清空现有显示
        for widget in self.selection_info_frame.winfo_children():
            widget.destroy()
        
        if not self.selected_points:
            self.no_selection_label = ttk.Label(self.selection_info_frame, text="未选中任何数据点")
            self.no_selection_label.pack(anchor=tk.W)
            return
        
        # 显示选中点信息
        info_text = f"选中 {len(self.selected_points)} 个数据点: "
        for i, (channel_idx, point_idx) in enumerate(self.selected_points):
            if i < 3:  # 只显示前3个选中点
                channel = self.channels[channel_idx]
                time_str = channel['times'][point_idx].strftime('%H:%M:%S')
                value = channel['values'][point_idx]
                info_text += f"{channel['name']}[{time_str}: {value:.2f}] "
            elif i == 3:
                info_text += "..."
                break
        
        selection_label = ttk.Label(self.selection_info_frame, text=info_text, foreground='blue')
        selection_label.pack(anchor=tk.W)
    
    def update_chart(self):
        """更新图表 - 多通道版本"""
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
            self.ax.clear()
            
            has_visible_data = False
            
            # 绘制所有可见通道的数据
            for i, channel in enumerate(self.channels):
                if not channel['visible'] or not channel['times'] or not channel['values']:
                    continue
                    
                has_visible_data = True
                times = channel['times']
                values = channel['values']
                
                # 绘制曲线
                line = self.ax.plot(times, values, '-', color=channel['color'], 
                                  linewidth=1.5, alpha=0.6, label=channel['name'])[0]
                
                # 绘制散点，并为每个点设置picker属性
                scatter = self.ax.scatter(times, values, c=channel['color'], s=30, alpha=0.8, 
                                        zorder=5, picker=True, pickradius=5)
                
                # 存储散点对象引用
                channel['scatter'] = scatter
                
                # 高亮选中的点
                for channel_idx, point_idx in self.selected_points:
                    if channel_idx == i and point_idx < len(times):
                        self.ax.scatter([times[point_idx]], [values[point_idx]], 
                                      c='gold', s=100, alpha=1.0, zorder=6, 
                                      edgecolors='red', linewidths=2)
            
            if not has_visible_data:
                self.ax.text(0.5, 0.5, '暂无数据或所有通道已隐藏', 
                           transform=self.ax.transAxes, ha='center', va='center', fontsize=12)
                self.ax.set_xlim(0, 1)
                self.ax.set_ylim(0, 1)
            else:
                # 设置图表样式
                self.ax.set_title("多通道实时数值监控 (支持拖动和点选)")
                self.ax.set_xlabel("时间")
                self.ax.set_ylabel("数值")
                self.ax.grid(True, alpha=0.3)
                self.ax.legend(loc='upper right')
                
                # 设置时间格式化
                self.ax.xaxis.set_major_formatter(self.mdates.DateFormatter('%H:%M:%S'))
                self.ax.xaxis.set_major_locator(self.mdates.AutoDateLocator())
                
                # 自动调整Y轴范围
                all_values = []
                for channel in self.channels:
                    if channel['visible'] and channel['values']:
                        all_values.extend(channel['values'])
                
                if all_values:
                    y_min, y_max = min(all_values), max(all_values)
                    margin = (y_max - y_min) * 0.1 if y_max > y_min else 1
                    self.ax.set_ylim(max(0, y_min - margin), y_max + margin)
            
            # 安全地刷新画布
            if hasattr(self, 'canvas'):
                self.canvas.draw()
            
        except Exception as e:
            print(f"[DEBUG] 图表更新异常: {e}")
            import traceback
            traceback.print_exc()
    
    # 图表交互功能
    def on_click(self, event):
        """鼠标点击事件处理"""
        if event.inaxes != self.ax:
            return
        
        if event.button == 1:  # 左键
            # 点选功能
            if event.dblclick:  # 双击选择点
                self.select_point(event)
            else:  # 单击开始拖动
                self.drag_start = (event.xdata, event.ydata)
                self.original_xlim = self.ax.get_xlim()
                self.original_ylim = self.ax.get_ylim()
        
        elif event.button == 3:  # 右键
            # 右键取消选择
            self.selected_points.clear()
            self.update_selection_info()
            self.update_chart()
    
    def on_motion(self, event):
        """鼠标移动事件处理"""
        if event.inaxes != self.ax or self.drag_start is None:
            return
        
        if event.button == 1:  # 左键拖动
            dx = event.xdata - self.drag_start[0]
            dy = event.ydata - self.drag_start[1]
            
            xlim = self.ax.get_xlim()
            ylim = self.ax.get_ylim()
            
            new_xlim = [xlim[0] - dx, xlim[1] - dx]
            new_ylim = [ylim[0] - dy, ylim[1] - dy]
            
            self.ax.set_xlim(new_xlim)
            self.ax.set_ylim(new_ylim)
            self.canvas.draw()
    
    def on_release(self, event):
        """鼠标释放事件处理"""
        self.drag_start = None
    
    def on_scroll(self, event):
        """鼠标滚轮缩放事件"""
        if event.inaxes != self.ax:
            return
        
        scale_factor = 1.1 if event.button == 'up' else 0.9
        
        # X轴缩放
        xlim = self.ax.get_xlim()
        x_center = (xlim[0] + xlim[1]) / 2
        x_width = (xlim[1] - xlim[0]) * scale_factor
        new_xlim = [x_center - x_width/2, x_center + x_width/2]
        
        # Y轴缩放
        ylim = self.ax.get_ylim()
        y_center = (ylim[0] + ylim[1]) / 2
        y_height = (ylim[1] - ylim[0]) * scale_factor
        new_ylim = [y_center - y_height/2, y_center + y_height/2]
        
        self.ax.set_xlim(new_xlim)
        self.ax.set_ylim(new_ylim)
        self.canvas.draw()
    
    def select_point(self, event):
        """选择数据点"""
        if not self.channels:
            return
        
        # 查找最近的数据点
        min_distance = float('inf')
        selected_channel = -1
        selected_point = -1
        
        for i, channel in enumerate(self.channels):
            if not channel['visible'] or not channel['times'] or not channel['values']:
                continue
            
            times = channel['times']
            values = channel['values']
            
            for j, (t, v) in enumerate(zip(times, values)):
                # 计算距离
                dist = ((self.mdates.date2num(t) - event.xdata) ** 2 + 
                       (v - event.ydata) ** 2)
                
                if dist < min_distance:
                    min_distance = dist
                    selected_channel = i
                    selected_point = j
        
        # 如果找到足够近的点
        if min_distance < 0.001:  # 距离阈值
            point_key = (selected_channel, selected_point)
            
            # 如果按住Ctrl键，可以多选
            if event.key and 'control' in event.key:
                if point_key in self.selected_points:
                    self.selected_points.remove(point_key)
                else:
                    self.selected_points.append(point_key)
            else:
                # 单选模式
                self.selected_points = [point_key]
            
            self.update_selection_info()
            self.update_chart()
    
    def zoom_in(self):
        """放大图表"""
        self.zoom_chart(0.8)
    
    def zoom_out(self):
        """缩小图表"""
        self.zoom_chart(1.2)
    
    def zoom_chart(self, factor):
        """缩放图表"""
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()
        
        x_center = (xlim[0] + xlim[1]) / 2
        y_center = (ylim[0] + ylim[1]) / 2
        
        x_width = (xlim[1] - xlim[0]) * factor
        y_height = (ylim[1] - ylim[0]) * factor
        
        self.ax.set_xlim([x_center - x_width/2, x_center + x_width/2])
        self.ax.set_ylim([y_center - y_height/2, y_center + y_height/2])
        self.canvas.draw()
    
    def reset_view(self):
        """重置视图到自动范围"""
        self.ax.relim()
        self.ax.autoscale_view()
        self.canvas.draw()
    
    def clear_selection(self):
        """清除所有选中的点"""
        self.selected_points.clear()
        self.update_selection_info()
        self.update_chart()
    
    def clear_all_data(self):
        """清空所有通道数据"""
        if not self.channels:
            return
            
        if messagebox.askyesno("确认清空", "确定要清空所有通道的数据吗？"):
            # 清空数据队列
            while True:
                try:
                    self.data_queue.get_nowait()
                except queue.Empty:
                    break
            
            # 清空所有通道数据
            for channel in self.channels:
                channel['times'].clear()
                channel['values'].clear()
            
            # 清除选择
            self.selected_points.clear()
            
            # 更新显示
            self.update_current_values_display()
            self.update_selection_info()
            self.update_channels_display()
            self.update_chart()
            
            print("[DEBUG] 所有通道数据已清空")
    
    def save_chart(self):
        """保存图表"""
        if not hasattr(self, 'fig') or not any(len(channel['values']) > 0 for channel in self.channels):
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
        """导出CSV - 多通道版本"""
        if not any(len(channel['values']) > 0 for channel in self.channels):
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
                    
                    # 写入表头
                    headers = ['时间']
                    for channel in self.channels:
                        if channel['values']:
                            headers.append(channel['name'])
                    
                    writer.writerow(headers)
                    
                    # 找到所有通道的最大数据长度
                    max_length = max(len(channel['times']) for channel in self.channels if channel['values'])
                    
                    # 写入数据
                    for i in range(max_length):
                        row = []
                        
                        # 时间列（使用第一个有数据的通道的时间）
                        for channel in self.channels:
                            if channel['times']:
                                if i < len(channel['times']):
                                    row.append(channel['times'][i].strftime('%Y-%m-%d %H:%M:%S'))
                                else:
                                    row.append('')
                                break
                        
                        # 数值列
                        for channel in self.channels:
                            if channel['values']:
                                if i < len(channel['values']):
                                    row.append(f"{channel['values'][i]:.2f}")
                                else:
                                    row.append('')
                        
                        writer.writerow(row)
                        
                messagebox.showinfo("成功", f"数据已导出到: {filename}")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {e}")
    
    def run(self):
        """启动程序"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ScreenOCROrMonitor()
    app.run()
