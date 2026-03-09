import os
import sys
import ctypes
from pathlib import Path
def get_application_path():
    """หา Path ของโฟลเดอร์ที่โปรแกรมกำลังทำงานอยู่"""
    if getattr(sys, 'frozen', False):
        return Path(os.path.dirname(sys.executable))
    else:
        return Path(os.path.dirname(os.path.abspath(__file__)))
APPLICATION_PATH = get_application_path()
print(f"Application base path detected at: {APPLICATION_PATH}")

VLC_DIRECTORY = APPLICATION_PATH / 'vlc-64'
if VLC_DIRECTORY.is_dir():
    print(f"Found 'vlc-64' directory at: {VLC_DIRECTORY}")
    try:
        os.add_dll_directory(str(VLC_DIRECTORY))
        print("Successfully added VLC path to DLL search directory.")
    except Exception as e:
        print(f"Warning: Could not add VLC path. Error: {e}")
else:
    print("!!! CRITICAL WARNING: 'vlc-64' DIRECTORY NOT FOUND !!!")

if hasattr(sys, '_MEIPASS'):
    base_path = Path(sys._MEIPASS)
    vlc_plugin_path = base_path / 'vlc-64' / 'plugins'
    if vlc_plugin_path.exists():
        os.environ['VLC_PLUGIN_PATH'] = str(vlc_plugin_path)
    fonts_path = base_path / 'fonts'
    if fonts_path.exists() and fonts_path.is_dir():
        FR_PRIVATE = 0x10
        gdi32 = ctypes.WinDLL('gdi32')
        for font_file in os.listdir(fonts_path):
            if font_file.lower().endswith(('.ttf', '.otf')):
                font_full_path = str(fonts_path / font_file)
                path_buffer = ctypes.create_unicode_buffer(font_full_path)
                if gdi32.AddFontResourceExW(path_buffer, FR_PRIVATE, 0) != 0:
                    print(f"Font loaded: {font_file}")
                else:
                    print(f"Warning: Failed to load font {font_file}")
    tkdnd_path = base_path / 'tkdnd'
    if tkdnd_path.exists():
        os.environ['TKDND_LIBRARY'] = str(tkdnd_path)
import random
import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import time
from mutagen.mp3 import MP3
from mutagen.mp4 import MP4
from mutagen.wave import WAVE
import json
from tkinterdnd2 import DND_FILES, TkinterDnD
try:
    import pystray
    from PIL import Image, ImageDraw
except ImportError:
    pass

try:
    import pygetwindow as gw
except ImportError:
    pass

import winreg

try:
    import vlc
    print("Successfully imported 'vlc'.")
except ImportError as e:
    print(f"FATAL ERROR: Failed to import 'vlc'. Error: {e}")
    root_for_error = tk.Tk()
    root_for_error.withdraw()
    messagebox.showerror("VLC Error",
                         "ไม่สามารถโหลด VLC ได้\nกรุณาตรวจสอบความถูกต้อง")
    sys.exit(1)

def load_font(font_path):
    try:
        if not os.path.exists(font_path): return False
        FR_PRIVATE = 0x10
        gdi32 = ctypes.WinDLL('gdi32')
        path_buffer = ctypes.create_unicode_buffer(str(font_path))
        result = gdi32.AddFontResourceExW(path_buffer, FR_PRIVATE, 0)
        return result != 0
    except Exception as e:
        print(f"Error loading font {font_path}: {e}")
        return False
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, bg="#f0f0f0", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollable_frame = ttk.Frame(self.canvas, style='TFrame')
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.bind_all("<MouseWheel>", self.on_mousewheel)
    def on_frame_configure(self, event):
        """อัปเดต scroll region ให้ครอบคลุม Frame ด้านในทั้งหมด"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    def on_canvas_configure(self, event):
        """ปรับความกว้างของ Frame ด้านในให้เท่ากับความกว้างของ Canvas"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    def on_mousewheel(self, event):
        """เลื่อน Canvas ตามการหมุนของ Mouse Wheel"""
        widget = self.canvas.winfo_containing(event.x_root, event.y_root)
        if widget:
            # เช็คว่าเราอยู่บนออบเจกต์ใน Canvas เพื่อให้สามารถเลื่อนแม้เมาส์จะทับ Label หรือปุ่มอยู่
            if str(widget).startswith(str(self.canvas)) or widget == self.canvas:
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
class AudioSystemApp:
    def __init__(self, root):
        self.root = root
        self.root.title("โปรแกรมเล่นสื่ออัตโนมัติ ออกแบบเเละพัฒนาโดย Phenphet.K")
        try:
            icon_path = APPLICATION_PATH / 'icon.ico'
            if icon_path.exists():
                self.root.iconbitmap(icon_path)
            else:
                print("Warning: Icon file 'icon.ico' not found.")
        except Exception as e:
            print(f"Warning: Could not set window icon. Error: {e}")
            
        # ทำให้หน้าต่างพอดีกับหน้าจออัตโนมัติ
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # ตั้งค่าขนาดเริ่มต้นให้เกือบเต็มจอ (เผื่อ Taskbar) แต่ไม่ต้อง Zoomed ผูกมัดตั้งแต่แรก
        window_width = int(screen_width * 0.95)
        window_height = int(screen_height * 0.9)
        
        # จัดตำแหน่งให้อยู่ตรงกลางจอ
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # ตั้งค่าขั้นต่ำเพื่อไม่ให้ UI บีบจนพัง
        self.root.minsize(int(screen_width * 0.7), int(screen_height * 0.7))
        self.root.state('zoomed') # ขยายเต็มจออัตโนมัติเมื่อเปิด
        self.root.configure(bg="#f0f0f0")
        self.font_normal = ("TH Sarabun New", 15)
        self.font_bold = ("TH Sarabun New", 17, "bold")
        self.font_header = ("TH Sarabun New", 20, "bold")
        self.vlc_instance = vlc.Instance('--audio-filter=equalizer', '--no-xlib', '--fullscreen', '--video-on-top', '--aout=waveout')
        self.main_player = self.vlc_instance.media_player_new()
        self.interrupt_player = self.vlc_instance.media_player_new()
        self.status_var = tk.StringVar(value="พร้อมใช้งาน")
        self.time_var = tk.StringVar(value="")
        self.current_directory = tk.StringVar(value="ยังไม่ได้เลือกโฟลเดอร์ หรือนำเข้าไฟล์สื่อ")
        self.total_duration_var = tk.StringVar(value="ระยะเวลารวม: 00:00:00")
        self.end_time_var = tk.StringVar(value="เวลาสิ้นสุดโดยประมาณ: --:--:--")
        self.media_list = []
        self.current_media_item = None
        self.is_playing_main = False
        self.is_playing_interrupt = False
        self.interrupted_media_item = None
        self.play_random_mode = tk.BooleanVar(value=True)
        self.shutdown_on_finish = tk.BooleanVar(value=False)
        self.loop_media_var = tk.BooleanVar(value=False)
        self.stop_threads = False
        self.eq_values = [0.0] * 10
        self.main_volume = 100
        self.minimize_to_tray_var = tk.BooleanVar(value=True)
        self.auto_start_var = tk.BooleanVar(value=False)
        self.stats = {"main_played": 0, "interrupt_played": 0}
        
        appdata_path = Path(os.getenv('APPDATA'))
        app_settings_dir = appdata_path / "RadioSystemApp"
        app_settings_dir.mkdir(parents=True, exist_ok=True)
        self.settings_file = app_settings_dir / "radiosettings.json"
        self.create_ui()
        self.load_settings()
        self.start_threads()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.setup_keyboard_shortcuts()
    def create_ui(self):
        """สร้างส่วนติดต่อผู้ใช้ทั้งหมดของโปรแกรม"""
        top_frame = tk.Frame(self.root, bg="#4CAF50", height=60)
        top_frame.pack(fill=tk.X)
        tk.Label(top_frame, textvariable=self.status_var, font=self.font_header, bg="#4CAF50", fg="white").pack(
            side=tk.LEFT, padx=20, pady=10)
        tk.Label(top_frame, textvariable=self.time_var, font=self.font_header, bg="#4CAF50", fg="white").pack(
            side=tk.RIGHT, padx=20, pady=10)
            
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_main = ttk.Frame(self.notebook, style='TFrame')
        self.tab_dashboard = ttk.Frame(self.notebook, style='TFrame')
        
        self.notebook.add(self.tab_main, text="   เครื่องเล่นหลัก   ")
        self.notebook.add(self.tab_dashboard, text="   แดชบอร์ดสรุปผล   ")

        main_paned_window = ttk.PanedWindow(self.tab_main, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)
        left_frame = ttk.Frame(main_paned_window, style='TFrame')
        main_paned_window.add(left_frame, weight=3)
        right_frame = ttk.Frame(main_paned_window, style='TFrame')
        main_paned_window.add(right_frame, weight=2)
        scrollable_settings_panel = ScrollableFrame(right_frame)
        scrollable_settings_panel.pack(fill=tk.BOTH, expand=True)
        self.create_settings_panel(scrollable_settings_panel.scrollable_frame)
        media_frame = ttk.LabelFrame(left_frame,
                                     text="รายการสื่อหลัก (หรือลากไฟล์มาวางที่นี่ เลือกและกดปุ่ม Del เพื่อนำออกจากรายการ)",
                                     style='TLabelframe')
        media_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.create_media_list(media_frame)
        main_playback_frame = ttk.LabelFrame(left_frame, text="ควบคุมการเล่นสื่อหลัก", style='TLabelframe')
        main_playback_frame.pack(fill=tk.X, padx=5, pady=5)
        self.create_main_playback_controls(main_playback_frame)
        bottom_status_frame = tk.Frame(left_frame, bg="#f0f0f0")
        bottom_status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        tk.Label(bottom_status_frame, textvariable=self.total_duration_var, font=self.font_normal, bg="#f0f0f0").pack(
            side=tk.LEFT, padx=10)
        tk.Label(bottom_status_frame, textvariable=self.end_time_var, font=self.font_normal, bg="#f0f0f0").pack(
            side=tk.RIGHT, padx=10)
            
        self.create_dashboard_ui()

        style = ttk.Style()
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabelframe', background='#f0f0f0', borderwidth=1)
        style.configure('TLabelframe.Label', background='#f0f0f0', font=self.font_bold)
        style.configure("Treeview", font=self.font_normal, rowheight=30)
        style.configure("Treeview.Heading", font=self.font_bold)
    def create_settings_panel(self, parent_frame):
        main_settings_frame = ttk.LabelFrame(parent_frame, text="การตั้งค่าสื่อหลัก", style='TLabelframe')
        main_settings_frame.pack(fill=tk.X, padx=5, pady=5, ipady=5)

        dir_frame = tk.Frame(main_settings_frame, bg="#f0f0f0")
        dir_frame.pack(fill=tk.X, padx=10, pady=5)
        select_btn = tk.Button(dir_frame, text="เลือกโฟลเดอร์", command=self.select_directory, bg="#2196F3", fg="white",
                               font=self.font_normal)
        select_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        tk.Label(dir_frame, textvariable=self.current_directory, font=self.font_normal, bg="white", relief=tk.SUNKEN,
                 anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        mode_frame = tk.Frame(main_settings_frame, bg="#f0f0f0")
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        self.random_btn = tk.Button(mode_frame, text="เล่นแบบสุ่ม", command=self.set_random_mode, font=self.font_normal)
        self.random_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.sequential_btn = tk.Button(mode_frame, text="เล่นแบบเรียงลำดับ", command=self.set_sequential_mode,
                                        font=self.font_normal)
        self.sequential_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.update_play_mode_buttons()
        self.main_schedule_entries = []
        for i in range(3):
            self.main_schedule_entries.append(self.create_schedule_row(main_settings_frame, i, "เล่นรอบหลัก"))
        audio_settings_frame = ttk.LabelFrame(parent_frame, text="ตั้งค่าเสียง (EQ)", style='TLabelframe')
        audio_settings_frame.pack(fill=tk.X, padx=5, pady=10, ipady=5)
        self.create_audio_controls(audio_settings_frame)
        interrupt_frame = ttk.LabelFrame(parent_frame, text="ตั้งเวลาเล่นไฟล์/คั่นรายการ (Interrupt)",
                                         style='TLabelframe')
        interrupt_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=10, ipady=5)
        self.interrupt_schedule_entries = []
        for i in range(3):
            self.interrupt_schedule_entries.append(self.create_interrupt_row(interrupt_frame, i))
    def create_schedule_row(self, parent, index, label_prefix):
        row_frame = tk.Frame(parent, bg="#f0f0f0")
        row_frame.pack(fill=tk.X, padx=10, pady=3)
        tk.Label(row_frame, text=f"{label_prefix} {index + 1}:", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT,
                                                                                                           padx=5)
        open_hour = ttk.Combobox(row_frame, values=[f"{h:02d}" for h in range(24)], width=3, font=self.font_normal, state="readonly")
        open_hour.set("00")
        open_hour.pack(side=tk.LEFT)
        tk.Label(row_frame, text=":", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT)
        open_minute = ttk.Combobox(row_frame, values=[f"{m:02d}" for m in range(60)], width=3, font=self.font_normal, state="readonly")
        open_minute.set("00")
        open_minute.pack(side=tk.LEFT)
        tk.Label(row_frame, text=" ปิดเวลา ", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        close_hour = ttk.Combobox(row_frame, values=[f"{h:02d}" for h in range(24)], width=3, font=self.font_normal, state="readonly")
        close_hour.set("00")
        close_hour.pack(side=tk.LEFT)
        tk.Label(row_frame, text=":", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT)
        close_minute = ttk.Combobox(row_frame, values=[f"{m:02d}" for m in range(60)], width=3, font=self.font_normal, state="readonly")
        close_minute.set("00")
        close_minute.pack(side=tk.LEFT)

        return {"open_hour": open_hour, "open_minute": open_minute, "close_hour": close_hour,
                "close_minute": close_minute}
    def create_interrupt_row(self, parent, index):
        row_frame = tk.Frame(parent, bg="#f0f0f0")
        row_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(row_frame, text=f"รอบ {index + 1}", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT, anchor='n',
                                                                                               padx=5)
        details_frame = tk.Frame(row_frame, bg="#f0f0f0")
        details_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
        file_frame = tk.Frame(details_frame, bg="#f0f0f0")
        file_frame.pack(fill=tk.X)
        file_path_var = tk.StringVar(value="ยังไม่ได้เลือกไฟล์")
        select_file_btn = tk.Button(file_frame, text="เลือกไฟล์คั่นรายการ", font=self.font_normal,
                                    command=lambda i=index: self.select_interrupt_file(i))
        select_file_btn.pack(side=tk.LEFT, padx=5)
        tk.Label(file_frame, textvariable=file_path_var, font=self.font_normal, bg="white", relief=tk.SUNKEN, width=20,
                 anchor='w').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        time_frame = tk.Frame(details_frame, bg="#f0f0f0")
        time_frame.pack(fill=tk.X, pady=3)
        open_status_var = tk.StringVar(value="")
        tk.Label(time_frame, text="เปิด:", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT)
        open_hour = ttk.Combobox(time_frame, values=[f"{h:02d}" for h in range(24)], width=3, font=self.font_normal, state="readonly")
        open_hour.set("00")
        open_hour.pack(side=tk.LEFT)
        tk.Label(time_frame, text=":", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT)
        open_minute = ttk.Combobox(time_frame, values=[f"{m:02d}" for m in range(60)], width=3, font=self.font_normal, state="readonly")
        open_minute.set("00")
        open_minute.pack(side=tk.LEFT)
        tk.Label(time_frame, textvariable=open_status_var, font=(self.font_normal[0], 16, 'bold'), fg="green",
                 bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        close_status_var = tk.StringVar(value="")
        tk.Label(time_frame, text="ปิด:", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT)
        close_hour = ttk.Combobox(time_frame, values=[f"{h:02d}" for h in range(24)], width=3, font=self.font_normal, state="readonly")
        close_hour.set("00")
        close_hour.pack(side=tk.LEFT)
        tk.Label(time_frame, text=":", font=self.font_normal, bg="#f0f0f0").pack(side=tk.LEFT)
        close_minute = ttk.Combobox(time_frame, values=[f"{m:02d}" for m in range(60)], width=3, font=self.font_normal, state="readonly")
        close_minute.set("00")
        close_minute.pack(side=tk.LEFT)
        tk.Label(time_frame, textvariable=close_status_var, font=(self.font_normal[0], 16, 'bold'), fg="red",
                 bg="#f0f0f0").pack(side=tk.LEFT, padx=5)
        open_hour.bind("<<ComboboxSelected>>", lambda e, i=index: self.update_interrupt_status(i))
        open_minute.bind("<<ComboboxSelected>>", lambda e, i=index: self.update_interrupt_status(i))
        close_hour.bind("<<ComboboxSelected>>", lambda e, i=index: self.update_interrupt_status(i))
        close_minute.bind("<<ComboboxSelected>>", lambda e, i=index: self.update_interrupt_status(i))
        return {
            "file_path_var": file_path_var, "file_path": None,
            "open_hour": open_hour, "open_minute": open_minute, "open_status_var": open_status_var,
            "close_hour": close_hour, "close_minute": close_minute, "close_status_var": close_status_var,
        }
    def create_audio_controls(self, parent_frame):
        volume_frame = tk.Frame(parent_frame, bg="#f0f0f0")
        volume_frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(volume_frame, text="ระดับเสียง:", bg="#f0f0f0", font=self.font_normal).pack(side=tk.LEFT, pady=2)
        self.volume_slider = tk.Scale(volume_frame, from_=0, to=100, orient=tk.HORIZONTAL, length=300,
                                      command=self.adjust_volume, font=self.font_normal, bg="#f0f0f0",
                                      highlightthickness=0)
        self.volume_slider.set(100)
        self.volume_slider.pack(fill=tk.X, expand=True, padx=5)
        
        eq_bands = ["60", "170", "310", "600", "1K", "3K", "6K", "12K", "14K", "16K"]
        self.eq_sliders = []
        eq_frame = tk.Frame(parent_frame, bg="#f0f0f0")
        eq_frame.pack(fill=tk.X, expand=True, padx=5, pady=5)

        for i, band_name in enumerate(eq_bands):
            band_frame = tk.Frame(eq_frame, bg="#f0f0f0")
            band_frame.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
            tk.Label(band_frame, text=band_name, bg="#f0f0f0", font=("Arial", 10)).pack()
            eq_slider = tk.Scale(band_frame, from_=20, to=-20, orient=tk.VERTICAL, showvalue=0, length=80,
                                 command=lambda val, idx=i: self.adjust_equalizer(idx, val), bg="#f0f0f0",
                                 highlightthickness=0)
            eq_slider.set(0)
            eq_slider.pack()
            self.eq_sliders.append(eq_slider)
    def create_media_list(self, parent_frame):
        tree_frame = tk.Frame(parent_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        tree_scroll_y = tk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.media_tree = ttk.Treeview(tree_frame, columns=("Number", "Name", "Type", "Duration", "Status"),
                                       show="headings",
                                       yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.media_tree.heading("Number", text="ลำดับ")
        self.media_tree.heading("Name", text="ชื่อไฟล์")
        self.media_tree.heading("Type", text="ประเภท")
        self.media_tree.heading("Duration", text="ความยาว")
        self.media_tree.heading("Status", text="สถานะ")

        self.media_tree.column("Number", width=60, anchor='center')
        self.media_tree.column("Name", width=400)
        self.media_tree.column("Type", width=100, anchor='center')
        self.media_tree.column("Duration", width=100, anchor='center')
        self.media_tree.column("Status", width=120, anchor='center')
        self.media_tree.pack(fill=tk.BOTH, expand=True)
        tree_scroll_y.config(command=self.media_tree.yview)
        tree_scroll_x.config(command=self.media_tree.xview)
        self.media_tree.drop_target_register(DND_FILES)
        self.media_tree.dnd_bind('<<Drop>>', self.handle_drop)
        self.media_tree.bind("<Delete>", self.delete_selected_media)
        
        # สำหรับ Drag and Drop เพื่อเลื่อนลำดับ
        self.media_tree.bind("<ButtonPress-1>", self.on_tree_button_press)
        self.media_tree.bind("<B1-Motion>", self.on_tree_drag)
        self.media_tree.bind("<ButtonRelease-1>", self.on_tree_button_release)
        self.drag_data = {"item": None, "y": 0}
        
    def on_tree_button_press(self, event):
        item = self.media_tree.identify_row(event.y)
        if item:
            self.drag_data["item"] = item
            self.drag_data["y"] = event.y

    def on_tree_drag(self, event):
        if not self.drag_data["item"]: return
        item = self.media_tree.identify_row(event.y)
        if item and item != self.drag_data["item"]:
            target_index = self.media_tree.index(item)
            current_index = self.media_tree.index(self.drag_data["item"])
            
            # ย้ายใน Treeview
            self.media_tree.move(self.drag_data["item"], self.media_tree.parent(item), target_index)
            
            # ย้ายใน media_list ด้วย เพื่อให้เล่นถูกลำดับ
            path_to_move = self.media_list.pop(current_index)
            self.media_list.insert(target_index, path_to_move)
            
            # อัปเดตตัวเลขลำดับข้างหน้าใหม่ทั้งหมด
            for i, child in enumerate(self.media_tree.get_children()):
                values = list(self.media_tree.item(child, 'values'))
                values[0] = i + 1
                self.media_tree.item(child, values=tuple(values))

    def on_tree_button_release(self, event):
        self.drag_data["item"] = None
        self.drag_data["y"] = 0

    def create_main_playback_controls(self, parent_frame):
        tk.Button(parent_frame, text="▶ เล่น", command=self.play_main_media, bg="#4CAF50", fg="white",
                  font=self.font_bold, width=12).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(parent_frame, text="■ หยุด", command=self.stop_all_playback, bg="#F44336", fg="white",
                  font=self.font_bold, width=12).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(parent_frame, text="บันทึกการตั้งค่า", command=self.save_settings, bg="#FF9800", fg="white",
                  font=self.font_bold).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(parent_frame, text="รีเซ็ตการตั้งค่า", command=self.clear_all_settings, bg="#9E9E9E", fg="white",
                  font=self.font_bold).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(parent_frame, text="โหลดบันทึกตั้งค่า", command=self.load_settings, bg="#2196F3", fg="white",
                  font=self.font_bold).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Checkbutton(parent_frame, text="ปิดเครื่องหลังเล่นจบ", variable=self.shutdown_on_finish, font=self.font_bold,
                       bg="#f0f0f0").pack(side=tk.LEFT, padx=10, pady=10)
        tk.Checkbutton(parent_frame, text="เล่นวนซ้ำ (Loop)", variable=self.loop_media_var, font=self.font_bold,
                       bg="#f0f0f0").pack(side=tk.LEFT, padx=10, pady=10)
        
        sys_frame = tk.Frame(parent_frame, bg="#f0f0f0")
        sys_frame.pack(side=tk.RIGHT, padx=10)
        tk.Checkbutton(sys_frame, text="เปิดโปรแกรมตอนเปิดเครื่อง", variable=self.auto_start_var, font=self.font_normal,
                       bg="#f0f0f0", command=self.toggle_autostart).pack(side=tk.TOP, anchor='e')
        tk.Checkbutton(sys_frame, text="ย่อโปรแกรมไว้ที่มุมจอ (Tray)", variable=self.minimize_to_tray_var, font=self.font_normal,
                       bg="#f0f0f0").pack(side=tk.TOP, anchor='e')

    def create_dashboard_ui(self):
        dashboard_frame = tk.Frame(self.tab_dashboard, bg="white")
        dashboard_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title = tk.Label(dashboard_frame, text="สรุปสถิติการเล่น", font=("TH Sarabun New", 30, "bold"), bg="white", fg="#4CAF50")
        title.pack(pady=20)
        
        stats_frame = tk.Frame(dashboard_frame, bg="white")
        stats_frame.pack(pady=20)
        
        self.lbl_stat_main = tk.Label(stats_frame, text="🎵 สื่อหลักที่เล่นจบไปแล้ว: 0 รอบ", font=self.font_header, bg="white")
        self.lbl_stat_main.pack(anchor='w', pady=10)
        
        self.lbl_stat_int = tk.Label(stats_frame, text="📢 สื่อคั่นรายการที่เล่นไปแล้ว: 0 รอบ", font=self.font_header, bg="white")
        self.lbl_stat_int.pack(anchor='w', pady=10)
        
        tk.Label(dashboard_frame, text="(หมายเหตุ: สถิตินี้จะรีเซ็ตใหม่เมื่อปิดและเปิดโปรแกรมใหม่)", font=self.font_normal, bg="white", fg="gray").pack(side=tk.BOTTOM, pady=20)

    def update_dashboard_stats(self):
        self.lbl_stat_main.config(text=f"🎵 สื่อหลักที่เล่นจบไปแล้ว: {self.stats['main_played']} รอบ")
        self.lbl_stat_int.config(text=f"📢 สื่อคั่นรายการที่เล่นไปแล้ว: {self.stats['interrupt_played']} รอบ")

    def toggle_autostart(self):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE | winreg.KEY_READ)
            app_name = "RadioSystemApp"
            if self.auto_start_var.get():
                if getattr(sys, 'frozen', False):
                    exe_path = f'"{sys.executable}"'
                else:
                    exe_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Failed to set auto-start: {e}")

    def start_threads(self):
        threading.Thread(target=self.update_clock, daemon=True).start()
        threading.Thread(target=self.scheduler_loop, daemon=True).start()
        self.vlc_event_manager = self.main_player.event_manager()
        self.vlc_event_manager.event_attach(vlc.EventType.MediaPlayerEndReached, self.handle_main_media_end)
        self.vlc_interrupt_event_manager = self.interrupt_player.event_manager()
        self.vlc_interrupt_event_manager.event_attach(vlc.EventType.MediaPlayerEndReached,
                                                      self.handle_interrupt_media_end)
    def update_clock(self):
        while not self.stop_threads:
            now = datetime.datetime.now().strftime("%H:%M:%S")
            self.time_var.set(f"เวลาปัจจุบัน: {now}")
            time.sleep(1)
    def scheduler_loop(self):
        last_checked_minute = -1
        while not self.stop_threads:
            now = datetime.datetime.now()
            current_time = now.strftime("%H:%M")

            if now.minute != last_checked_minute:
                last_checked_minute = now.minute
                for entry in self.main_schedule_entries:
                    open_time = f"{entry['open_hour'].get()}:{entry['open_minute'].get()}"
                    close_time = f"{entry['close_hour'].get()}:{entry['close_minute'].get()}"
                    if open_time == current_time and open_time != "00:00":
                        self.root.after(0, self.play_main_media)
                    if close_time == current_time and close_time != "00:00":
                        self.root.after(0, self.stop_all_playback, True)
                for entry in self.interrupt_schedule_entries:
                    if entry.get("file_path"):
                        open_time = f"{entry['open_hour'].get()}:{entry['open_minute'].get()}"
                        close_time = f"{entry['close_hour'].get()}:{entry['close_minute'].get()}"
                        if open_time == current_time and open_time != "00:00":
                            self.root.after(0, self.start_interrupt, entry)
                        if close_time == current_time and close_time != "00:00":
                            self.root.after(0, self.stop_interrupt)
            time.sleep(1)
    def play_main_media(self, event=None):
        if self.is_playing_interrupt:
            self.status_var.set("กำลังเล่นไฟล์คั่นรายการ ไม่สามารถเล่นสื่อหลักได้")
            return
        if self.is_playing_main:
            self.handle_main_media_end()
            return
        next_media_item = self.get_next_media()
        if not next_media_item:
            self.status_var.set("ไม่มีสื่อในรายการให้เล่น")
            return
        try:
            self.current_media_item = next_media_item
            media = self.vlc_instance.media_new(self.current_media_item)
            self.main_player.set_media(media)
            if self.is_video_file(self.current_media_item):
                self.main_player.set_fullscreen(True)
            else:
                self.main_player.set_fullscreen(False)
            
            # Fade in
            self.main_player.audio_set_volume(0)
            self.main_player.play()
            self.fade_volume(self.main_player, 0, self.main_volume, duration=2.0)
            
            self.is_playing_main = True
            file_name = os.path.basename(self.current_media_item)
            self.status_var.set(f"กำลังเล่น: {file_name}")
            self.update_media_status_by_path(self.current_media_item, "▶ กำลังเล่น")
        except Exception as e:
            self.status_var.set(f"เกิดข้อผิดพลาด: {e}")
            self.is_playing_main = False
    def get_next_media(self):
        if not self.media_list: return None
        if self.play_random_mode.get():
            return random.choice(self.media_list)
        else:
            if not self.current_media_item or self.current_media_item not in self.media_list: return self.media_list[0]
            try:
                current_index = self.media_list.index(self.current_media_item)
                next_index = (current_index + 1) % len(self.media_list)
                return self.media_list[next_index]
            except ValueError:
                return self.media_list[0]
    def handle_main_media_end(self, event=None):
        # รันบน Thread หลักของ Tkinter เพื่อป้องกันอาการค้างหรือหลุดเวลาสั่งงานผ่าน Thread ของ VLC
        self.root.after(0, self._process_main_media_end)

    def _process_main_media_end(self):
        self.update_media_status_by_path(self.current_media_item, "✔ เล่นแล้ว")
        self.is_playing_main = False
        self.stats["main_played"] += 1
        self.update_dashboard_stats()
        # ตรวจสอบว่าเล่นจบทั้งหมดแล้วหรือยัง
        # เลิกใช้ is_last_file ที่ผูกกับลำดับไฟล์แบบเก่า เพื่อรองรับระบบเล่นสุ่ม (เล่นครบตามจำนวนรายการสื่อ)
        is_all_played = False
        
        # รีเซ็ตสถานะทั้งหมดเมื่อเล่นครบแล้ว
        if self.stats["main_played"] >= len(self.media_list) and len(self.media_list) > 0:
            is_all_played = True
            self.stats["main_played"] = 0 # reset stat for next loop
            self.update_dashboard_stats()

        if is_all_played:
            if self.loop_media_var.get() and not self.shutdown_on_finish.get():
                # เคลียร์สถานะในหน้าจอให้เริ่มใหม่ทั้งหมด
                for item in self.media_tree.get_children():
                    values = list(self.media_tree.item(item, 'values'))
                    values[4] = "ยังไม่เล่น"
                    self.media_tree.item(item, values=tuple(values))
                self.main_player.set_media(None)
            else:
                self.main_player.set_media(None) # ปิดหน้าต่างวิดีโอเมื่อเล่นครบทั้งหมด
                if self.shutdown_on_finish.get():
                    self.shutdown_computer("เล่นสื่อหลักครบทุกรายการแล้ว", force=False)
                self.status_var.set("รายการสื่อสิ้นสุดลงแล้ว")
                return
            
        self.root.after(100, self.play_main_media)
        
    def start_interrupt(self, interrupt_entry):
        if self.is_playing_interrupt: return
        
        file_path = interrupt_entry["file_path"]
        
        if self.main_player.is_playing():
            self.interrupted_media_item = self.current_media_item
            
            # Fade out before pausing main media
            def pause_after_fade():
                self.main_player.set_pause(1)
                self.is_playing_main = False
                self.update_media_status_by_path(self.interrupted_media_item, "⏸ หยุดชั่วคราว")
                self._start_interrupt_actual(file_path)

            self.fade_volume(self.main_player, self.main_volume, 0, duration=1.5, on_complete=pause_after_fade)
        else:
            self._start_interrupt_actual(file_path)
            
    def _start_interrupt_actual(self, file_path):
        try:
            media = self.vlc_instance.media_new(file_path)
            self.interrupt_player.set_media(media)
            if self.is_video_file(file_path):
                self.interrupt_player.set_fullscreen(True)
            else:
                self.interrupt_player.set_fullscreen(False)
            
            # เล่นไฟล์ก่อน ค่อยตั้งระดับเสียง เพื่อป้องกัน Error ถ้าระบบเสียงยังไม่พร้อม
            self.interrupt_player.play()
            self.root.after(300, lambda: self.interrupt_player.audio_set_volume(self.main_volume))
            
            self.is_playing_interrupt = True
            file_name_interrupt = os.path.basename(file_path)
            self.status_var.set(f"กำลังเล่นไฟล์คั่นรายการ: {file_name_interrupt}")
        except Exception as e:
            self.status_var.set(f"ไม่สามารถเล่นไฟล์คั่นรายการได้: {e}")
            self.resume_main_playback()
    def stop_interrupt(self):
        if not self.is_playing_interrupt: return
        self.interrupt_player.stop()
        self.handle_interrupt_media_end()
    def handle_interrupt_media_end(self, event=None):
        # รันบน Thread หลักของ Tkinter เพื่อป้องกันการเรียกคำสั่งของ Tkinter ซ้อนกับ Thread VLC
        self.root.after(0, self._process_interrupt_media_end)

    def _process_interrupt_media_end(self):
        if not self.is_playing_interrupt: return
        self.is_playing_interrupt = False
        self.stats["interrupt_played"] += 1
        self.update_dashboard_stats()
        self.status_var.set("ไฟล์คั่นรายการเล่นจบแล้ว")
        
        # ปิดบังคับล้าง Media และหน้าต่างของ VLC โดยตรง
        self.interrupt_player.set_media(None)
        self.close_vlc_video_windows()
        
        self.resume_main_playback()
        
    def resume_main_playback(self):
        if self.interrupted_media_item:
            self.status_var.set(f"กลับมาเล่น: {os.path.basename(self.interrupted_media_item)}")
            self.current_media_item = self.interrupted_media_item
            self.interrupted_media_item = None
            
            self.main_player.play()
            # Fade in resuming media
            self.fade_volume(self.main_player, 0, self.main_volume, duration=1.5)
            self.is_playing_main = True
        else:
            self.status_var.set("พร้อมใช้งาน")
            self.play_main_media()
    def stop_all_playback(self, from_schedule=False):
        self.main_player.stop()
        self.interrupt_player.stop()
        
        # ปิดหน้าต่างหน้าจอวีดีโอทุกชนิดเวลาสั่งหยุด
        self.main_player.set_media(None)
        self.interrupt_player.set_media(None)
        self.close_vlc_video_windows()
        
        self.is_playing_main = False
        self.is_playing_interrupt = False
        self.interrupted_media_item = None
        self.current_media_item = None
        
        for item in self.media_tree.get_children():
            values = list(self.media_tree.item(item, 'values'))
            values[4] = "ยังไม่เล่น"
            self.media_tree.item(item, values=tuple(values))
            
        if from_schedule:
            self.status_var.set("หยุดทำงานเนื่องจากล่วงเลยเวลาที่กำหนด")
        else:
            self.status_var.set("หยุดการเล่นทั้งหมด")
            
        if from_schedule and self.shutdown_on_finish.get(): self.shutdown_computer("ถึงเวลาปิดเครื่องตามที่ตั้งไว้", force=True)
    def select_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.current_directory.set(directory)
            self.media_list.clear()
            self.add_files_from_directory(directory)
            self.status_var.set(f"เลือกโฟลเดอร์: {os.path.basename(directory)}")
    def handle_drop(self, event):
        files_str = self.root.tk.splitlist(event.data)
        self.add_files_to_list(files_str)
        self.current_directory.set("นำเข้าไฟล์สื่อ")
    def add_files_from_directory(self, directory):
        supported_formats = ('.mp3', '.mp4', '.wav', '.avi', '.mkv', '.flac', '.ogg')
        files_to_add = [os.path.join(directory, f) for f in os.listdir(directory) if
                        f.lower().endswith(supported_formats)]
        self.add_files_to_list(files_to_add)
    def add_files_to_list(self, file_paths):
        newly_added = False
        for file_path in file_paths:
            if os.path.isfile(file_path) and file_path not in self.media_list:
                self.media_list.append(file_path)
                newly_added = True
        if newly_added: self.refresh_media_treeview()
    def refresh_media_treeview(self):
        for item in self.media_tree.get_children(): self.media_tree.delete(item)
        for i, file_path in enumerate(self.media_list):
            file_name = os.path.basename(file_path)
            file_ext = os.path.splitext(file_name)[1].lower()
            file_type = file_ext.replace('.', '').upper()
            duration = self.get_media_duration(file_path)
            self.media_tree.insert("", "end", iid=file_path,
                                   values=(i + 1, file_name, file_type, duration, "ยังไม่เล่น"))
        self.calculate_total_duration()
    def delete_selected_media(self, event=None):
        selected_items = self.media_tree.selection()
        if not selected_items: return
        if messagebox.askyesno("ยืนยันการลบ", f"คุณต้องการลบ {len(selected_items)} รายการที่เลือกใช่หรือไม่?"):
            for item_id in selected_items:
                if item_id in self.media_list: self.media_list.remove(item_id)
            self.refresh_media_treeview()
            self.status_var.set(f"ลบ {len(selected_items)} รายการสำเร็จ")
    def select_interrupt_file(self, index):
        file_path = filedialog.askopenfilename(title=f"เลือกไฟล์คั่นรายการรอบที่ {index + 1}",
                                               filetypes=[("Media Files", "*.mp3 *.mp4 *.wav *.avi *.mkv"),
                                                          ("All files", "*.*")])
        if file_path:
            entry = self.interrupt_schedule_entries[index]
            entry["file_path"] = file_path
            entry["file_path_var"].set(os.path.basename(file_path))
            self.update_interrupt_status(index)
    def update_interrupt_status(self, index):
        entry = self.interrupt_schedule_entries[index]
        if entry["open_hour"].get() != "00" or entry["open_minute"].get() != "00":
            entry["open_status_var"].set("✓")
        else:
            entry["open_status_var"].set("")
        if entry["close_hour"].get() != "00" or entry["close_minute"].get() != "00":
            entry["close_status_var"].set("✓")
        else:
            entry["close_status_var"].set("")
    def get_media_duration(self, file_path):
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.mp3':
                audio = MP3(file_path)
            elif ext == '.mp4':
                audio = MP4(file_path)
            elif ext == '.wav':
                audio = WAVE(file_path)
            else:
                return "--:--"
            seconds = int(audio.info.length)
            return f"{seconds // 60:02d}:{seconds % 60:02d}"
        except Exception:
            return "N/A"
    def calculate_total_duration(self):
        total_seconds = 0
        for file_path in self.media_list:
            duration_str = self.get_media_duration(file_path)
            if duration_str not in ["N/A", "--:--"]:
                try:
                    minutes, seconds = map(int, duration_str.split(':'))
                    total_seconds += (minutes * 60) + seconds
                except ValueError:
                    continue
        h = total_seconds // 3600;
        m = (total_seconds % 3600) // 60;
        s = total_seconds % 60
        self.total_duration_var.set(f"ระยะเวลารวม: {h:02d}:{m:02d}:{s:02d}")
        if total_seconds > 0:
            end_time = datetime.datetime.now() + datetime.timedelta(seconds=total_seconds)
            self.end_time_var.set(f"เวลาสิ้นสุดโดยประมาณ: {end_time.strftime('%H:%M:%S')}")
        else:
            self.end_time_var.set("เวลาสิ้นสุดโดยประมาณ: --:--:--")
    def update_media_status_by_path(self, file_path, status):
        if file_path and self.media_tree.exists(file_path):
            values = list(self.media_tree.item(file_path, 'values'))
            values[4] = status
            self.media_tree.item(file_path, values=tuple(values))
            
    def fade_volume(self, player, start_vol, end_vol, duration=2.0, steps=20, on_complete=None):
        """Fade audio volume linearly over a specified duration using tkinter root.after"""
        step_duration = int((duration * 1000) / steps)
        vol_change = (end_vol - start_vol) / steps

        def step_fade(current_step):
            if current_step <= steps:
                new_vol = int(start_vol + (vol_change * current_step))
                player.audio_set_volume(new_vol)
                self.root.after(step_duration, lambda: step_fade(current_step + 1))
            else:
                player.audio_set_volume(end_vol)
                if on_complete:
                    on_complete()
                    
        step_fade(1)

    def adjust_volume(self, value):
        self.main_volume = int(value)
        if hasattr(self, 'main_player'):
            self.main_player.audio_set_volume(self.main_volume)
        if hasattr(self, 'interrupt_player'):
            self.interrupt_player.audio_set_volume(self.main_volume)

    def adjust_equalizer(self, band, value):
        band_index = int(band);
        gain = float(value)
        self.eq_values[band_index] = gain
        if not hasattr(self, 'equalizer'): self.equalizer = vlc.libvlc_audio_equalizer_new()
        vlc.libvlc_audio_equalizer_set_amp_at_index(self.equalizer, gain, band_index)
        self.main_player.set_equalizer(self.equalizer)
        self.interrupt_player.set_equalizer(self.equalizer)
    def is_video_file(self, file_path):
        return file_path.lower().endswith(('.mp4', '.avi', '.mkv', '.mov'))

    def exit_fullscreen(self, event=None):
        self.main_player.set_fullscreen(False)
        self.interrupt_player.set_fullscreen(False)
        return "break"
    def save_settings(self):
        settings = {
            "main_schedule": [],
            "interrupt_schedule": [],
            "audio": {"volume": self.volume_slider.get(), "eq": [s.get() for s in self.eq_sliders]},
            "play_mode": "random" if self.play_random_mode.get() else "sequential",
            "media_list": self.media_list,
            "system": {
                "minimize_to_tray": self.minimize_to_tray_var.get(),
                "auto_start": self.auto_start_var.get(),
                "loop_media": self.loop_media_var.get()
            }
        }
        for entry in self.main_schedule_entries:
            settings["main_schedule"].append(
                {"open": f"{entry['open_hour'].get()}:{entry['open_minute'].get()}",
                 "close": f"{entry['close_hour'].get()}:{entry['close_minute'].get()}"}
            )
        for entry in self.interrupt_schedule_entries:
            settings["interrupt_schedule"].append(
                {"file": entry["file_path"],
                 "open": f"{entry['open_hour'].get()}:{entry['open_minute'].get()}",
                 "close": f"{entry['close_hour'].get()}:{entry['close_minute'].get()}"}
            )
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
            self.status_var.set("บันทึกการตั้งค่าสำเร็จแล้ว")
            messagebox.showinfo("บันทึกสำเร็จ", "บันทึกการตั้งค่าเรียบร้อยแล้ว")
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถบันทึกการตั้งค่าได้: {e}")
    def load_settings(self):
        if not self.settings_file.exists():
            self.status_var.set("ไม่พบบันทึกการตั้งค่า")
            return
        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                settings = json.load(f)
            for i, schedule_data in enumerate(settings.get("main_schedule", [])):
                if i < len(self.main_schedule_entries):
                    open_h, open_m = schedule_data.get("open", "00:00").split(":")
                    close_h, close_m = schedule_data.get("close", "00:00").split(":")
                    self.main_schedule_entries[i]["open_hour"].set(open_h)
                    self.main_schedule_entries[i]["open_minute"].set(open_m)
                    self.main_schedule_entries[i]["close_hour"].set(close_h)
                    self.main_schedule_entries[i]["close_minute"].set(close_m)
            for i, schedule_data in enumerate(settings.get("interrupt_schedule", [])):
                if i < len(self.interrupt_schedule_entries):
                    entry = self.interrupt_schedule_entries[i]
                    file_path = schedule_data.get("file")
                    if file_path and os.path.exists(file_path):
                        entry["file_path"] = file_path
                        entry["file_path_var"].set(os.path.basename(file_path))
                    else:
                        entry["file_path"] = None
                        entry["file_path_var"].set("ยังไม่ได้เลือกไฟล์")
                    open_h, open_m = schedule_data.get("open", "00:00").split(":")
                    close_h, close_m = schedule_data.get("close", "00:00").split(":")
                    entry["open_hour"].set(open_h)
                    entry["open_minute"].set(open_m)
                    entry["close_hour"].set(close_h)
                    entry["close_minute"].set(close_m)
                    self.update_interrupt_status(i)
            audio = settings.get("audio", {})
            self.volume_slider.set(audio.get("volume", 100))
            self.main_volume = int(audio.get("volume", 100))
            
            eq_values_loaded = audio.get("eq", [0.0] * 10)
            for i, val in enumerate(eq_values_loaded):
                if i < len(self.eq_sliders):
                    self.eq_sliders[i].set(val)
                    self.adjust_equalizer(i, val) # Apply EQ directly to vlc instance too
                    
            sys_settings = settings.get("system", {})
            self.minimize_to_tray_var.set(sys_settings.get("minimize_to_tray", True))
            self.auto_start_var.set(sys_settings.get("auto_start", False))
            self.loop_media_var.set(sys_settings.get("loop_media", False))

            if settings.get("play_mode") == "sequential":
                self.set_sequential_mode()
            else:
                self.set_random_mode()
            loaded_media = settings.get("media_list", [])
            # ตรวจสอบว่าไฟล์ยังมีอยู่จริงหรือไม่ ก่อนจะเพิ่มกลับเข้าไปในลิสต์
            self.media_list = [path for path in loaded_media if os.path.exists(path)]
            if self.media_list:
                self.refresh_media_treeview()
                self.current_directory.set("โหลดรายการสื่อจากไฟล์บันทึกแล้ว")
            self.status_var.set("โหลดการตั้งค่าสำเร็จ")
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถโหลดการตั้งค่าได้ กรุณาออกจากโปรแกรมเเละเปิดใหม่อีกครั้ง: {e}")

    def clear_all_settings(self):
        if messagebox.askyesno("ยืนยัน", "ต้องการลบการตั้งค่าทั้งหมดใช่หรือไม่?"):
            for entry in self.main_schedule_entries:
                entry["open_hour"].set("00")
                entry["open_minute"].set("00")
                entry["close_hour"].set("00")
                entry["close_minute"].set("00")

            for i, entry in enumerate(self.interrupt_schedule_entries):
                entry["file_path_var"].set("ยังไม่ได้เลือกไฟล์")
                entry["file_path"] = None
                entry["open_hour"].set("00")
                entry["open_minute"].set("00")
                entry["close_hour"].set("00")
                entry["close_minute"].set("00")
                self.update_interrupt_status(i)

            self.media_list.clear()
            for item in self.media_tree.get_children():
                self.media_tree.delete(item)
            self.refresh_media_treeview()

            self.volume_slider.set(100)
            for s in self.eq_sliders: s.set(0)
            for i in range(10): self.adjust_equalizer(i, 0)
                
            self.shutdown_on_finish.set(False)
            self.loop_media_var.set(False)
            self.minimize_to_tray_var.set(True)
            self.auto_start_var.set(False)
            self.toggle_autostart()
                
            self.set_random_mode()
            self.current_directory.set("ยังไม่ได้เลือกโฟลเดอร์ หรือนำเข้าไฟล์สื่อ")
            self.calculate_total_duration()
            self.stop_all_playback()
            self.status_var.set("รีเซ็ตการตั้งค่าทั้งหมดแล้ว")
            
            self.save_settings()

    def close_vlc_video_windows(self):
        try:
            # ใช้ไลบรารี pygetwindow เพื่อหาหน้าต่างของ VLC ทั้งหมดแล้วสั่งปิด
            if 'gw' in globals():
                windows = gw.getAllWindows()
                for window in windows:
                    if 'VLC' in window.title and 'Direct3D' in window.title:
                        window.close()
        except Exception as e:
            print(f"Error closing VLC windows: {e}")
    def set_random_mode(self):
        self.play_random_mode.set(True);
        self.update_play_mode_buttons()
        self.status_var.set("เปลี่ยนเป็นโหมด: เล่นแบบสุ่ม")
    def set_sequential_mode(self):
        self.play_random_mode.set(False);
        self.update_play_mode_buttons()
        self.status_var.set("เปลี่ยนเป็นโหมด: เล่นแบบเรียงลำดับ")
    def update_play_mode_buttons(self):
        if self.play_random_mode.get():
            self.random_btn.config(bg="#4CAF50", fg="white", relief=tk.SUNKEN)
            self.sequential_btn.config(bg="#E0E0E0", fg="black", relief=tk.RAISED)
        else:
            self.random_btn.config(bg="#E0E0E0", fg="black", relief=tk.RAISED)
            self.sequential_btn.config(bg="#4CAF50", fg="white", relief=tk.SUNKEN)
    def shutdown_computer(self, reason="", force=False):
        if force:
            self._execute_shutdown()
            return
            
        dialog = tk.Toplevel(self.root)
        dialog.title("ยืนยันการปิดเครื่อง")
        dialog.geometry("450x220")
        dialog.attributes('-topmost', True)
        
        # Center dialog
        dialog.update_idletasks()
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 225
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 110
            dialog.geometry(f"+{x}+{y}")
        except:
            pass

        countdown = [60]
        lbl_msg = tk.Label(dialog, text=f"{reason}\nระบบจะทำการปิดเครื่องคอมพิวเตอร์โดยอัตโนมัติ\nในอีก {countdown[0]} วินาที", font=self.font_normal)
        lbl_msg.pack(pady=20)
        
        def cancel_shutdown():
            self.shutdown_on_finish.set(False)
            self.status_var.set("ยกเลิกการปิดเครื่อง")
            dialog.destroy()
            
        def proceed_shutdown():
            dialog.destroy()
            self._execute_shutdown()
            
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="ปิดเครื่องเดี๋ยวนี้", command=proceed_shutdown, bg="#4CAF50", fg="white", font=self.font_normal, width=15).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="ยกเลิก", command=cancel_shutdown, bg="#F44336", fg="white", font=self.font_normal, width=15).pack(side=tk.RIGHT, padx=10)
        
        def update_countdown():
            if not dialog.winfo_exists():
                return
            countdown[0] -= 1
            if countdown[0] <= 0:
                proceed_shutdown()
            else:
                lbl_msg.config(text=f"{reason}\nระบบจะทำการปิดเครื่องคอมพิวเตอร์โดยอัตโนมัติ\nในอีก {countdown[0]} วินาที")
                dialog.after(1000, update_countdown)
                
        dialog.after(1000, update_countdown)

    def _execute_shutdown(self):
        self.status_var.set("กำลังจะปิดเครื่องใน 1 นาที...")
        print("Shutdown command initiated...")
        if os.name == 'nt':
            os.system("shutdown /s /t 60")
        else:
            os.system("sudo shutdown -h +1")

    def setup_keyboard_shortcuts(self):
        self.root.bind("<Escape>", self.exit_fullscreen)
        
    def create_tray_icon(self):
        # Create a simple icon image if we don't have one
        image = Image.new('RGB', (64, 64), color=(76, 175, 80))
        d = ImageDraw.Draw(image)
        d.text((10, 20), "Radio", fill=(255, 255, 255))
        
        try:
            icon_path = APPLICATION_PATH / 'icon.ico'
            if icon_path.exists():
                image = Image.open(icon_path)
        except Exception:
            pass

        menu = pystray.Menu(
            pystray.MenuItem('Show Program', self.restore_window, default=True),
            pystray.MenuItem('Exit', self.quit_from_tray)
        )
        self.tray_icon = pystray.Icon("name", image, "Auto Radio Player", menu)

    def hide_to_tray(self):
        self.root.withdraw()
        self.create_tray_icon()
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restore_window(self, icon, item):
        icon.stop()
        self.root.after(0, self.root.deiconify)

    def quit_from_tray(self, icon, item):
        icon.stop()
        self.stop_threads = True
        self.main_player.stop()
        self.interrupt_player.stop()
        self.root.after(0, self.root.destroy)

    def on_closing(self):
        # ถามผู้ใช้งานว่าต้องการบันทึกตั้งค่าก่อนปิดหรือไม่
        ans = messagebox.askyesnocancel("ออกจากโปรแกรม", "ต้องการบันทึกการตั้งค่าก่อนปิดโปรแกรมหรือไม่?")
        if ans is None:
            return # ผู้ใช้กด Cancel ยกเลิกการปิดโปรแกรม
        
        if ans:
            self.save_settings()

        if self.minimize_to_tray_var.get():
            self.hide_to_tray()
            return

        self.stop_threads = True
        self.main_player.stop()
        self.interrupt_player.stop()
        self.root.destroy()
if __name__ == "__main__":
    fonts_dir = APPLICATION_PATH / "fonts"
    if os.path.isdir(fonts_dir):
        for font_file in os.listdir(fonts_dir):
            if font_file.lower().endswith(".ttf"):
                load_font(fonts_dir / font_file)

    root = TkinterDnD.Tk()
    app = AudioSystemApp(root)
    root.mainloop()