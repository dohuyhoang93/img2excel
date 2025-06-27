import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue
import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime
import ttkbootstrap as tb
import json

class Image2ExcelGUI:
    def __init__(self, core_class):
        self.core = core_class()
        self.root = tb.Window(themename="solar")
        self.root.title("Image2Excel")
        self.root.minsize(600, 500)
        # icon setup
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "icon.ico")
        else:
            icon_path = "icon.ico"
        self.root.iconbitmap(icon_path)

        # status and progress variables
        self.status_text = tk.StringVar(value="Ready")
        self.progress_value = tk.DoubleVar(value=0)
        self.filter_var = tk.StringVar(value="Tất cả")
        self.all_iids = []
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.export_folder = None
        self.start_time = None
        self.suffix_vars = [tk.StringVar() for _ in range(4)]
        self.settings_file = Path(os.getenv('APPDATA')) / "Image2Excel" / "settings.json"
        self.settings_file.parent.mkdir(exist_ok=True)  # Ensure directory exists

        self.load_settings()
        self.build_ui()
        self.root.after(50, self.poll_queues)

        # keyboard shortcuts
        self.root.bind('<Control-r>', lambda e: self.on_run())
        self.root.bind('<Control-p>', lambda e: self.on_pause())
        self.root.bind('<Control-o>', lambda e: self.on_resume())
        self.root.bind('<Control-s>', lambda e: self.on_stop())

    def load_settings(self):
        if self.settings_file.exists():
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                self.product_path.set(settings.get('product_path', ''))
                self.image_path.set(settings.get('image_path', ''))
                self.matched_path.set(settings.get('matched_path', ''))
                for i, suffix in enumerate(settings.get('suffixes', ['', '', '', ''])):
                    if i < len(self.suffix_vars):
                        self.suffix_vars[i].set(suffix)
            except Exception as e:
                print(f"Error loading settings: {e}")

    def save_settings(self):
        settings = {
            'product_path': self.product_path.get(),
            'image_path': self.image_path.get(),
            'matched_path': self.matched_path.get(),
            'suffixes': [var.get().strip() for var in self.suffix_vars]
        }
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def build_ui(self):
        main_frame = tb.Labelframe(self.root, text="Controls", padding=10)
        main_frame.grid(row=0, column=0, sticky='nsew')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1, minsize=200)
        self.root.rowconfigure(1, weight=2, minsize=300)
        main_frame.columnconfigure(0, weight=1)

        # input fields
        input_frame = tb.Frame(main_frame)
        input_frame.grid(row=0, column=0, sticky='ew')
        input_frame.columnconfigure(1, weight=1)
        for i, (label, var, cmd) in enumerate([
            ("Product list:", self.product_path, self.browse_product_file),
            ("Image folder:", self.image_path, self.browse_image_folder),
            ("Output folder:", self.matched_path, self.browse_matched_folder)
        ]):
            tb.Label(input_frame, text=label).grid(row=i, column=0, sticky='w', padx=5, pady=5)
            tb.Entry(input_frame, textvariable=var).grid(row=i, column=1, sticky='ew', padx=5, pady=5)
            tb.Button(input_frame, text="…", command=cmd).grid(row=i, column=2, padx=5, pady=5)

        # suffix inputs
        suffix_frame = tb.Frame(main_frame)
        suffix_frame.grid(row=1, column=0, sticky='ew', pady=(5, 5))
        tb.Label(suffix_frame, text="Image Suffixes:").grid(row=0, column=0, sticky='w', padx=5)
        for i, var in enumerate(self.suffix_vars):
            tb.Entry(suffix_frame, textvariable=var, width=10).grid(row=0, column=i+1, padx=5)

        # control buttons
        button_frame = tb.Frame(main_frame)
        button_frame.grid(row=2, column=0, sticky='ew', pady=(5, 5))
        for i, (text, cmd) in enumerate([
            ("Run", self.on_run),
            ("Pause", self.on_pause),
            ("Resume", self.on_resume),
            ("Stop", self.on_stop)
        ]):
            btn = tb.Button(button_frame, text=text, command=cmd)
            btn.grid(row=0, column=i, padx=5, sticky='ew')
            button_frame.columnconfigure(i, weight=1)

        # progress bar
        self.progressbar = tb.Progressbar(main_frame, variable=self.progress_value, mode='determinate')
        self.progressbar.grid(row=3, column=0, sticky='ew', padx=5, pady=5)

        # filter combo
        filter_frame = tb.Frame(main_frame)
        filter_frame.grid(row=4, column=0, sticky='ew', pady=(5, 10))
        filter_frame.columnconfigure(0, weight=1)
        tb.Label(filter_frame, text="Filter:").grid(row=0, column=0, sticky='w', padx=5)
        tb.Combobox(filter_frame, textvariable=self.filter_var, values=["Tất cả", "OK", "Thiếu ảnh", "Lỗi"], state='readonly').grid(row=0, column=1, sticky='w', padx=5)
        self.filter_var.trace('w', self.filter_log)

        # log table and status bar
        table_frame = tb.Labelframe(self.root, text="Log Table", padding=10)
        table_frame.grid(row=1, column=0, sticky='nsew')
        table_frame.columnconfigure(0, weight=1)

        canvas = tk.Canvas(table_frame)
        scrollbar = tb.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tb.Frame(canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        table_frame.rowconfigure(0, weight=1)

        self.tree = tb.Treeview(self.scrollable_frame, columns=("Code", "Status"), show='headings', height=30)
        self.tree.heading("Code", text="Mã SP")
        self.tree.heading("Status", text="Trạng thái")
        self.tree.column("Code", width=300)
        self.tree.column("Status", width=400)
        self.tree.tag_configure('ok', foreground='#00ff00')
        self.tree.tag_configure('warning', foreground='#ffff00')
        self.tree.tag_configure('error', foreground='#ff0000')
        self.tree.grid(row=0, column=0, sticky='nsew')
        self.scrollable_frame.columnconfigure(0, weight=2)

        status_frame = tb.Frame(self.root)
        status_frame.grid(row=2, column=0, sticky='ew')
        self.status_label = tb.Label(status_frame, textvariable=self.status_text)
        self.status_label.grid(row=0, column=0, sticky='w', padx=5)
        self.progress_label = tb.Label(status_frame, text="Progress: 0%")
        self.progress_label.grid(row=0, column=1, sticky='e', padx=5)

        # menu
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Product List", command=self.browse_product_file)
        file_menu.add_command(label="Select Image Folder", command=self.browse_image_folder)
        file_menu.add_command(label="Select Output Folder", command=self.browse_matched_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Open Export Folder", command=self.open_export_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Light Mode", command=lambda: self.toggle_theme("cosmo"))
        view_menu.add_command(label="Dark Mode", command=lambda: self.toggle_theme("solar"))
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)

    def on_run(self):
        self.save_settings()
        self.run_process()
        self.status_text.set("Running")

    def on_pause(self):
        self.core.pause()
        self.status_text.set("Paused")

    def on_resume(self):
        self.core.resume()
        self.status_text.set("Running")

    def on_stop(self):
        self.core.stop()
        self.status_text.set("Stopped")

    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("About Image2Excel")
        if getattr(sys, 'frozen', False):
            icon_path = os.path.join(sys._MEIPASS, "icon.ico")
        else:
            icon_path = "icon.ico"
        about_window.iconbitmap(icon_path)
        about_window.geometry("300x200")
        about_label = tk.Label(about_window, text="Image2Excel\nVersion: 1.0\nDeveloped by: Do Huy Hoang Fujikin Vietnam\nDate: June 26, 2025", justify="center")
        about_label.pack(expand=True)
        close_button = tk.Button(about_window, text="Close", command=about_window.destroy)
        close_button.pack(pady=10)

    

    def set_tooltip(self, widget, text):
        tooltip = tk.Toplevel(widget)
        tooltip.withdraw()
        tooltip.attributes('-topmost', True)
        tk.Label(tooltip, text=text, background='#ffffe0', relief='solid', borderwidth=1).pack()
        widget.bind("<Enter>", lambda e: tooltip.deiconify())
        widget.bind("<Leave>", lambda e: tooltip.withdraw())

    def browse_product_file(self):
        path = filedialog.askopenfilename(filetypes=[("Product List", "*.txt *.xlsx")])
        if path:
            self.product_path.set(path)
            self.save_settings()

    def browse_image_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.image_path.set(path)
            self.save_settings()

    def browse_matched_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.matched_path.set(path)
            self.save_settings()

    def run_process(self):
        product = self.product_path.get()
        images = self.image_path.get()
        matched = self.matched_path.get()
        suffixes = [var.get().strip() for var in self.suffix_vars if var.get().strip()]
        if not product or not images or not matched:
            self.root.after(0, lambda: self.log(None, "⚠️ Chưa chọn đầy đủ đường dẫn.", 'error'))
            return
        if not suffixes:
            self.root.after(0, lambda: self.log(None, "⚠️ Chưa nhập hậu tố nào.", 'error'))
            return
        self.status_text.set(f"Running since {datetime.now().strftime('%H:%M:%S')}")
        self.start_time = datetime.now()
        self.progress_value.set(0)
        self.tree.delete(*self.tree.get_children())
        self.all_iids.clear()
        self.export_folder = Path(matched) / "Image2Excel_Export"
        self.core.start(product, images, matched, suffixes,
                       log_queue=self.enqueue_log,
                       progress_queue=self.enqueue_progress)

    def open_export_folder(self):
        folder = getattr(self, "export_folder", None)
        if not folder or not folder.exists():
            self.log(None, "⚠️ Chưa có thư mục xuất hoặc nó chưa tồn tại.", 'error')
            return
        if sys.platform.startswith('win'):
            os.startfile(folder)
        elif sys.platform == 'darwin':
            subprocess.run(['open', folder])
        else:
            subprocess.run(['xdg-open', folder])

    def enqueue_log(self, item):
        self.log_queue.put(item)

    def enqueue_progress(self, item):
        self.progress_queue.put(item)

    def log(self, code, message, tag):
        if code is None:
            iid = self.tree.insert("", "end", values=("", message), tags=(tag,))
        else:
            iid = self.tree.insert("", "end", values=(code, message), tags=(tag,))
        print("LOGGED IID:", iid, "status:", message, "tag:", tag)
        self.all_iids.append((iid, tag))
        self.tree.yview_moveto(1)
        f = self.filter_var.get()
        if f != "Tất cả":
            ok = (f == "OK" and tag == 'ok') or \
                 (f == "Thiếu ảnh" and tag == 'warning') or \
                 (f == "Lỗi" and tag == 'error')
            if not ok:
                self.tree.detach(iid)

    def update_progress(self, total, current):
        percent = (current / total) * 100
        self.root.after(0, lambda: self.progress_value.set(percent))
        self.root.after(0, lambda: self.status_text.set(f"Started: {self.start_time.strftime('%H:%M:%S')}"))
        self.root.after(0, lambda: self.progress_label.configure(text=f"Progress: {percent:.1f}%"))

    def filter_log(self, *args):
        f = self.filter_var.get()
        for iid, tag in self.all_iids:
            self.tree.detach(iid)
        for iid, tag in self.all_iids:
            cond = (f == "Tất cả") or \
                   (f == "OK" and tag == 'ok') or \
                   (f == "Thiếu ảnh" and tag == 'warning') or \
                   (f == "Lỗi" and tag == 'error')
            if cond:
                self.tree.reattach(iid, "", "end")

    def poll_queues(self):
        while not self.log_queue.empty():
            code, message, tag = self.log_queue.get_nowait()
            self.root.after(0, lambda c=code, m=message, t=tag: self.log(c, m, t))
        while not self.progress_queue.empty():
            total, current = self.progress_queue.get_nowait()
            self.root.after(0, lambda t=total, c=current: self.update_progress(t, c))
        self.root.after(50, self.poll_queues)

    def toggle_theme(self, theme):
        self.root.style.theme_use(theme)
        self.status_label.configure(foreground=self.root.style.colors.fg)
        self.progress_label.configure(foreground=self.root.style.colors.fg)
        self.status_text.set(f"Switched to {theme.capitalize()} Mode")

    @property
    def product_path(self):
        if not hasattr(self, '_product_path'):
            self._product_path = tk.StringVar()
        return self._product_path

    @property
    def image_path(self):
        if not hasattr(self, '_image_path'):
            self._image_path = tk.StringVar()
        return self._image_path

    @property
    def matched_path(self):
        if not hasattr(self, '_matched_path'):
            self._matched_path = tk.StringVar()
        return self._matched_path

def run(core_class=None):
    if core_class is None:
        from main import Image2ExcelCore
        core_class = Image2ExcelCore
    app = Image2ExcelGUI(core_class)
    app.root.mainloop()