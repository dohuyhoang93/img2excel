import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import queue

class Image2ExcelGUI:
    def __init__(self, core_class):
        self.core = core_class()
        self.root = tk.Tk()
        self.root.title("Image2Excel")
        self.root.configure(bg="#222222")
        self.root.minsize(600, 400)
        self.status_text = tk.StringVar(value="Ready")
        self.progress_value = tk.DoubleVar(value=0)
        self.filter_var = tk.StringVar(value="Tất cả")
        self.all_iids = []  # Danh sách master chứa (IID, tag)
        self.log_queue = queue.Queue()
        self.progress_queue = queue.Queue()
        self.build_ui()
        self.root.after(50, self.poll_queues)

    def build_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', background='#222222', foreground='#ffffff')
        style.configure('TEntry', fieldbackground='#333333', foreground='#ffffff')
        style.configure('TButton', background='#007acc', foreground='#ffffff')
        style.configure('Treeview', background='#111111', fieldbackground='#111111', foreground='#ffffff')
        style.configure('Treeview.Heading', background='#333333', foreground='#ffffff')

        # MenuBar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Product List", command=self.browse_product_file)
        file_menu.add_command(label="Select Image Folder", command=self.browse_image_folder)
        file_menu.add_command(label="Select Output Folder", command=self.browse_matched_folder)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.grid(row=0, column=0, sticky='nsew')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Input frame
        input_frame = ttk.Frame(main_frame)
        input_frame.grid(row=0, column=0, sticky='ew')
        input_frame.columnconfigure(1, weight=1)

        # Row 1: Product list
        ttk.Label(input_frame, text="Product list:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.product_path = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.product_path).grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(input_frame, text="…", command=self.browse_product_file).grid(row=0, column=2, padx=5, pady=5)

        # Row 2: Image folder
        ttk.Label(input_frame, text="Image folder:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.image_path = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.image_path).grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(input_frame, text="…", command=self.browse_image_folder).grid(row=1, column=2, padx=5, pady=5)

        # Row 3: Output folder
        ttk.Label(input_frame, text="Output folder:").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.matched_path = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.matched_path).grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        ttk.Button(input_frame, text="…", command=self.browse_matched_folder).grid(row=2, column=2, padx=5, pady=5)

        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky='ew', pady=10)
        button_frame.columnconfigure(0, weight=1)
        ttk.Button(button_frame, text="Run", command=self.run_process).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Pause", command=self.core.pause).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="Resume", command=self.core.resume).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="Stop", command=self.core.stop).grid(row=0, column=3, padx=5)

        # Progressbar
        self.progressbar = ttk.Progressbar(main_frame, variable=self.progress_value, mode='determinate')
        self.progressbar.grid(row=4, column=0, sticky='ew', padx=5, pady=5)

        # Filter frame
        filter_frame = ttk.Frame(main_frame)
        filter_frame.grid(row=5, column=0, sticky='ew', padx=5, pady=5)
        filter_frame.columnconfigure(0, weight=1)
        ttk.Label(filter_frame, text="Filter:").grid(row=0, column=0, sticky='w', padx=5)
        ttk.Combobox(filter_frame, textvariable=self.filter_var, values=["Tất cả", "OK", "Thiếu ảnh", "Lỗi"], state='readonly').grid(row=0, column=1, sticky='w', padx=5)
        self.filter_var.trace('w', self.filter_log)

        # Treeview
        self.tree = ttk.Treeview(main_frame, columns=("Code", "Status"), show='headings', height=10)
        self.tree.heading("Code", text="Mã SP")
        self.tree.heading("Status", text="Trạng thái")
        self.tree.column("Code", width=100)
        self.tree.column("Status", width=300)
        self.tree.tag_configure('ok', foreground='#00ff00')
        self.tree.tag_configure('warning', foreground='#ffff00')
        self.tree.tag_configure('error', foreground='#ff0000')
        self.tree.grid(row=6, column=0, sticky='nsew', padx=5, pady=5)
        main_frame.rowconfigure(6, weight=1)

        # Status bar
        ttk.Label(main_frame, textvariable=self.status_text, background='#222222', foreground='#ffffff').grid(row=7, column=0, sticky='ew', padx=5, pady=5)

    def enqueue_log(self, item):
        self.log_queue.put(item)

    def enqueue_progress(self, item):
        self.progress_queue.put(item)

    def browse_product_file(self):
        path = filedialog.askopenfilename(filetypes=[("Product List", "*.txt *.xlsm")])
        if path:
            self.product_path.set(path)

    def browse_image_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.image_path.set(path)

    def browse_matched_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.matched_path.set(path)

    def run_process(self):
        product = self.product_path.get()
        images = self.image_path.get()
        matched = self.matched_path.get()
        if not product or not images or not matched:
            self.root.after(0, lambda: self.log(None, "⚠️ Chưa chọn đầy đủ đường dẫn.", 'error'))
            return
        self.status_text.set("Processing...")
        self.progress_value.set(0)
        self.tree.delete(*self.tree.get_children())
        self.all_iids.clear()
        self.core.start(product, images, matched,
                       log_queue=self.enqueue_log,
                       progress_queue=self.enqueue_progress)

    def log(self, code, message, tag):
        if code is None:  # Lỗi chung
            iid = self.tree.insert("", "end", values=("", message), tags=(tag,))
        else:
            iid = self.tree.insert("", "end", values=(code, message), tags=(tag,))
        print("LOGGED IID:", iid, "status:", message, "tag:", tag)  # Debug
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
        self.root.after(0, lambda: self.progress_value.set((current / total) * 100))

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

def run(core_class=None):
    if core_class is None:
        from main import Image2ExcelCore
        core_class = Image2ExcelCore
    app = Image2ExcelGUI(core_class)
    app.root.mainloop()