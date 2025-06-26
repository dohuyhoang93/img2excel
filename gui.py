import tkinter as tk
from tkinter import filedialog, scrolledtext
from main import Image2ExcelCore
from pathlib import Path

class Image2ExcelGUI:
    def __init__(self, core_class):
        self.core = core_class(logger_callback=self.log)
        self.root = tk.Tk()
        self.root.title("Image2Excel")
        self.root.configure(bg="#222222")
        self.build_ui()

    def build_ui(self):
        label_fg = "#ffffff"
        entry_bg = "#333333"
        entry_fg = "#ffffff"
        button_bg = "#007acc"
        button_active = "#005f99"

        def dark_button(master, text, cmd):
            return tk.Button(master, text=text, bg=button_bg, fg="white",
                             activebackground=button_active, command=cmd)

        # Row 1: Chọn file mã sản phẩm
        tk.Label(self.root, text="Select product list:", fg=label_fg, bg="#222222").pack(anchor='w', padx=10, pady=5)
        self.product_path = tk.StringVar()
        entry1 = tk.Entry(self.root, textvariable=self.product_path, bg=entry_bg, fg=entry_fg, width=80)
        entry1.pack(padx=10, pady=2)
        dark_button(self.root, "Browse", self.browse_product_file).pack(padx=10, pady=2)

        # Row 2: Chọn thư mục ảnh
        tk.Label(self.root, text="Select image folder:", fg=label_fg, bg="#222222").pack(anchor='w', padx=10, pady=5)
        self.image_path = tk.StringVar()
        entry2 = tk.Entry(self.root, textvariable=self.image_path, bg=entry_bg, fg=entry_fg, width=80)
        entry2.pack(padx=10, pady=2)
        dark_button(self.root, "Browse", self.browse_image_folder).pack(padx=10, pady=2)

        # Row 3: Chọn thư mục ImageMatched
        tk.Label(self.root, text="Select output folder (ImageMatched):", fg=label_fg, bg="#222222").pack(anchor='w', padx=10, pady=5)
        self.matched_path = tk.StringVar()
        entry3 = tk.Entry(self.root, textvariable=self.matched_path, bg=entry_bg, fg=entry_fg, width=80)
        entry3.pack(padx=10, pady=2)
        dark_button(self.root, "Browse", self.browse_matched_folder).pack(padx=10, pady=2)

        # Row 4: Run - Pause - Resume - Stop
        frame_buttons = tk.Frame(self.root, bg="#222222")
        frame_buttons.pack(pady=10)
        dark_button(frame_buttons, "Run", self.run_process).pack(side='left', padx=5)
        dark_button(frame_buttons, "Pause", self.core.pause).pack(side='left', padx=5)
        dark_button(frame_buttons, "Resume", self.core.resume).pack(side='left', padx=5)
        dark_button(frame_buttons, "Stop", self.core.stop).pack(side='left', padx=5)

        # Log frame
        self.log_box = scrolledtext.ScrolledText(self.root, bg="#111111", fg="#00ffcc", height=20, width=200)
        self.log_box.pack(padx=10, pady=10)

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
            self.log("⚠️ Chưa chọn đầy đủ đường dẫn.")
            return
        self.core.start(product, images, matched)

    def log(self, message):
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)

def run(core_class=Image2ExcelCore):
    app = Image2ExcelGUI(core_class)
    app.root.mainloop()