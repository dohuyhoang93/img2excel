import os
import shutil
import threading
import time
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment
from itertools import chain

class Image2ExcelCore:
    def __init__(self, logger_callback=None):
        self._pause_event = threading.Event()
        self._stop_event = threading.Event()
        self._thread = None
        self.logger_callback = logger_callback or print

    def _log(self, msg):
        self.logger_callback(msg)

    def start(self, product_file_path, image_folder, matched_folder):
        self.product_file = Path(product_file_path)
        self.image_folder = Path(image_folder)
        self.matched_folder = Path(matched_folder)
        self._pause_event.set()
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run)
        self._thread.start()

    def _run(self):
        try:
            # Đọc danh sách mã sản phẩm
            if self.product_file.suffix.lower() == '.txt':
                with open(self.product_file, 'r', encoding='utf-8') as f:
                    codes = [line.strip() for line in f if line.strip()]
            elif self.product_file.suffix.lower() == '.xlsm':
                from openpyxl import load_workbook
                wb = load_workbook(self.product_file, data_only=True)
                sheet = wb.active
                codes = [str(cell.value).strip() for cell in sheet['A'] if cell.value]
            else:
                self._log("❌ Định dạng file không hợp lệ.")
                return

            # Kiểm tra và làm trống ImageMatched
            if self.matched_folder.exists():
                for file in chain(
                    self.matched_folder.glob("*.[jJ][pP][gG]"),
                    self.matched_folder.glob("*.[pP][nN][gG]"),
                    self.matched_folder.glob("*.[jJ][pP][eE][gG]")
                ):
                    file.unlink()
                    self._log(f"🗑️ Đã xóa: {file}")
            else:
                self.matched_folder.mkdir(exist_ok=True)

            wb = Workbook()
            ws = wb.active
            ws.append(["Mã SP", "Ảnh 01", "Ảnh 06"])
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 35
            ws.column_dimensions['C'].width = 35

            # Xây dựng chỉ mục
            suffixes = ["01", "06"]
            self.files_by_suffix = {suffix: [] for suffix in suffixes}
            for file in self.image_folder.rglob("*"):
                if file.suffix.lower() in ['.jpg', '.png', '.jpeg']:
                    if self.matched_folder in file.parents:
                        continue
                    name = file.stem.lower()
                    normalized_name = name.replace("-", "").replace("_", "").replace(".", "").replace(" ", "")
                    for suffix in suffixes:
                        if normalized_name.endswith(suffix):
                            self.files_by_suffix[suffix].append((file, normalized_name))

            for idx, code in enumerate(codes, start=2):
                if self._stop_event.is_set():
                    self._log("🛑 Đã dừng xử lý.")
                    break
                while not self._pause_event.is_set():
                    time.sleep(0.1)

                img1_path = self._find_image(code, "01")
                img6_path = self._find_image(code, "06")

                cell = ws.cell(row=idx, column=1)
                cell.value = code
                cell.alignment = Alignment(horizontal="center", vertical="center")

                if img1_path:
                    copied1 = shutil.copy(img1_path, self.matched_folder)
                    excel_img1 = ExcelImage(copied1)
                    excel_img1.width, excel_img1.height = 200, 200
                    ws.add_image(excel_img1, f"B{idx}")
                if img6_path:
                    copied6 = shutil.copy(img6_path, self.matched_folder)
                    excel_img6 = ExcelImage(copied6)
                    excel_img6.width, excel_img6.height = 200, 200
                    ws.add_image(excel_img6, f"C{idx}")

                ws.row_dimensions[idx].height = 150

                if img1_path and img6_path:
                    self._log(f"✅ {code}: OK")
                elif img1_path:
                    self._log(f"⚠️ {code}: Thiếu ảnh .06")
                elif img6_path:
                    self._log(f"⚠️ {code}: Thiếu ảnh .01")
                else:
                    self._log(f"⚠️ {code}: Thiếu cả hai ảnh .01 và .06")

            out_file = self.matched_folder / f"output_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            wb.save(out_file)
            self._log(f"📦 Đã lưu Excel: {out_file}")

        except Exception as e:
            self._log(f"❌ Lỗi: {e}")

    def _find_image(self, product_code, suffix):
        code = product_code.lower()
        normalized_code = code.replace("-", "").replace("_", "").replace(".", "").replace(" ", "")
        if suffix in self.files_by_suffix:
            for file, normalized_name in self.files_by_suffix[suffix]:
                if normalized_name.startswith(normalized_code) and normalized_name.endswith(suffix):
                    return file
        return None

    def pause(self):
        self._pause_event.clear()
        self._log("⏸️ Đã tạm dừng")

    def resume(self):
        self._pause_event.set()
        self._log("▶️ Đã tiếp tục")

    def stop(self):
        self._stop_event.set()
        self._pause_event.set()
        self._log("🛑 Đã dừng")