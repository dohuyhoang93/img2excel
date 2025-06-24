# main.py
import code
import os
import shutil
import threading
from pathlib import Path
from datetime import datetime
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment
from PIL import Image as PILImage

class Image2ExcelCore:
    def __init__(self, logger_callback=None):
        self._pause_event = threading.Event()
        self._stop_event = threading.Event()
        self._thread = None
        self.logger_callback = logger_callback or print

    def _log(self, msg):
        self.logger_callback(msg)

    def start(self, product_file_path, image_folder, output_dir):
        self.product_file = Path(product_file_path)
        self.image_folder = Path(image_folder)
        self.output_dir = Path(output_dir)
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

            matched_folder = self.image_folder / "ImageMatched"
            matched_folder.mkdir(exist_ok=True)

            wb = Workbook()
            ws = wb.active
            ws.append(["Mã SP", "Ảnh 01", "Ảnh 06"])
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 35
            ws.column_dimensions['C'].width = 35

            for idx, code in enumerate(codes, start=2):
                if self._stop_event.is_set():
                    self._log("🛑 Đã dừng xử lý.")
                    break
                while not self._pause_event.is_set():
                    threading.Event().wait(0.1)

                # img1_name = f"{code}.01"
                # img6_name = f"{code}.06"
                # img1_path = self._find_image(img1_name)
                # img6_path = self._find_image(img6_name)
                img1_path = self._find_image(code, "01")
                img6_path = self._find_image(code, "06")

                # ws.cell(row=idx, column=1).value = code
                cell = ws.cell(row=idx, column=1)
                cell.value = code
                cell.alignment = Alignment(horizontal="center", vertical="center")

                if img1_path:
                    copied1 = shutil.copy(img1_path, matched_folder)
                    excel_img1 = ExcelImage(copied1)
                    excel_img1.width, excel_img1.height = 200, 200
                    ws.add_image(excel_img1, f"B{idx}")
                if img6_path:
                    copied6 = shutil.copy(img6_path, matched_folder)
                    excel_img6 = ExcelImage(copied6)
                    excel_img6.width, excel_img6.height = 200, 200
                    ws.add_image(excel_img6, f"C{idx}")

                # Tăng chiều cao hàng cho phù hợp ảnh
                ws.row_dimensions[idx].height = 150


                self._log(f"✅ {code}: {'OK' if img1_path and img6_path else 'Thiếu ảnh'}")

            out_file = matched_folder / f"output_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            wb.save(out_file)
            self._log(f"📦 Đã lưu Excel: {out_file}")

        except Exception as e:
            self._log(f"❌ Lỗi: {e}")

    def _find_image(self, product_code, suffix: str):
        code = product_code.lower()
        suffix = suffix.strip()

        for ext in ['.jpg', '.png', '.jpeg']:
            for file in self.image_folder.rglob(f"*{ext}"):
                name = file.stem.lower()
                if code in name and name.endswith(suffix):
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
