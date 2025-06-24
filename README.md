# Image2Excel
## Mô tả chức năng
Phần mềm Image2Excel sẽ:
1. Lần lượt lấy tên của từng sản phẩm (có dạng mã_sản_phẩm) trong file chứa danh sách tên sản phẩm (txt/xlsm).<br>
Tìm trong 1 thư mục chứa rất nhiều ảnh (png/jpg/img) các ảnh có tên file: `mã_sản_phẩm - 01` và `mã_sản_phẩm - 06`. Copy các ảnh này sang một thư mục mới (có tên: ImageMatched) cùng vị trí với thư mục ảnh.(không phải bên trong thư mục ảnh).<br>

2. Tạo 1 file exel xlsm mới với tên dạng: `output_date.xlsm`.(date : là ngày tháng hiện tại).<br>
Lần lượt dán từng: mã sản phẩm và cặp ảnh mã_sản_phẩm - 01; mã_sản_phẩm - 06 vào các cell tại cùng 1 hàng. Sau đó lưu file excel. Bắt đầu từ A2. Lặp lại cho tất cả mã tên sản phẩm có trong file danh sách tên mã sản phẩm.<br>

Ví dụ bố cụ exel đầu ra:<br>
|A|B|C|D|E|
|---|---|---|---|---|
||||||
|abc0123|abc012-01.jpg|abc012-06.jpg|
|abc456|abc456-01.jpg|abc456-06.jpg|

Image2Excel chỉ chạy trên Windows.

## Input
1. Thư mục (hỗ trợ local/SMB Shared) chứa rất nhiều ảnh
2. File danh sách đầu vào (txt/xlsm cũng hỗ trợ local/SMB Shared) chứa: danh sách mã các sản phẩm. Mỗi mã 1 dòng.

## Output
1. Thư mục chứa các ảnh đã tìm ra khớp với tên mã sản phẩm của tất cả mã trong file danh sách.
2. File excel đã chứa mã_sản_phẩm - ảnh 01 - ảnh 06 của tất cả mã sản phẩm trong file danh sách.

## Yêu cầu lập trình:
1. Ngôn ngữ: Python
2. Thư viện xử lý exel: OpenPyxl
3. GUI: dùng tkinter, định vị các đối tượng theo cú pháp `.pack` để dễ dàng co kéo layout khi cửa sổ thay đổi.
4. Tách biệt mã GUI và phần xử lý logic bằng: gui.py và main.py
5. Image2Excel sẽ được biên dịch sang `.exe` bằng `pyinstaller`. icon ico sẽ được nhúng trực tiếp vào `.exe` và sẽ giải nén vào temp file khi chạy để làm biểu tượng chính của chương trình.

### Mô tả GUI
**Layout sẽ như sau:**

|Object1|Object2|Object3|
|---|---|---|
|label: select product list|path view entry (hiện path tới file txt/xlsm)|button: Browser (mở dialog)|
|label: select image folder|path view entry (hiện path tới thư mục ảnh gốc)|button: Browser (mở dialog)|
|button: Run|button: Pause|button: Stop|
|log frame: hiển thị tiến độ, trạng thái, lỗi|

**Theme:**
* tối màu (modern dark)
* Button: màu xanh dương. Khi hover chuột/click chuột -> chuyển sang màu xanh tối hơn