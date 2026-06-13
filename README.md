# ZOP HR Inbox — Dispute Automation Tool

Tự động tạo chứng từ Word/PDF và reply hàng loạt lên Freshdesk cho team OP xử lý dispute ZaloPay.

---

## Tính năng chính

### 1. Tạo chứng từ tự động từ Excel
- Upload file Excel giao dịch → tool tự đọc dữ liệu, điền vào template Word, nhóm theo ticket
- Mỗi ticket được tạo thành 1 file `.docx` riêng, đặt tên theo mã ticket

### 2. Xem chứng từ trực tiếp trên trình duyệt
- Click vào tên chứng từ → mở tab mới xem ngay, không cần tải về
- Nếu đã convert sang PDF thì hiển thị PDF (giữ định dạng gốc); nếu chưa thì render HTML từ docx

### 3. Chỉnh sửa và thay thế chứng từ
- Tải file `.docx` về, chỉnh sửa bằng Word, rồi upload lại để thay thế bản cũ
- Có thể thay thế 1 file hoặc nhiều file; các file không upload lại vẫn dùng bản tool đã tạo

### 4. Convert sang PDF trước khi gửi
- Bước 3 có nút **Convert sang PDF** — chuyển toàn bộ `.docx` sang `.pdf` để giữ đúng font/layout
- Chạy tự động bằng **LibreOffice** (Linux/Greennode) hoặc **Microsoft Word** (Windows)
- Nếu convert thất bại, tool tự fallback sang gửi file `.docx` và hiện cảnh báo

### 5. Chọn ticket muốn gửi
- Tick chọn từng ticket hoặc dùng **Chọn tất cả / Bỏ chọn tất cả**
- Chỉ những ticket được tick mới được gửi reply

### 6. Reply Freshdesk hàng loạt
- Gửi reply kèm đính kèm chứng từ (PDF hoặc docx) lên Freshdesk một lần cho tất cả ticket đã chọn
- Có CC tự động theo cấu hình từng ticket

### 7. Log kết quả
- Sau khi gửi, xem log ngay trên web: ticket nào thành công, thất bại, lỗi gì
- Tải log về file Excel (`.xlsx`) để lưu trữ

---

## Luồng sử dụng (4 bước)

```
Bước 1: Upload Excel       →  Bước 2: Xem & thay chứng từ  →  Bước 3: Convert PDF  →  Bước 4: Gửi & Log
Upload file giao dịch          Xem / tải / replace file           (tuỳ chọn)               Chọn ticket → Gửi
```

---

## Cấu trúc file

```
app.py                          # Flask app chính
requirements.txt                # Thư viện Python
Dockerfile                      # Deploy lên Greennode (có LibreOffice)

TEMPLATE_RUN_MEGER.xlsx         # ← File dữ liệu giao dịch (upload ở Bước 1)
freshdesk_reply_list_auto.xlsx  # Mapping ticket_id ↔ tên file chứng từ
freshdesk_reply_log.xlsx        # Log kết quả gửi (tự sinh)

template_chung_tu_clean.docx    # Template Word cho chứng từ

uploads/                        # File Excel tạm sau khi upload
output_grouped/                 # File chứng từ đã tạo (.docx / .pdf)
archive_data/                   # Lưu trữ dữ liệu cũ
```

---

## Chuẩn bị file trước khi dùng

### `TEMPLATE_RUN_MEGER.xlsx` — dữ liệu giao dịch
File Excel chứa các cột dữ liệu giao dịch cần điền vào chứng từ. Đây là file bạn upload ở **Bước 1**.

### `freshdesk_reply_list_auto.xlsx` — danh sách ticket
| Cột | Mô tả |
|-----|-------|
| `ticket_id` | ID ticket trên Freshdesk |
| `file_name` | Tên file chứng từ tương ứng (không cần đuôi `.docx`) |

### `template_chung_tu_clean.docx` — template Word
File Word chứa các placeholder `{{tên_cột}}` tương ứng với tên cột trong Excel. Tool sẽ tự điền dữ liệu vào.

---

## Chạy local (Windows)

```bash
# Cài thư viện
pip install -r requirements.txt

# Chạy app
python app.py

# Mở trình duyệt
http://localhost:8080
```

> **Lưu ý:** Trên Windows, bước Convert PDF dùng Microsoft Word (cần cài Word). Trên Linux dùng LibreOffice.

---

## Deploy lên Greennode

### Yêu cầu
- Greennode AgentBase (container runtime)
- Repo kết nối với Greennode qua GitHub

### Các bước

1. **Push code lên GitHub:**
   ```bash
   git add .
   git commit -m "update"
   git push origin main
   ```

2. **Greennode tự build Docker image** từ `Dockerfile` trong repo.

3. **Dockerfile đã bao gồm LibreOffice** — không cần cài thêm gì trên server:
   ```dockerfile
   RUN apt-get install -y --no-install-recommends libreoffice-writer ...
   ```

4. Sau khi deploy, truy cập qua URL Greennode cung cấp.

### Lưu ý khi deploy
- Các file Excel (`TEMPLATE_RUN_MEGER.xlsx`, `freshdesk_reply_list_auto.xlsx`) và template Word cần có trong repo hoặc được upload sau khi container chạy
- Thư mục `uploads/`, `output_grouped/`, `archive_data/` được tạo tự động khi khởi động
- Dữ liệu trong container **không persistent** — nếu container restart, file output bị xóa. Lưu log Excel về máy sau mỗi lần gửi.

---

## Tech stack

| Thành phần | Thư viện |
|-----------|---------|
| Web framework | Flask |
| Đọc/ghi Excel | pandas, openpyxl |
| Tạo file Word | python-docx |
| Xem docx trên web | mammoth |
| Convert sang PDF (Windows) | docx2pdf |
| Convert sang PDF (Linux) | LibreOffice headless |
| Gọi Freshdesk API | requests |
