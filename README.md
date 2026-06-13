# ZaloPay Dispute Automation Agent

## Mô tả
Tự động tạo chứng từ Word và reply Freshdesk ticket cho ZaloPay OP dispute team.

## Tính năng
- Upload file Excel giao dịch → tạo hàng loạt file Word chứng từ
- Mock Freshdesk API (giả lập ticket, reply, CC)
- Giao diện web demo cho giám khảo
- Tải về ZIP file Word + log Excel

## Chạy local
```bash
pip install -r requirements.txt
python app.py
# Mở: http://localhost:8080
```

## Cấu trúc file
```
app.py                          # Web app chính + Mock Freshdesk API
requirements.txt                # Thư viện Python
Dockerfile                      # Deploy lên GreenNode
TEMPLATE_RUN_MEGER_-_Copy.xlsx  # Data giao dịch giả lập
freshdesk_reply_list_auto.xlsx  # Danh sách ticket giả lập
freshdesk_ticket_list.xlsx      # Ticket IDs giả lập
template_chung_tu_clean.docx    # Template Word (nếu có)
```

## Deploy GreenNode AgentBase
Xem hướng dẫn trong tài liệu cuộc thi.
