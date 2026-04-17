# 1. Sử dụng Python bản nhẹ để tiết kiệm chi phí chạy Cloud
FROM python:3.11-slim

# 2. Thiết lập thư mục làm việc
WORKDIR /app

# 3. Copy file requirements và cài đặt thư viện
COPY requirements.txt .

# Cài đặt thư viện (requirements.txt đã sạch)
RUN pip install --no-cache-dir -r requirements.txt

# 4. Copy mã nguồn xử lý lõi (Chỉ lấy phần logic, bỏ qua phần giao diện UI)
COPY transform.py .
COPY local_excel_runner.py .
COPY local_fact_builder.py .
COPY gsheet_sync.py .
COPY cloud_job_main.py .

# 5. Thiết lập biến môi trường mặc định (Có thể ghi đè khi chạy Job)
ENV GS_SPREADSHEET_ID=""
ENV FISCAL_YEAR="2026"
ENV GS_SERVICE_ACCOUNT_JSON=""

# 6. Lệnh khởi chạy
CMD ["python", "cloud_job_main.py"]
