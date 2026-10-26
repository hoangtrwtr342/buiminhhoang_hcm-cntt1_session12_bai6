# Bài 6 : Hoàn thiện bộ hồ sơ triển khai nền tảng Smart-Home


## PHẦN 1. PHÂN TÍCH: BẢNG "XẤU" VS "ĐÚNG"
Gộp ô (Merge Cell): Bảng xấu gộp ô gây lỗi khi lọc/hàm; bảng đúng không gộp ô, mỗi ô là một dữ liệu độc lập.

Cấu trúc dữ liệu: Bảng xấu trộn thông tin hoặc ngắt dòng tùy tiện; bảng đúng chuẩn hóa mỗi cột là 1 Field, mỗi dòng là 1 Record.

Kiểu dữ liệu: Bảng xấu nhập chữ lẫn vào ô số; bảng đúng để số thuần túy và chỉnh hiển thị bằng Format Cells.

## PHẦN 2. TRIỂN KHAI DỰ ÁN SMART-HOME
### 1. Khung SRS trong Word (Heading chuẩn):

    Heading 1: 1. TỔNG QUAN HỆ THỐNG
    
    Heading 2: 1.1. Mục tiêu & Người dùng
    
    Heading 1: 2. YÊU CẦU CHỨC NĂNG
    
    Heading 2: 2.1. Đăng tin & Tra cứu BĐS
    
    Heading 2: 2.2. Đấu giá trực tuyến Real-time
    
    Heading 3: 2.2.1. Đặt cọc & Bước giá
    
    Heading 3: 2.2.2. Chốt hợp đồng tự động
    
    Heading 1: 3. YÊU CẦU PHI CHỨC NĂNG
    
    Heading 2: 3.1. Độ trễ < 500ms & Bảo mật 2 lớp

### 2. Công thức Excel:

  Giá thuê sau chiết khấu: =C2 * (1 - D2)
  
  Dò tỷ lệ hoa hồng: =VLOOKUP(TRIM(B2), $H$2:$I$5, 2, FALSE)
  
  Tiền hoa hồng: =E2 * F2

### 3. 10 Slide Pitch Deck (1 Slide - 1 Thông điệp):

  Tiêu đề: Smart-Home – Sàn đấu giá nhà đất trực tuyến minh bạch.
  
  Vấn đề: Bất động sản truyền thống thiếu minh bạch, môi giới cắt phế cao.
  
  Giải pháp: Đấu giá online theo thời gian thực (Real-time).
  
  Thị trường: Quy mô 15 tỷ USD, tăng trưởng 22%/năm.
  
  Sản phẩm: Đấu giá 1-chạm kết hợp Smart Contract tự động.
  
  Kinh doanh: Thu phí sàn 1.5% trên mỗi giao dịch thành công.
  
  Kế hoạch: Tiếp cận 5.000 căn hộ chung cư năm đầu.
  
  Cạnh tranh: Rút ngắn 50% thời gian chốt giao dịch so với môi giới cũ.
  
  Đội ngũ: Kết hợp chuyên môn CNTT và Bất động sản.

Kêu gọi vốn: Gọi 100.000 USD cho 10% cổ phần để phát triển MVP.

## PHẦN 3. KIỂM THỬ: SỬA LỖI TRÔI VÙNG DỮ LIỆU
Nguyên nhân: Dùng địa chỉ tương đối (H2:I5) nên khi kéo công thức xuống, vùng dò tìm bị trượt theo (H3:I6, H4:I7) làm thiếu dữ liệu, sinh lỗi #N/A.

Khắc phục: Dùng phím F4 để thêm dấu $ khóa cố định vùng dò tìm thành tuyệt đối: $H$2:$I$5.
