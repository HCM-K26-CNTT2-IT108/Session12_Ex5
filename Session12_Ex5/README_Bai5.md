# Bài 5: Báo cáo sự cố an ninh mạng NetGuard

## Các file nộp
| File | Nội dung |
|---|---|
| `NetGuard_BaoCaoSuCo.docx` | Incident Report 7 trang, Auto TOC đã cập nhật số trang, Heading 1/2 dùng được với Navigation Pane |
| `NetGuard_NhatKyCanhBao.xlsx` | `NhatKyHeThong` (300 cảnh báo), `CSDL_Virus` (bảng tra), `BaoCao` (thống kê) |
| `NetGuard_PhongThu3Tang.pptx` | Slide sơ đồ phân cấp phòng thủ 3 tầng, font Montserrat |

## Phần 1. AutoFilter là "Google Search của bảng tính"
| Google Search | AutoFilter trong Excel |
|---|---|
| Hỏi một từ khóa trên cả kho dữ liệu khổng lồ | Gõ từ khóa vào ô tìm kiếm của danh sách lọc trên tiêu đề cột |
| Trả về đúng kết quả khớp, ẩn phần còn lại | Chỉ hiện các dòng khớp điều kiện, các dòng khác bị ẩn (không bị xóa) |
| Có bộ lọc nâng cao (thời gian, loại) | Lọc theo văn bản, số, ngày, màu; kết hợp nhiều cột cùng lúc (điều kiện AND) |
| Xóa bộ lọc là về trạng thái đầy đủ | Clear Filter đưa bảng về nguyên trạng |

**Cơ chế:** bật bằng Ctrl+Shift+L trên dòng tiêu đề, Excel nhận vùng dữ liệu liên tục xung quanh ô đang chọn rồi tạo danh sách giá trị duy nhất của mỗi cột. Điều kiện để lọc chính xác: dòng 1 là header duy nhất, không hàng/cột trống, không merge cell, và dữ liệu đã được làm sạch (dùng `TRIM`), nếu không các giá trị trông giống nhau sẽ bị tách thành nhiều mục.

**Ví dụ trong báo cáo:** lọc cột `MucDo` = "Cao" trên sheet `NhatKyHeThong` cho ngay 126 cảnh báo cần xử lý trong ngày.

## Phần 2. Triển khai báo cáo NetGuard
- **Excel:** `=IFERROR(VLOOKUP(TRIM($D2),CSDL_Virus!$A$2:$D$11,2,FALSE),"Chưa có trong CSDL")`. `TRIM` bọc ngoài `lookup_value` để mã có khoảng trắng thừa vẫn tra đúng. `$A$2:$D$11` khóa vùng bảng virus. Cột 2, 3, 4 lần lượt lấy loại mã độc, mức độ, cách xử lý. 3 mã lạ (`MD999`) hiện "Chưa có trong CSDL".
- **Kết quả:** 300 cảnh báo; mức Cao 126, Trung bình 120, Thấp 51; Trojan nhiều nhất (47 cảnh báo).
- **Word:** dùng Heading 1/2 để Navigation Pane hiển thị cây tài liệu. Mục lục tự động đã cập nhật. Khi tài liệu dài lên hàng chục trang, chuột phải vào mục lục > Update Field > Update entire table (hoặc F9) để số trang nhảy đúng.
- **PowerPoint:** sơ đồ phân cấp gồm tài sản cần bảo vệ và 3 tầng (mạng; máy chủ và ứng dụng; dữ liệu và con người). Font Montserrat. Nếu máy chiếu chưa cài Montserrat, cài font trước khi trình chiếu để hiển thị đúng.
- Toàn bộ dữ liệu cảnh báo, địa chỉ IP và sự cố là dữ liệu giả định.
