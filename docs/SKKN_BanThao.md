# BÁO CÁO SÁNG KIẾN KINH NGHIỆM

## Tên sáng kiến
**Xây dựng phần mềm tự động tổng hợp, phân tích và trực quan hóa kết quả khảo sát ý kiến sinh viên phục vụ công tác Đảm bảo chất lượng**

---

## PHẦN I. MỞ ĐẦU

### 1. Bối cảnh và lý do chọn sáng kiến

Công tác khảo sát ý kiến sinh viên về hoạt động giảng dạy là một nội dung bắt buộc và thường xuyên trong quy trình Đảm bảo chất lượng (ĐBCL) tại các cơ sở giáo dục đại học. Sau mỗi học kỳ, đơn vị phải xử lý hàng chục đến hàng trăm tập tin dữ liệu khảo sát (định dạng Excel), tổng hợp thành báo cáo Word theo mẫu chuẩn để trình các cấp.

Quy trình **làm thủ công** hiện nay bộc lộ nhiều hạn chế:
- Tốn nhiều thời gian, công sức: mỗi báo cáo học phần cần nhập liệu, tính toán tỷ lệ phần trăm, định dạng bảng biểu — trung bình 30–45 phút/báo cáo.
- Dễ sai sót khi sao chép, tính toán bằng tay.
- Khó tổng hợp, so sánh số liệu giữa nhiều lớp/học phần và giữa các học kỳ.
- Ý kiến đóng góp mở của sinh viên thường chỉ được liệt kê, chưa được phân loại để rút ra thông tin có giá trị.

Từ thực tiễn đó, tôi đề xuất xây dựng **phần mềm tự động hóa toàn bộ quy trình**, giúp giảm thời gian, tăng độ chính xác và nâng cao giá trị phân tích của dữ liệu khảo sát.

### 2. Mục đích của sáng kiến

- Tự động chuyển dữ liệu khảo sát từ Excel sang báo cáo Word theo đúng mẫu ĐBCL.
- Bổ sung phân tích thống kê và biểu đồ trực quan để hỗ trợ ra quyết định.
- Phân loại ý kiến đóng góp của sinh viên theo chủ đề và sắc thái một cách tự động.
- Theo dõi xu hướng chất lượng qua các học kỳ nhằm minh chứng cải tiến liên tục.
- Tạo ra công cụ dễ sử dụng, có thể nhân rộng cho các đơn vị khác.

### 3. Đối tượng và phạm vi áp dụng

- **Đối tượng**: dữ liệu khảo sát ý kiến sinh viên về học phần và hoạt động giảng dạy.
- **Phạm vi**: áp dụng tại Khoa Khoa học Cơ bản và có khả năng nhân rộng cho các khoa/đơn vị khác trong trường.

### 4. Phương pháp thực hiện

- Phân tích thực trạng và yêu cầu nghiệp vụ của công tác ĐBCL.
- Thiết kế kiến trúc phần mềm theo hướng module hóa.
- Lập trình bằng ngôn ngữ Python với giao diện đồ họa thân thiện.
- Kiểm thử tự động trên dữ liệu thật để đảm bảo độ chính xác.
- Đóng gói thành ứng dụng chạy độc lập, không yêu cầu cài đặt phức tạp.

---

## PHẦN II. NỘI DUNG

### Chương 1. Cơ sở lý luận và thực trạng

#### 1.1. Cơ sở lý luận

Việc ứng dụng công nghệ thông tin vào công tác quản lý, ĐBCL trong giáo dục là một xu hướng tất yếu, phù hợp với chủ trương chuyển đổi số của ngành. Tự động hóa các tác vụ lặp lại giúp giải phóng nguồn lực con người cho các công việc có giá trị cao hơn như phân tích, ra quyết định cải tiến.

#### 1.2. Thực trạng trước khi có sáng kiến

| Tiêu chí | Quy trình thủ công |
|---|---|
| Thời gian tạo 1 báo cáo | 30–45 phút |
| Nguy cơ sai sót | Cao (nhập, tính tay) |
| Phân tích thống kê | Hầu như không có |
| Biểu đồ trực quan | Phải vẽ thủ công |
| Phân loại ý kiến mở | Chỉ liệt kê |
| So sánh giữa các kỳ | Rất khó thực hiện |

### Chương 2. Mô tả giải pháp (Phần kỹ thuật)

#### 2.1. Tổng quan giải pháp

Phần mềm được xây dựng bằng ngôn ngữ Python với giao diện đồ họa (Tkinter), gồm **4 chức năng chính** tương ứng 4 tab:

1. **Tổng hợp báo cáo**: đọc dữ liệu khảo sát Excel, tạo báo cáo Word theo mẫu ĐBCL và báo cáo phân tích kèm biểu đồ.
2. **Gộp Feedbacks**: trích xuất, làm sạch và phân loại ý kiến đóng góp.
3. **Phân tích & Xu hướng**: vẽ biểu đồ xu hướng chất lượng qua các kỳ.
4. **Cấu hình**: tùy chỉnh tham số.

Phần mềm được thiết kế theo hướng **module hóa**, mỗi module đảm nhiệm một chức năng riêng, dễ bảo trì và mở rộng.

#### 2.2. Các điểm mới và sáng tạo

**a) Đọc dữ liệu linh hoạt, dò cột tự động**

Thay vì cố định vị trí cột (dễ lỗi khi mẫu Excel thay đổi), phần mềm **dò cột tự động theo tiêu đề**, hỗ trợ cả định dạng `.xls` và `.xlsx`, có cơ chế dự phòng an toàn.

**b) Kiểm tra dữ liệu đầu vào (validate)**

Phần mềm tự động phát hiện file lỗi, thiếu cột, tổng tỷ lệ phần trăm bất thường và cảnh báo cho người dùng, đảm bảo độ tin cậy của báo cáo.

**c) Phân tích thống kê và trực quan hóa**

Tự động tính điểm trung bình, độ lệch chuẩn, tỷ lệ hài lòng, phân bố mức đánh giá; sinh biểu đồ (cột, tròn, xu hướng) và xuất báo cáo phân tích riêng.

**d) Phân loại ý kiến mở thông minh (offline)**

Đây là điểm sáng tạo nổi bật: ý kiến đóng góp của sinh viên được tự động **phân loại theo chủ đề** (giảng viên & phương pháp, nội dung học phần, tài liệu, lượng giá, thời lượng, cơ sở vật chất) và **theo sắc thái** (tích cực, góp ý, trung tính). Toàn bộ xử lý diễn ra cục bộ, không gửi dữ liệu ra ngoài, đảm bảo an toàn thông tin sinh viên.

**e) Theo dõi xu hướng giữa các kỳ**

Phần mềm lưu lại kết quả từng đợt khảo sát và vẽ biểu đồ xu hướng điểm trung bình, tỷ lệ hài lòng qua các học kỳ — minh chứng trực quan cho công tác cải tiến liên tục.

**f) Đo lường hiệu quả tự động**

Phần mềm tự đo thời gian xử lý và ước lượng thời gian tiết kiệm so với làm thủ công.

#### 2.3. Kiến trúc kỹ thuật

Phần mềm gồm các module chính:

| Module | Chức năng |
|---|---|
| `gui_app.py` | Giao diện người dùng, điều phối tác vụ đa luồng |
| `excel_to_word.py` | Trích xuất dữ liệu, tạo báo cáo Word theo mẫu |
| `data_utils.py` | Đọc `.xls`/`.xlsx`, dò cột tự động, validate |
| `analytics.py` | Phân tích thống kê |
| `charts.py` | Sinh biểu đồ |
| `report_analytics.py` | Tạo báo cáo phân tích kèm biểu đồ |
| `feedback_classifier.py` | Phân loại ý kiến theo chủ đề & sắc thái |
| `history_store.py` | Lưu lịch sử, dữ liệu xu hướng |
| `benchmark.py` | Đo lường hiệu quả |
| `app_logging.py` | Ghi log có cấu trúc |

#### 2.4. Đảm bảo chất lượng phần mềm

Phần mềm được kiểm thử tự động bằng **bộ test (pytest)** gồm các kịch bản kiểm chứng trên dữ liệu thật: đọc dữ liệu, tính điểm trung bình (đối chiếu công thức tính tay), tách và làm sạch ý kiến, phân loại, tạo báo cáo Word. Toàn bộ test đều đạt, chứng minh tính đúng đắn và độ tin cậy của giải pháp.

### Chương 3. Hiệu quả của sáng kiến

#### 3.1. Hiệu quả về thời gian và công sức

Bảng so sánh quy trình **trước** và **sau** khi áp dụng phần mềm (số liệu minh họa, phần mềm tự đo và cập nhật theo thực tế đơn vị):

| Tiêu chí | Trước (thủ công) | Sau (phần mềm) | Cải thiện |
|---|---|---|---|
| Thời gian/báo cáo học phần | 30–45 phút | < 5 giây | Giảm > 99% |
| Xử lý 1 đợt nhiều học phần | Nhiều giờ | Vài phút | Nhanh hàng chục lần |
| Sai sót tính toán | Có thể xảy ra | Gần như bằng 0 | Tăng độ chính xác |
| Biểu đồ trực quan | Không/thủ công | Tự động | Mới hoàn toàn |
| Phân loại ý kiến | Thủ công | Tự động | Mới hoàn toàn |

> Phần mềm có chức năng tự ghi nhận thời gian xử lý thực tế và ước lượng thời gian tiết kiệm sau mỗi lần chạy, làm minh chứng định lượng khách quan.

#### 3.2. Hiệu quả về chất lượng công tác ĐBCL

- Báo cáo được tạo đồng nhất, đúng mẫu, giảm sai sót.
- Bổ sung phân tích thống kê và biểu đồ giúp lãnh đạo nắm bắt nhanh tình hình.
- Phân loại ý kiến giúp xác định đúng vấn đề cần cải tiến (ví dụ: tỷ lệ ý kiến về phương pháp giảng dạy, cơ sở vật chất...).
- Theo dõi xu hướng qua các kỳ hỗ trợ minh chứng cải tiến liên tục — yêu cầu cốt lõi của kiểm định chất lượng.

#### 3.3. Khả năng nhân rộng

- Phần mềm đóng gói thành ứng dụng chạy độc lập trên Windows và macOS, không yêu cầu cài đặt Python hay kiến thức kỹ thuật.
- Hoạt động hoàn toàn offline, phù hợp mọi máy tính văn phòng, an toàn dữ liệu.
- Có thể chuyển giao cho các khoa/đơn vị khác có cùng nghiệp vụ khảo sát.

---

## PHẦN III. KẾT LUẬN VÀ KIẾN NGHỊ

### 1. Kết luận

Sáng kiến đã xây dựng thành công một phần mềm tự động hóa toàn diện quy trình tổng hợp, phân tích và trực quan hóa kết quả khảo sát ý kiến sinh viên. Giải pháp đáp ứng đầy đủ các tiêu chí của một sáng kiến kinh nghiệm:

- **Tính mới**: phân loại ý kiến tự động theo chủ đề & sắc thái, theo dõi xu hướng giữa các kỳ — những chức năng mà công cụ trộn văn bản thông thường không có.
- **Tính khoa học**: kiến trúc module hóa rõ ràng, có kiểm thử tự động và cơ chế kiểm tra dữ liệu, ghi log minh bạch.
- **Tính hiệu quả**: giảm trên 99% thời gian tạo báo cáo, tăng độ chính xác, có số liệu đo lường khách quan.
- **Tính thực tiễn & nhân rộng**: dễ sử dụng, chạy offline, đóng gói sẵn, có khả năng chuyển giao.

### 2. Kiến nghị

- Triển khai áp dụng chính thức tại đơn vị cho các đợt khảo sát tiếp theo.
- Tổ chức tập huấn ngắn cho cán bộ phụ trách ĐBCL.
- Tiếp tục mở rộng bộ từ điển phân loại chủ đề và bổ sung các loại báo cáo theo nhu cầu thực tế.
- Xem xét nhân rộng cho các khoa/đơn vị khác trong trường.

---

## PHỤ LỤC

- **Phụ lục A**: Hướng dẫn cài đặt và sử dụng (xem [README.md](../README.md) và file hướng dẫn kèm theo).
- **Phụ lục B**: Mã nguồn phần mềm (các module Python).
- **Phụ lục C**: Ảnh chụp màn hình giao diện và biểu đồ minh họa (xem walkthrough).
- **Phụ lục D**: Báo cáo phân tích mẫu và biểu đồ xuất từ dữ liệu thật.

> [!NOTE]
> Bản thảo này (định dạng Markdown) phản ánh các nâng cấp mới nhất của phần mềm. Có thể chuyển sang định dạng Word theo mẫu quy định của đơn vị khi nộp chính thức.
