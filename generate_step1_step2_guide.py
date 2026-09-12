import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_step1_2_docx():
    doc = docx.Document()

    # Cấu hình Margins A4 chuẩn
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # Bảng màu kỹ thuật cao cấp
    NAVY = RGBColor(0, 51, 102)        # #003366 - Tiêu đề chính
    STEEL = RGBColor(41, 128, 185)     # #2980B9 - Tiêu đề phụ
    CHARCOAL = RGBColor(45, 55, 72)    # Văn bản thông thường
    MUTED = RGBColor(113, 128, 150)
    HEX_HEADER_BG = "003366"
    HEX_ROW_LIGHT = "F8FAFC"
    HEX_CODE_BG = "F1F5F9"

    def set_cell_background(cell, hex_color):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

    def set_table_borders(table, color="D0D7DE", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = NAVY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11.5)
        run.font.bold = True
        run.font.color.rgb = STEEL
        return p

    def add_body(text, bold_prefix="", italic=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.name = 'Calibri'
            r_bold.font.size = Pt(10)
            r_bold.font.bold = True
            r_bold.font.color.rgb = CHARCOAL
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.italic = italic
        run.font.color.rgb = CHARCOAL
        return p

    def add_bullet(text, bold_prefix="", level=0):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.name = 'Calibri'
            r_bold.font.size = Pt(10)
            r_bold.font.bold = True
            r_bold.font.color.rgb = CHARCOAL
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.color.rgb = CHARCOAL
        return p

    def add_code_block(code_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, HEX_CODE_BG)
        set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'<w:left w:val="single" w:sz="18" w:space="0" w:color="003366"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(code_text)
        r.font.name = 'Consolas'
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(30, 41, 59)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def add_callout(text, title="NGUYÊN TẮC KỸ THUẬT QUAN TRỌNG"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, "F0F4F8")
        set_cell_margins(cell, top=100, bottom=100, left=160, right=140)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="single" w:sz="20" w:space="0" w:color="003366"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r_title = p.add_run(f"📌 {title}: ")
        r_title.font.name = 'Calibri'
        r_title.font.size = Pt(10)
        r_title.font.bold = True
        r_title.font.color.rgb = NAVY
        
        r_txt = p.add_run(text)
        r_txt.font.name = 'Calibri'
        r_txt.font.size = Pt(9.5)
        r_txt.font.color.rgb = CHARCOAL
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # -------------------------------------------------------------
    # HEADER BLOCK
    # -------------------------------------------------------------
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(6)
    p_top.paragraph_format.space_after = Pt(2)
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_top.add_run("DỰ ÁN KHẢO SÁT HIỆN TRƯỜNG & TỰ ĐỘNG HÓA BÁO CÁO (GIS CONNECT)\n")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(11)
    r_sub.font.bold = True
    r_sub.font.color.rgb = STEEL

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p_title.add_run("HƯỚNG DẪN KỸ THUẬT CHI TIẾT:\nCƠ CHẾ XỬ LÝ BƯỚC 1 (FORMAT DỮ LIỆU FORM)\nVÀ BƯỚC 2 (MÃ HÓA, NÉN ẢNH, WATERMARK TẠI MÁY KHÁCH)")
    r_main.font.name = 'Calibri'
    r_main.font.size = Pt(15)
    r_main.font.bold = True
    r_main.font.color.rgb = NAVY

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(2)
    p_meta.paragraph_format.space_after = Pt(14)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Tài liệu đặc tả giải pháp kỹ thuật (Technical Specifications & Implementation Guide) — Tháng 09/2026")
    r_meta.font.name = 'Calibri'
    r_meta.font.size = Pt(9.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = MUTED

    # -------------------------------------------------------------
    # PHẦN I: MỤC TIÊU CỦA BƯỚC 1 & BƯỚC 2
    # -------------------------------------------------------------
    add_h1("I. MỤC TIÊU KỸ THUẬT CỦA BƯỚC 1 & BƯỚC 2")
    add_body("Toàn bộ sự thành bại của việc xuất báo cáo tự động và liên kết bản đồ GIS đều phụ thuộc 100% vào **khâu thu thập đầu vào (Bước 1 & Bước 2)**. Nếu đầu vào bị sai, ảnh đặt tên ngẫu nhiên hoặc file quá nặng làm nghẽn mạng thì toàn bộ các bước sau đều sụp đổ.")
    
    add_bullet("Biến biểu mẫu nhập liệu từ một form thông thường thành một form có cấu trúc chặt chẽ (Structured Schema). Khảo sát viên không cần gõ tay nhiều, giảm thiểu 95% sai sót con người.", "Mục tiêu Bước 1 (Format Biểu Mẫu): ")
    add_bullet("Toàn bộ công tác xử lý ảnh (Đổi tên, Nén dung lượng, Đóng dấu Watermark tọa độ/thời gian) phải được thực thi trực tiếp trên trình duyệt điện thoại của khảo sát viên (Client-side) ngay khoảnh khắc chụp ảnh, trước khi gửi bất kỳ byte dữ liệu nào về máy chủ.", "Mục tiêu Bước 2 (Xử Lý Ảnh Máy Khách): ")

    # -------------------------------------------------------------
    # PHẦN II: CHI TIẾT XỬ LÝ BƯỚC 1 - FORMAT DỮ LIỆU BIỂU MẪU
    # -------------------------------------------------------------
    add_h1("II. CHI TIẾT CƠ CHẾ XỬ LÝ BƯỚC 1: FORMAT BIỂU MẪU KHẢO SÁT")

    add_h2("1. Chuẩn Hóa Danh Mục Nhà Ga (Master Data Binding)")
    add_body("Thay vì để khảo sát viên tự gõ tên ga (rất dễ gõ sai chính tả như 'Phan van hon', 'S1_PVH', 'Ga s1' khiến máy tính không gom nhóm được), biểu mẫu sử dụng kỹ thuật liên kết dữ liệu danh mục:")
    add_bullet("Tải danh mục 26 nhà ga từ file cơ sở dữ liệu stations_db.json vào bộ nhớ ứng dụng.", "Nạp danh mục: ")
    add_bullet("Giao diện hiển thị Dropdown gồm [Mã Ga] - [Tên Ga] chuẩn mực. Ví dụ: S1 - Phan Văn Hớn, S2 - Đông Hưng Thuận, ..., S26 - Bến Thành.", "Giao diện chọn: ")
    add_bullet("Khi khảo sát viên chọn một ga, hệ thống lập tức gán ngầm các biến hệ thống: station_id = 'S1', station_lat = 10.836067, station_lng = 106.618255. Đây chính là khóa ngoại (Foreign Key) để tự động ghép ảnh và liên kết bản đồ GIS.", "Khóa dữ liệu ngầm: ")

    add_h2("2. Cơ Chế Tự Động Bắt Tọa Độ Vệ Tinh (GPS Auto-Capture)")
    add_body("Sử dụng chuẩn Geolocation API của HTML5 với độ chính xác cao nhất (High Accuracy):")
    add_code_block(
        "// Cơ chế tự động bắt tọa độ khi mở form\n"
        "navigator.geolocation.getCurrentPosition(\n"
        "    (pos) => {\n"
        "        surveyRecord.device_lat = pos.coords.latitude;\n"
        "        surveyRecord.device_lng = pos.coords.longitude;\n"
        "        surveyRecord.device_accuracy = pos.coords.accuracy; // Sai số tính bằng mét (±m)\n"
        "        surveyRecord.gps_captured_time = new Date().toISOString();\n"
        "    },\n"
        "    (err) => console.warn('Lỗi bắt GPS:', err),\n"
        "    { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }\n"
        ");"
    )
    add_body("Khảo sát viên hoàn toàn không phải mở app bản đồ hay gõ tọa độ bằng tay. Hệ thống tự động ghi nhận và kiểm soát sai số vệ tinh.")

    add_h2("3. Cấu Trúc 6 Khung Chụp Ảnh Kỹ Thuật (Fixed Photo Slots)")
    add_body("Trong biểu mẫu, 6 nút chụp ảnh được gán mã định danh bất biến (Slot Code) tương ứng với 6 góc chụp bắt buộc:")

    tbl_code_slots = doc.add_table(rows=7, cols=4)
    tbl_code_slots.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_code_slots)

    slot_h = ["Mã Slot", "Tên Hiển Thị Trên Form", "Ràng Buộc Kỹ Thuật", "Ý Nghĩa Đối Chiếu Báo Cáo"]
    for i, h in enumerate(slot_h):
        c = tbl_code_slots.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 100, 100)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(9.5)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    slot_rows = [
        ("A1", "Hiện trạng mặt bằng xây ga", "Bắt buộc (Required)", "Chèn vào Chương 1 của Báo cáo (Vị trí tim ga quy hoạch)"),
        ("A2", "Vỉa hè & Lối tiếp cận đi bộ", "Bắt buộc (Required)", "Chèn vào Chương 2 (Đánh giá bề rộng và chất lượng vỉa hè)"),
        ("A3", "Trạm dừng xe buýt kết nối", "Bắt buộc (Required)", "Chèn vào Chương 3 (Hiện trạng kết nối giao thông công cộng)"),
        ("A4", "Bãi đỗ xe cá nhân (P&R)", "Tùy chọn (Optional)", "Chèn vào Chương 4 (Khu đất tiềm năng làm bãi giữ xe máy/ô tô)"),
        ("A5", "Điểm đón trả khách (PUDO)", "Tùy chọn (Optional)", "Chèn vào Chương 4 (Vị trí vịnh dừng cho Taxi / Xe ôm công nghệ)"),
        ("A6", "Chướng ngại vật & Rào cản", "Tùy chọn (Optional)", "Chèn vào Chương 5 (Cột điện, cây xanh, miệng cống, lấn chiếm)")
    ]

    for row_idx, r in enumerate(slot_rows, start=1):
        row = tbl_code_slots.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(r):
            cell = row.cells[col_idx]
            cell.text = text
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 100, 100)
            p = cell.paragraphs[0]
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(9.5)
            if col_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------
    # PHẦN III: CHI TIẾT XỬ LÝ BƯỚC 2 - MÃ HÓA, NÉN & WATERMARK
    # -------------------------------------------------------------
    add_h1("III. CHI TIẾT CƠ CHẾ XỬ LÝ BƯỚC 2: MÃ HÓA, NÉN ẢNH & WATERMARK")
    add_body("Khi khảo sát viên chọn file hoặc chụp trực tiếp từ camera, sự kiện onchange kích hoạt hàm xử lý ảnh ngầm trên trình duyệt qua chuỗi 3 thuật toán liên tiếp:")

    add_h2("1. Thuật Toán Tự Động Đổi Tên Ảnh (Auto-Naming Algorithm)")
    add_body("Để triệt tiêu tình trạng file ảnh bị đặt tên vô nghĩa (như IMG_20260912_...jpg hoặc mã hash Zalo), hệ thống sinh tên file mới ngay tại bộ nhớ RAM:")
    add_code_block(
        "function generateCleanFileName(stationId, slotCode, timestamp) {\n"
        "    // Chuyển timestamp về chuỗi YYYYMMDD_HHmmss\n"
        "    const dateStr = timestamp.toISOString().replace(/[-:T]/g, '').slice(0, 15);\n"
        "    // Công thức bất biến: [MÃ_GA]_[MÃ_SLOT]_[TIMESTAMP].jpg\n"
        "    return `${stationId}_${slotCode}_${dateStr}.jpg`;\n"
        "}\n"
        "// Kết quả đầu ra chuẩn mực: S1_A1_20260912_143025.jpg"
    )
    add_bullet("Tên file ngắn gọn, không dấu tiếng Việt, không chứa khoảng trắng hay ký tự đặc biệt.", "Quy tắc an toàn: ")
    add_bullet("Khi nhìn vào tên file, cả con người và máy tính đều biết ngay bức ảnh này chụp tại Ga nào (S1), chụp hạng mục gì (A1) và chụp lúc nào.", "Tính tường minh: ")

    add_h2("2. Thuật Toán Nén Ảnh Tại Máy Khách (Client-Side Compression)")
    add_body("Điện thoại thông minh hiện nay chụp ảnh có độ phân giải rất cao (từ 12MP đến 48MP, dung lượng 5MB – 12MB/ảnh). Nếu nộp cả 6 bức ảnh gốc qua mạng 4G, tổng dung lượng lên đến 40MB – 60MB, gây ra lỗi timeout hoặc đứng app.")
    add_body("Giải pháp kỹ thuật sử dụng HTML5 Canvas để co kích thước và nén chất lượng:")
    add_bullet("Giới hạn chiều dài nhất (Max Dimension) của bức ảnh là 1920px (chuẩn Full HD).", "Co tỷ lệ: ")
    add_bullet("Tính toán tỷ lệ Aspect Ratio để ảnh không bao giờ bị méo hình (bảo toàn tỷ lệ gốc).", "Bảo toàn hình ảnh: ")
    add_bullet("Xuất ảnh dưới định dạng image/jpeg với hệ số nén Quality = 0.82. Đây là 'điểm vàng' giữa chất lượng và dung lượng: mắt thường không phân biệt được với ảnh gốc, nhưng dung lượng giảm từ 8MB xuống chỉ còn 600KB - 800KB!", "Chất lượng tối ưu: ")

    add_h2("3. Thuật Toán Đóng Dấu Watermark Tọa Độ & Thời Gian (Proof Overlay)")
    add_body("Để đảm bảo tính pháp lý và độ tin cậy của dữ liệu khảo sát (chứng minh khảo sát viên thực sự có mặt tại hiện trường đúng thời điểm), Canvas tự động vẽ một thanh băng đen mờ ở cạnh đáy bức ảnh:")

    add_code_block(
        "// Thuật toán vẽ Watermark lên Canvas\n"
        "const ctx = canvas.getContext('2d');\n"
        "\n"
        "// 1. Vẽ dải băng đen mờ (Chiều cao 60px)\n"
        "const bannerHeight = Math.max(50, canvas.height * 0.05);\n"
        "ctx.fillStyle = 'rgba(0, 0, 0, 0.65)';\n"
        "ctx.fillRect(0, canvas.height - bannerHeight, canvas.width, bannerHeight);\n"
        "\n"
        "// 2. Định dạng chữ Watermark sắc nét\n"
        "ctx.fillStyle = '#FFFFFF';\n"
        "ctx.font = `bold ${Math.round(bannerHeight * 0.32)}px Arial, sans-serif`;\n"
        "ctx.textBaseline = 'middle';\n"
        "\n"
        "// 3. In thông tin định danh và tọa độ vệ tinh\n"
        "const line1 = `📍 ${stationName.toUpperCase()} | HẠNG MỤC: ${slotName}`;\n"
        "const line2 = `GPS: ${lat.toFixed(6)} N, ${lng.toFixed(6)} E (±${accuracy}m) | ${timeStr}`;\n"
        "\n"
        "ctx.fillText(line1, 20, canvas.height - bannerHeight * 0.65);\n"
        "ctx.fillText(line2, 20, canvas.height - bannerHeight * 0.25);"
    )

    add_callout("Sau khi hoàn tất Bước 2, bức ảnh gốc nặng 8MB có tên 'IMG_2026.jpg' đã biến thành bức ảnh 'S1_A1_20260912_143025.jpg' nặng 700KB, trên góc ảnh in sẵn tọa độ GPS và thời gian. Toàn bộ quá trình này diễn ra trên máy khách trong đúng 0.3 giây mà không cần mạng Internet!", "HIỆU QUẢ CỦA BƯỚC 2")

    # -------------------------------------------------------------
    # PHẦN IV: CƠ CHẾ ĐÓNG GÓI & CHẠY OFFLINE
    # -------------------------------------------------------------
    add_h1("IV. CƠ CHẾ ĐÓNG GÓI DỮ LIỆU & CHẠY OFFLINE KHI MẤT MẠNG")
    add_body("Khi khảo sát viên khảo sát tại các vị trí không có sóng 4G (nhà ga ngầm, hầm chui), ứng dụng kích hoạt chế độ **Offline-First Resilience**:")
    add_bullet("Phiếu khảo sát và toàn bộ chuỗi Base64 của 6 bức ảnh đã nén được lưu trữ trực tiếp vào IndexedDB của trình duyệt điện thoại.", "Lưu tạm trên máy: ")
    add_bullet("Giao diện thông báo: 'Đã lưu cục bộ vào máy (Chờ kết nối mạng)'. Khảo sát viên tiếp tục di chuyển sang các ga khác khảo sát bình thường.", "Trải nghiệm không gián đoạn: ")
    add_bullet("Khi điện thoại kết nối lại mạng 4G/Wifi, Service Worker ngầm lập tức đẩy toàn bộ các phiếu tồn đọng lên máy chủ một cách an toàn mà không bị mất dữ liệu.", "Tự động đồng bộ (Auto-Sync): ")

    # -------------------------------------------------------------
    # PHẦN V: CẤU TRÚC GÓI TIN DỮ LIỆU (PAYLOAD SCHEMA)
    # -------------------------------------------------------------
    add_h1("V. CẤU TRÚC GÓI TIN DỮ LIỆU HOÀN CHỈNH (PAYLOAD SCHEMA)")
    add_body("Dữ liệu sau khi kết thúc Bước 1 và Bước 2 được đóng gói thành một đối tượng JSON chuẩn mực sẵn sàng gửi lên Server:")

    add_code_block(
        "{\n"
        '  "survey_id": "REC_S1_20260912_143025",\n'
        '  "station_id": "S1",\n'
        '  "station_name": "Phan Văn Hớn",\n'
        '  "surveyor_name": "Nguyễn Văn A",\n'
        '  "survey_time": "2026-09-12T14:30:25.000Z",\n'
        '  "gps_location": {\n'
        '    "latitude": 10.836068,\n'
        '    "longitude": 106.618255,\n'
        '    "accuracy_meters": 3.5\n'
        '  },\n'
        '  "assessments": {\n'
        '    "via_he_width_m": 3.2,\n'
        '    "via_he_quality": "Tốt",\n'
        '    "bus_stop_distance_m": 120,\n'
        '    "notes": "Mặt bằng thông thoáng, vỉa hè rộng rãi nhưng có 1 trụ điện hạ thế cần giải tỏa."\n'
        '  },\n'
        '  "photos": [\n'
        '    {\n'
        '      "slot_code": "A1",\n'
        '      "category_name": "Hiện trạng mặt bằng",\n'
        '      "file_name": "S1_A1_20260912_143025.jpg",\n'
        '      "file_size_kb": 745.2,\n'
        '      "data_base64": "data:image/jpeg;base64,...",\n'
        '      "captured_gps": { "lat": 10.836068, "lng": 106.618255 }\n'
        '    },\n'
        '    {\n'
        '      "slot_code": "A2",\n'
        '      "category_name": "Vỉa hè tiếp cận đi bộ",\n'
        '      "file_name": "S1_A2_20260912_143110.jpg",\n'
        '      "file_size_kb": 680.5,\n'
        '      "data_base64": "data:image/jpeg;base64,...",\n'
        '      "captured_gps": { "lat": 10.836072, "lng": 106.618260 }\n'
        '    }\n'
        '  ]\n'
        "}"
    )

    add_callout("Nhờ gói tin này: Khi lưu lên Server ở Bước 3, Server chỉ việc bóc tách chuỗi data_base64 để ghi thẳng thành file ảnh vật lý với đúng tên file_name vào thư mục photos/S1/. Khi làm báo cáo, script Python chỉ việc gọi đúng tên file để chèn vào Word mà không cần lọc dữ liệu!", "KẾT NỐI VỚI BƯỚC TIẾP THEO")

    # Lưu tài liệu Word
    output_path = r"d:\gis_connect\HUONG_DAN_CHI_TIET_XU_LY_BUOC_1_VA_2.docx"
    doc.save(output_path)
    print("SUCCESS: File Word chi tiet Buoc 1 va 2 da duoc tao tai: " + output_path)

if __name__ == '__main__':
    create_step1_2_docx()
