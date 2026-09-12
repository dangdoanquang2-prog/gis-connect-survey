import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_step1_2_docx():
    doc = docx.Document()

    # Cấu hình lề trang A4 chuẩn
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # -------------------------------------------------------------
    # BẢNG MÀU TONE NÂU - BE (WARM COFFEE & BEIGE PALETTE)
    # -------------------------------------------------------------
    BROWN_PRIMARY = RGBColor(74, 53, 37)     # #4A3525 - Nâu cà phê đậm (Tiêu đề chính)
    BROWN_SECONDARY = RGBColor(122, 82, 48)  # #7A5230 - Nâu hổ phách / Caramel (Tiêu đề phụ)
    TEXT_DARK = RGBColor(44, 37, 35)         # #2C2523 - Chữ xám nâu than (Dịu mắt, không đen gắt)
    TEXT_MUTED = RGBColor(125, 115, 110)     # #7D736E - Chữ phụ chú
    
    HEX_HEADER_BG = "4A3525"                 # Nền tiêu đề bảng (Nâu đậm)
    HEX_ROW_LIGHT = "FAF7F2"                 # Nền hàng chẵn (Be sáng mềm)
    HEX_CALLOUT_BG = "F7F3ED"                # Nền hộp ghi chú (Be ấm)
    HEX_CODE_BG = "F4EFE6"                   # Nền khối mã (Giấy ngà / Be nhạt)
    HEX_BORDER = "D6CBBF"                    # Đường viền (Be xám nhạt)

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

    def set_table_borders(table, color=HEX_BORDER, sz="4", val="single"):
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
        run.font.size = Pt(13.5)
        run.font.bold = True
        run.font.color.rgb = BROWN_PRIMARY
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
        run.font.color.rgb = BROWN_SECONDARY
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
            r_bold.font.color.rgb = TEXT_DARK
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.italic = italic
        run.font.color.rgb = TEXT_DARK
        return p

    # Gạch đầu dòng thanh lịch (thay thế chấm tròn)
    def add_dash_item(text, bold_prefix="", level=0):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Inches(0.2 + 0.2 * level)
        
        # Dấu gạch đầu dòng
        r_dash = p.add_run("–  ")
        r_dash.font.name = 'Calibri'
        r_dash.font.size = Pt(10)
        r_dash.font.bold = True
        r_dash.font.color.rgb = BROWN_SECONDARY

        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.name = 'Calibri'
            r_bold.font.size = Pt(10)
            r_bold.font.bold = True
            r_bold.font.color.rgb = TEXT_DARK
            
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        run.font.color.rgb = TEXT_DARK
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
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
            f'<w:left w:val="single" w:sz="18" w:space="0" w:color="{HEX_HEADER_BG}"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="{HEX_BORDER}"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(code_text)
        r.font.name = 'Consolas'
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(60, 45, 35)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    def add_callout(text, title="ĐIỂM KỸ THUẬT LƯU Ý"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, HEX_CALLOUT_BG)
        set_cell_margins(cell, top=100, bottom=100, left=160, right=140)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="single" w:sz="20" w:space="0" w:color="{HEX_HEADER_BG}"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r_title = p.add_run(f"■ {title}: ")
        r_title.font.name = 'Calibri'
        r_title.font.size = Pt(10)
        r_title.font.bold = True
        r_title.font.color.rgb = BROWN_PRIMARY
        
        r_txt = p.add_run(text)
        r_txt.font.name = 'Calibri'
        r_txt.font.size = Pt(9.5)
        r_txt.font.color.rgb = TEXT_DARK
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # -------------------------------------------------------------
    # HEADER BLOCK (TONE NÂU BE)
    # -------------------------------------------------------------
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(6)
    p_top.paragraph_format.space_after = Pt(2)
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_top.add_run("DỰ ÁN KHẢO SÁT HIỆN TRƯỜNG & TỰ ĐỘNG HÓA DỮ LIỆU BÁO CÁO (GIS CONNECT)\n")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(10.5)
    r_sub.font.bold = True
    r_sub.font.color.rgb = BROWN_SECONDARY

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p_title.add_run("HƯỚNG DẪN KỸ THUẬT CHI TIẾT:\nCƠ CHẾ XỬ LÝ DỮ LIỆU BIỂU MẪU (BƯỚC 1)\nVÀ THUẬT TOÁN MÃ HÓA, NÉN ẢNH, WATERMARK TẠI MÁY KHÁCH (BƯỚC 2)")
    r_main.font.name = 'Calibri'
    r_main.font.size = Pt(14.5)
    r_main.font.bold = True
    r_main.font.color.rgb = BROWN_PRIMARY

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(2)
    p_meta.paragraph_format.space_after = Pt(14)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Tài liệu đặc tả giải pháp kỹ thuật luồng dữ liệu đầu vào — Cập nhật Tháng 09/2026")
    r_meta.font.name = 'Calibri'
    r_meta.font.size = Pt(9.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # PHẦN I: MỤC TIÊU CỐT LÕI
    # -------------------------------------------------------------
    add_h1("I. MỤC TIÊU KỸ THUẬT CỦA BƯỚC 1 VÀ BƯỚC 2")
    add_body("Hiệu quả của việc tự động hóa báo cáo và đồng bộ dữ liệu GIS phụ thuộc trực tiếp vào tính chuẩn xác của khâu thu thập hiện trường (Bước 1 và Bước 2). Khi dữ liệu đầu vào được kiểm soát chặt chẽ ngay tại nguồn, các công đoạn tổng hợp và xử lý phía sau sẽ diễn ra hoàn toàn tự động:")
    
    add_dash_item("Chuyển đổi hình thức nhập liệu từ các biểu mẫu tự do thành biểu mẫu có cấu trúc định sẵn. Khảo sát viên không cần nhập liệu thủ công nhiều, giúp hạn chế tối đa các sai sót phát sinh trong quá trình ghi nhận.", "Mục tiêu Bước 1 (Format dữ liệu): ")
    add_dash_item("Mọi công đoạn xử lý hình ảnh bao gồm đổi tên, nén dung lượng và đóng dấu thông tin thực địa (tọa độ GPS, thời gian) đều được thực hiện trực tiếp trên trình duyệt thiết bị di động (Client-side) ngay khi chụp, trước khi truyền tải về máy chủ.", "Mục tiêu Bước 2 (Xử lý ảnh tại máy khách): ")

    # -------------------------------------------------------------
    # PHẦN II: BƯỚC 1 - FORMAT DỮ LIỆU BIỂU MẪU
    # -------------------------------------------------------------
    add_h1("II. CHI TIẾT CƠ CHẾ XỬ LÝ BƯỚC 1: FORMAT BIỂU MẪU KHẢO SÁT")

    add_h2("1. Chuẩn Hóa Danh Mục Nhà Ga Bằng Cơ Chế Liên Kết Dữ Liệu Gốc")
    add_body("Để tránh việc người khảo sát nhập tên trạm tùy tiện (ví dụ: viết tắt, sai lỗi chính tả khiến hệ thống không thể tự động tổng hợp), biểu mẫu áp dụng cơ chế nạp danh mục cố định:")
    add_dash_item("Dữ liệu danh mục 26 nhà ga thuộc Tuyến Metro số 2 được tải trực tiếp từ tệp cơ sở dữ liệu 'stations_db.json' vào bộ nhớ của ứng dụng.", "Nạp danh mục trạm: ")
    add_dash_item("Giao diện cung cấp danh sách dạng Dropdown gồm định dạng chuẩn [Mã Ga] - [Tên Ga]. Ví dụ: 'S1 - Phan Văn Hớn', 'S2 - Đông Hưng Thuận', ..., 'S26 - Bến Thành'.", "Hiển thị lựa chọn: ")
    add_dash_item("Khi người khảo sát chọn trạm, hệ thống tự động gán ngầm các giá trị định danh gồm station_id = 'S1', station_lat = 10.836067, station_lng = 106.618255. Đây là trường khóa phục vụ ghép nối ảnh và liên kết không gian trên bản đồ GIS.", "Khóa liên kết ngầm: ")

    add_h2("2. Cơ Chế Tự Động Ghi Nhận Tọa Độ Vệ Tinh (GPS Auto-Capture)")
    add_body("Ứng dụng sử dụng hàm định vị tiêu chuẩn của trình duyệt (HTML5 Geolocation API) với cấu hình ưu tiên độ chính xác cao:")
    add_code_block(
        "// Ghi nhận tọa độ tự động khi mở biểu mẫu\n"
        "navigator.geolocation.getCurrentPosition(\n"
        "    (pos) => {\n"
        "        surveyRecord.device_lat = pos.coords.latitude;\n"
        "        surveyRecord.device_lng = pos.coords.longitude;\n"
        "        surveyRecord.device_accuracy = pos.coords.accuracy; // Sai số tính bằng mét (±m)\n"
        "        surveyRecord.gps_captured_time = new Date().toISOString();\n"
        "    },\n"
        "    (err) => console.warn('Lỗi định vị GPS:', err),\n"
        "    { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }\n"
        ");"
    )
    add_body("Khảo sát viên không cần thao tác kiểm tra tọa độ hay nhập số liệu thủ công. Hệ thống tự động ghi nhận vị trí thực tế kèm theo sai số đo đạc.")

    add_h2("3. Thiết Lập 6 Khung Chụp Ảnh Kỹ Thuật Cố Định")
    add_body("Trong biểu mẫu, 6 vị trí chụp ảnh được gán mã kỹ thuật cố định (Slot Code) tương ứng với các hạng mục khảo sát thực địa:")

    tbl_code_slots = doc.add_table(rows=7, cols=4)
    tbl_code_slots.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_code_slots)

    slot_h = ["Mã Vị Trí", "Tên Hạng Mục Khảo Sát", "Tính Chất Bắt Buộc", "Mục Đích Sử Dụng Trong Báo Cáo"]
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
        ("A1", "Hiện trạng mặt bằng vị trí ga", "Bắt buộc", "Chèn vào Chương 1 (Vị trí tim ga quy hoạch)"),
        ("A2", "Vỉa hè và lối tiếp cận đi bộ", "Bắt buộc", "Chèn vào Chương 2 (Đánh giá bề rộng và chất lượng vỉa hè)"),
        ("A3", "Trạm dừng xe buýt trung chuyển", "Bắt buộc", "Chèn vào Chương 3 (Hiện trạng kết nối vận tải công cộng)"),
        ("A4", "Bãi đỗ xe cá nhân (Park & Ride)", "Tùy chọn", "Chèn vào Chương 4 (Khu đất có tiềm năng làm bãi đỗ xe)"),
        ("A5", "Điểm đón trả khách nhanh (PUDO)", "Tùy chọn", "Chèn vào Chương 4 (Vị trí bố trí vịnh dừng đón/trả khách)"),
        ("A6", "Chướng ngại vật và điểm lấn chiếm", "Tùy chọn", "Chèn vào Chương 5 (Cột điện, cây xanh lớn, điểm nghẽn giao thông)")
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
                p.runs[0].font.color.rgb = BROWN_PRIMARY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------
    # PHẦN III: BƯỚC 2 - MÃ HÓA, NÉN & WATERMARK
    # -------------------------------------------------------------
    add_h1("III. CHI TIẾT CƠ CHẾ XỬ LÝ BƯỚC 2: MÃ HÓA, NÉN ẢNH VÀ WATERMARK")
    add_body("Khi người dùng chụp ảnh hoặc tải tệp lên tại từng khung khảo sát, ứng dụng sẽ thực thi đồng thời ba thuật toán xử lý nội bộ:")

    add_h2("1. Thuật Toán Tự Động Định Danh Tệp Ảnh (Auto-Naming Algorithm)")
    add_body("Hệ thống loại bỏ tên ảnh mặc định của thiết bị (dạng 'IMG_2026.jpg' hoặc mã số ngẫu nhiên) và khởi tạo tên tệp theo quy chuẩn thống nhất:")
    add_code_block(
        "function generateCleanFileName(stationId, slotCode, timestamp) {\n"
        "    // Chuyển đổi thời gian thành chuỗi YYYYMMDD_HHmmss\n"
        "    const dateStr = timestamp.toISOString().replace(/[-:T]/g, '').slice(0, 15);\n"
        "    // Quy chuẩn định danh: [Mã_Ga]_[Mã_Vị_Trí]_[Thời_Gian].jpg\n"
        "    return `${stationId}_${slotCode}_${dateStr}.jpg`;\n"
        "}\n"
        "// Kết quả định danh chuẩn: S1_A1_20260912_143025.jpg"
    )
    add_dash_item("Tên tệp ngắn gọn, không sử dụng dấu tiếng Việt, không chứa khoảng trắng hay ký tự đặc biệt.", "Đặc tính quy chuẩn: ")
    add_dash_item("Nhìn vào tên tệp có thể xác định chính xác ảnh thuộc nhà ga nào (S1), chụp hạng mục nào (A1) và thời điểm thực hiện.", "Khả năng nhận diện: ")

    add_h2("2. Thuật Toán Tối Ưu Dung Lượng Ảnh Tại Thiết Bị (Client-Side Compression)")
    add_body("Ảnh chụp trực tiếp từ điện thoại thông minh hiện nay có độ phân giải từ 12MP đến 48MP với dung lượng từ 5MB đến 12MB cho mỗi tệp. Việc gửi 6 ảnh nguyên bản qua mạng di động dễ gây hiện tượng chậm trễ hoặc lỗi kết nối:")
    add_dash_item("Ảnh được điều chỉnh kích thước về cạnh dài tối đa 1920px (chuẩn Full HD) thông qua Canvas.", "Điều chỉnh kích thước: ")
    add_dash_item("Tỷ lệ khung hình gốc (Aspect Ratio) được giữ nguyên để bảo đảm ảnh không bị méo mó.", "Bảo toàn tỷ lệ: ")
    add_dash_item("Ảnh xuất ra định dạng JPEG với hệ số nén chất lượng 0.82. Mức này bảo toàn độ rõ nét của biển báo, hiện trạng mặt đường trong khi giảm dung lượng từ 8MB xuống khoảng 600KB – 800KB.", "Mức nén tối ưu: ")

    add_h2("3. Thuật Toán Đóng Dấu Thông Tin Thực Địa Lên Ảnh (Watermarking)")
    add_body("Nhằm bảo đảm tính xác thực của dữ liệu khảo sát (xác nhận việc khảo sát được thực hiện tại hiện trường đúng thời điểm), ứng dụng tự động in dải thông tin định danh ở cạnh đáy bức ảnh:")

    add_code_block(
        "// Thuật toán ghi thông tin Watermark lên Canvas\n"
        "const ctx = canvas.getContext('2d');\n"
        "\n"
        "// 1. Tạo dải nền bán trong suốt ở đáy ảnh (chiều cao tương ứng 5% ảnh)\n"
        "const bannerHeight = Math.max(50, canvas.height * 0.05);\n"
        "ctx.fillStyle = 'rgba(0, 0, 0, 0.65)';\n"
        "ctx.fillRect(0, canvas.height - bannerHeight, canvas.width, bannerHeight);\n"
        "\n"
        "// 2. Thiết lập định dạng chữ hiển thị\n"
        "ctx.fillStyle = '#FFFFFF';\n"
        "ctx.font = `bold ${Math.round(bannerHeight * 0.32)}px Arial, sans-serif`;\n"
        "ctx.textBaseline = 'middle';\n"
        "\n"
        "// 3. Ghi thông tin trạm và tọa độ GPS thu nhận\n"
        "const line1 = `📍 ${stationName.toUpperCase()} | HẠNG MỤC: ${slotName}`;\n"
        "const line2 = `GPS: ${lat.toFixed(6)} N, ${lng.toFixed(6)} E (±${accuracy}m) | ${timeStr}`;\n"
        "\n"
        "ctx.fillText(line1, 20, canvas.height - bannerHeight * 0.65);\n"
        "ctx.fillText(line2, 20, canvas.height - bannerHeight * 0.25);"
    )

    add_callout("Sau khi xử lý qua Bước 2, tệp ảnh gốc dung lượng lớn được chuyển đổi thành tệp định danh chuẩn 'S1_A1_20260912_143025.jpg' dung lượng xấp xỉ 700KB, đồng thời in sẵn thông tin tọa độ và thời gian thực. Toàn bộ thao tác xử lý hoàn tất trong khoảng 0.3 giây ngay trên thiết bị mà không cần kết nối Internet.", "KẾT QUẢ XỬ LÝ BƯỚC 2")

    # -------------------------------------------------------------
    # PHẦN IV: CƠ CHẾ LƯU TẠM NGOẠI TUYẾN
    # -------------------------------------------------------------
    add_h1("IV. CƠ CHẾ LƯU TRỮ NGOẠI TUYẾN KHI MẤT SÓNG DI ĐỘNG")
    add_body("Tại các vị trí khảo sát không có sóng mạng 4G (khu vực ngầm hoặc tầng hầm), hệ thống kích hoạt cơ chế lưu trữ tạm thời (Offline-First):")
    add_dash_item("Dữ liệu phiếu khảo sát và chuỗi dữ liệu ảnh đã qua xử lý được lưu trữ an toàn trong bộ nhớ cục bộ (IndexedDB) của thiết bị.", "Lưu tạm trên thiết bị: ")
    add_dash_item("Ứng dụng thông báo trạng thái 'Đã lưu tạm trên thiết bị'. Khảo sát viên có thể tiếp tục công việc tại các điểm tiếp theo mà không bị gián đoạn.", "Duy trì khảo sát: ")
    add_dash_item("Khi thiết bị nhận lại tín hiệu mạng, hệ thống tự động đồng bộ ngầm các dữ liệu đã lưu lên máy chủ mà không đòi hỏi thao tác gửi lại.", "Tự động đồng bộ: ")

    # -------------------------------------------------------------
    # PHẦN V: ĐẶC TẢ CẤU TRÚC GÓI DỮ LIỆU
    # -------------------------------------------------------------
    add_h1("V. ĐẶC TẢ CẤU TRÚC GÓI DỮ LIỆU HOÀN CHỈNH (PAYLOAD SCHEMA)")
    add_body("Sau khi hoàn tất Bước 1 và Bước 2, dữ liệu được tổng hợp thành đối tượng chuẩn dạng JSON để truyền về máy chủ:")

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

    add_callout("Gói dữ liệu này cho phép máy chủ tách chuỗi ảnh ghi trực tiếp vào thư mục photos/S1/ với đúng tên định danh. Khi sinh báo cáo tự động, chương trình chỉ cần đối chiếu tên tệp để chèn ảnh vào văn bản mà không phải qua khâu lọc dữ liệu trung gian.", "KẾT NỐI VỚI HỆ THỐNG MÁY CHỦ")

    # Lưu tệp Word
    output_path = r"d:\gis_connect\HUONG_DAN_CHI_TIET_XU_LY_BUOC_1_VA_2.docx"
    doc.save(output_path)
    print("SUCCESS: File Word Buoc 1 va 2 (Tone Nau Be) da duoc tao tai: " + output_path)

if __name__ == '__main__':
    create_step1_2_docx()
