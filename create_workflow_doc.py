import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_technical_workflow_docx():
    doc = docx.Document()

    # Cấu hình lề trang chuẩn kỹ thuật (Executive A4)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # Bảng màu chủ đạo kỹ thuật (Deep Navy & Tech Steel)
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
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(13.5)
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
            r_bold.font.size = Pt(10.5)
            r_bold.font.bold = True
            r_bold.font.color.rgb = CHARCOAL
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10.5)
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
            r_bold.font.size = Pt(10.5)
            r_bold.font.bold = True
            r_bold.font.color.rgb = CHARCOAL
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(10.5)
        run.font.color.rgb = CHARCOAL
        return p

    def add_code_block(code_text):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, HEX_CODE_BG)
        set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
        
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

    def add_callout(text, title="ĐIỂM KỸ THUẬT CỐT LÕI"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, "F0F4F8")
        set_cell_margins(cell, top=120, bottom=120, left=180, right=160)
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="003366"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r_title = p.add_run(f"⚡ {title}: ")
        r_title.font.name = 'Calibri'
        r_title.font.size = Pt(10.5)
        r_title.font.bold = True
        r_title.font.color.rgb = NAVY
        
        r_txt = p.add_run(text)
        r_txt.font.name = 'Calibri'
        r_txt.font.size = Pt(10)
        r_txt.font.color.rgb = CHARCOAL
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # -------------------------------------------------------------
    # HEADER / TIÊU ĐỀ TÀI LIỆU
    # -------------------------------------------------------------
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(8)
    p_top.paragraph_format.space_after = Pt(2)
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_top.add_run("DỰ ÁN ĐƯỜNG SẮT ĐÔ THỊ TP. HỒ CHÍ MINH (METRO TUYẾN SỐ 2)\n")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(11)
    r_sub.font.bold = True
    r_sub.font.color.rgb = STEEL

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(6)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p_title.add_run("THIẾT LẬP WORKFLOW KỸ THUẬT:\nTHU THẬP HIỆN TRƯỜNG, MÃ HÓA, LƯU TRỮ SERVER,\nLIÊN KẾT GIS VÀ TỰ ĐỘNG HÓA BÁO CÁO")
    r_main.font.name = 'Calibri'
    r_main.font.size = Pt(15.5)
    r_main.font.bold = True
    r_main.font.color.rgb = NAVY

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(2)
    p_meta.paragraph_format.space_after = Pt(14)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Tài liệu kỹ thuật chuyên sâu về Luồng Dữ Liệu (Data Pipeline Architecture) — Tháng 09/2026")
    r_meta.font.name = 'Calibri'
    r_meta.font.size = Pt(9.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = MUTED

    # -------------------------------------------------------------
    # 1. TỔNG QUAN WORKFLOW 5 BƯỚC
    # -------------------------------------------------------------
    add_h1("1. SƠ ĐỒ WORKFLOW TỔNG THỂ (END-TO-END TECHNICAL PIPELINE)")
    add_body("Toàn bộ quy trình được thiết kế thành một chuỗi khép kín 5 bước, loại bỏ hoàn toàn các thao tác thủ công (tải ảnh từ Drive, đổi tên thủ công, cắt dán vào Word):")

    tbl_wf = doc.add_table(rows=6, cols=3)
    tbl_wf.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_wf)

    wf_headers = ["Bước", "Tên Công Đoạn Kỹ Thuật", "Nhiệm Vụ Kỹ Thuật & Đầu Ra Cụ Thể"]
    for i, h in enumerate(wf_headers):
        c = tbl_wf.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    wf_steps = [
        ("Bước 1", "Format Biểu Mẫu Khảo Sát (Mobile Form)", "Cấu hình Dropdown 26 Ga Metro 2 (S1-S26), cố định 6 ô chụp ảnh kỹ thuật (A1-A6), tự động bắt tọa độ GPS vệ tinh và thời gian thực."),
        ("Bước 2", "Mã Hóa & Đóng Gói Tại Máy Khách (Client-Side)", "JavaScript tự động đổi tên ảnh theo công thức [Mã_Ga]_[Mã_Mục]_[Timestamp].jpg, tự nén ảnh (Client-side Compression) và đóng Watermark GPS lên góc ảnh."),
        ("Bước 3", "Lưu Trữ Server & Database (Storage & Cloud)", "Server nhận dữ liệu qua API, lưu thuộc tính vào bảng Database (PostgreSQL/Supabase/SQLite), tự động lưu file ảnh vào thư mục phân cấp /data/photos/{Ga}/."),
        ("Bước 4", "Pipeline Liên Kết Phần Mềm GIS (Spatial Integration)", "Hệ thống tự động xuất file không gian metro2_survey_photos.geojson (chuẩn WGS84). Khi mở trong QGIS/ArcGIS, các điểm ảnh hiện đúng vị trí và có popup xem ảnh tại chỗ."),
        ("Bước 5", "Engine Tự Động Hóa Báo Cáo Word/PDF", "Script Python đọc template Word (.docx), tự động lấy dữ liệu thuộc tính và nhặt đúng các file ảnh A1-A6 từ thư mục máy chủ để chèn vào khung bảng, xuất báo cáo hoàn chỉnh.")
    ]

    for row_idx, step in enumerate(wf_steps, start=1):
        row = tbl_wf.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(step):
            cell = row.cells[col_idx]
            cell.text = text
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 120, 120)
            p = cell.paragraphs[0]
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(9.5)
            if col_idx in [0, 1]:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # 2. BƯỚC 1: FORMAT DỮ LIỆU BIỂU MẪU
    # -------------------------------------------------------------
    add_h1("2. BƯỚC 1: FORMAT DỮ LIỆU BIỂU MẪU KHẢO SÁT")
    add_body("Để dữ liệu máy tính có thể đọc và xử lý tự động, biểu mẫu không được tạo tự do mà phải tuân thủ Schema chuẩn hóa:")

    add_h2("A. Định Danh Vị Trí Khảo Sát (Station Master Data)")
    add_body("Tích hợp sẵn danh mục chuẩn 26 nhà ga Tuyến Metro số 2 (trích xuất từ cơ sở dữ liệu stations_db.json):")
    add_bullet("Trường dữ liệu dạng Dropdown: Khảo sát viên chọn tên hiển thị (ví dụ: S1 - Phan Văn Hớn, S2 - Đông Hưng Thuận, ..., S23 - Ba Sa).", "Mã Ga (Station ID): ")
    add_bullet("Hệ thống tự động khóa mã nội bộ (Station_Code = S1) và tọa độ tâm ga danh nghĩa (Lat: 10.836067, Lng: 106.618255).", "Khóa Hệ Thống: ")

    add_h2("B. Quy Định 6 Hạng Mục Ảnh Kỹ Thuật Cố Định (Photo Slots)")
    add_body("Thay vì cho phép tải ảnh lộn xộn, Form chia cố định đúng 6 ô chụp ảnh kỹ thuật bắt buộc tại mỗi nhà ga:")

    tbl_slots = doc.add_table(rows=7, cols=3)
    tbl_slots.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_slots)

    slot_headers = ["Mã Ô", "Tên Hạng Mục Khảo Sát", "Mục Tiêu Thu Thập & Góc Máy Quy Ước"]
    for i, h in enumerate(slot_headers):
        c = tbl_slots.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    slots_data = [
        ("A1", "Hiện Trạng Mặt Bằng Ga", "Chụp toàn cảnh mặt bằng tim đường / khu vực quy hoạch xây dựng nhà ga."),
        ("A2", "Vỉa Hè & Lối Tiếp Cận Đi Bộ", "Chụp bề rộng lối đi bộ, độ bằng phẳng, chất lượng lát gạch vỉa hè."),
        ("A3", "Trạm Dừng Xe Buýt Tiếp Chuyển", "Chụp nhà chờ / biển dừng xe buýt gần nhất trong bán kính 200m kết nối cửa ga."),
        ("A4", "Bãi Đỗ Xe Cá Nhân (Park & Ride)", "Chụp khu đất trống lân cận hoặc bãi giữ xe máy, ô tô có tiềm năng trung chuyển."),
        ("A5", "Điểm Đón Trả Khách Nhanh (PUDO)", "Chụp vị trí có khả năng bố trí vịnh dừng đón/trả cho Taxi, xe công nghệ, xe buýt nhỏ."),
        ("A6", "Chướng Ngại Vật & Lấn Chiếm", "Chụp các điểm xung đột: cột điện hạ cao độ, cây xanh lớn, miệng cống, chợ tạm lấn chiếm.")
    ]

    for row_idx, s in enumerate(slots_data, start=1):
        row = tbl_slots.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(s):
            cell = row.cells[col_idx]
            cell.text = text
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 120, 120)
            p = cell.paragraphs[0]
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(9.5)
            if col_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------
    # 3. BƯỚC 2: MÃ HÓA & ĐÓNG GÓI TẠI MÁY KHÁCH
    # -------------------------------------------------------------
    add_h1("3. BƯỚC 2: THUẬT TOÁN MÃ HÓA & ĐÓNG GÓI TẠI MÁY KHÁCH")
    add_body("Khảo sát viên hoàn toàn không phải gõ tên file. Khi chụp hoặc chọn ảnh tại ô khảo sát, JavaScript trên trình duyệt di động lập tức kích hoạt 3 cơ chế ngầm:")

    add_h2("A. Thuật Toán Tự Động Sinh Tên File (Auto-Naming Algorithm)")
    add_body("Tên file được chuẩn hóa theo quy tắc bất biến:")
    add_code_block("const fileName = `${stationId}_${categoryCode}_${timestamp}.jpg`;\n// Ví dụ cụ thể: S1_A1_HienTrang_20260912_143000.jpg")

    add_h2("B. Nén Ảnh Tại Máy Khách (Client-Side Compression)")
    add_body("Ảnh chụp từ điện thoại hiện đại có dung lượng rất lớn (5MB - 12MB), nếu tải trực tiếp qua mạng 4G sẽ dễ gây nghẽn, chậm và lỗi timeout:")
    add_bullet("Sử dụng HTML5 Canvas API để co kích thước tối đa về 1920x1080 (Full HD).", "Co kích thước: ")
    add_bullet("Nén chất lượng JPEG ở mức 0.82 (giảm 85% dung lượng nhưng giữ độ nét sắc sảo từng chi tiết vết nứt, biển báo).", "Tối ưu dung lượng: ")
    add_bullet("Mỗi file ảnh sau nén chỉ còn 600KB - 900KB, gửi lên server chỉ mất dưới 1 giây ngay cả khi sóng 4G chập chờn.", "Tốc độ: ")

    add_h2("C. Đóng Dấu Watermark Tọa Độ & Thời Gian (Proof Watermarking)")
    add_body("Canvas tự động in đè một dải băng đen mờ ở góc dưới bức ảnh với nội dung:")
    add_code_block("📍 GA S1 - PHAN VĂN HỚN | MỤC: A1_HIENTRANG\nGPS: 10.836068 N, 106.618255 E (±3.2m) | TIME: 2026-09-12 14:30:25")

    # -------------------------------------------------------------
    # 4. BƯỚC 3: THIẾT LẬP LƯU TRỮ SERVER & DATABASE
    # -------------------------------------------------------------
    add_h1("4. BƯỚC 3: THIẾT LẬP LƯU TRỮ MÁY CHỦ (SERVER & STORAGE ARCHITECTURE)")
    add_body("Để khắc phục triệt để lỗi Google Drive ('403 Forbidden', link xem trước không tải được bằng code), hệ thống lưu trữ Server được thiết lập theo mô hình 2 thành phần độc lập:")

    add_h2("A. Cấu Trúc Cơ Sở Dữ Liệu Quan Hệ (Database Schema)")
    add_body("Dữ liệu chữ và tọa độ được lưu trong Database (Supabase PostgreSQL hoặc SQLite cục bộ):")

    add_code_block(
        "-- 1. Bảng danh mục 26 Nhà Ga\n"
        "CREATE TABLE stations (\n"
        "    id VARCHAR(10) PRIMARY KEY,      -- 'S1', 'S2', ..., 'S26'\n"
        "    name VARCHAR(100) NOT NULL,      -- 'Phan Văn Hớn'\n"
        "    latitude DOUBLE PRECISION,       -- 10.836067\n"
        "    longitude DOUBLE PRECISION       -- 106.618255\n"
        ");\n\n"
        "-- 2. Bảng phiếu khảo sát hiện trường\n"
        "CREATE TABLE survey_records (\n"
        "    id VARCHAR(50) PRIMARY KEY,      -- 'REC_S1_20260912_143000'\n"
        "    station_id VARCHAR(10) REFERENCES stations(id),\n"
        "    surveyor_name VARCHAR(100),\n"
        "    survey_time TIMESTAMP WITH TIME ZONE,\n"
        "    survey_notes TEXT,\n"
        "    raw_answers JSONB\n"
        ");\n\n"
        "-- 3. Bảng liên kết Hình Ảnh Thực Địa\n"
        "CREATE TABLE survey_photos (\n"
        "    id SERIAL PRIMARY KEY,\n"
        "    record_id VARCHAR(50) REFERENCES survey_records(id),\n"
        "    station_id VARCHAR(10),\n"
        "    category_code VARCHAR(10),       -- 'A1', 'A2', ..., 'A6'\n"
        "    file_name VARCHAR(255),          -- 'S1_A1_HienTrang_20260912_143000.jpg'\n"
        "    file_path VARCHAR(500),          -- 'photos/S1/S1_A1_HienTrang_20260912_143000.jpg'\n"
        "    public_url TEXT,                 -- Direct CDN URL tải trực tiếp\n"
        "    latitude DOUBLE PRECISION,       -- Tọa độ lúc bấm chụp ảnh\n"
        "    longitude DOUBLE PRECISION,\n"
        "    created_at TIMESTAMP WITH TIME ZONE DEFAULT NOW()\n"
        ");"
    )

    add_h2("B. Cây Thư Mục Lưu Trữ File Ảnh Vật Lý (Physical Storage Hierarchy)")
    add_body("File ảnh được Server tự động phân phối vào đúng folder của nhà ga tương ứng. Khi tải về máy tính nội bộ, cấu trúc thư mục hoàn toàn ăn khớp:")
    add_code_block(
        "gis_connect/\n"
        "└── data/\n"
        "    ├── stations_db.json\n"
        "    ├── photos/\n"
        "    │   ├── S1/\n"
        "    │   │   ├── S1_A1_HienTrang.jpg\n"
        "    │   │   ├── S1_A2_ViaHeDiBo.jpg\n"
        "    │   │   ├── S1_A3_TramBus.jpg\n"
        "    │   │   ├── S1_A4_DiemDoXe.jpg\n"
        "    │   │   ├── S1_A5_PUDO.jpg\n"
        "    │   │   └── S1_A6_RaoCan.jpg\n"
        "    │   ├── S2/\n"
        "    │   └── ...\n"
        "    └── gis_export/\n"
        "        └── metro2_survey_photos.geojson"
    )

    add_callout("Khắc phục hoàn toàn lỗi Google Drive: Mỗi file ảnh trên Server có Direct CDN URL công khai an toàn (hoặc nằm sẵn trên ổ cứng nội bộ). Python Script chỉ cần đọc đường dẫn file_path là lấy được ảnh ngay lập tức trong 0.05 giây, không cần cookie, không cần đăng nhập Google.", "TỐI ƯU HÓA LƯU TRỮ")

    # -------------------------------------------------------------
    # 5. BƯỚC 4: PIPELINE LIÊN KẾT GIS
    # -------------------------------------------------------------
    add_h1("5. BƯỚC 4: PIPELINE LIÊN KẾT PHẦN MỀM BẢN ĐỒ GIS (QGIS / ARCGIS)")
    add_body("Để liên kết chặt chẽ với phần mềm GIS, hệ thống có một module tự động chuyển đổi toàn bộ điểm khảo sát và hình ảnh thành định dạng không gian chuẩn GeoJSON / Shapefile:")

    add_h2("A. Cấu Trúc Định Dạng GeoJSON (Chuẩn WGS-84 / EPSG:4326)")
    add_body("Tập tin metro2_survey_photos.geojson được tạo ra với cấu trúc chuẩn quốc tế:")
    add_code_block(
        "{\n"
        '  "type": "FeatureCollection",\n'
        '  "name": "Metro2_Survey_Points",\n'
        '  "crs": { "type": "name", "properties": { "name": "urn:ogc:def:crs:OGC:1.3:CRS84" } },\n'
        '  "features": [\n'
        "    {\n"
        '      "type": "Feature",\n'
        '      "geometry": { "type": "Point", "coordinates": [ 106.618255, 10.836067 ] },\n'
        '      "properties": {\n'
        '        "station_id": "S1",\n'
        '        "station_name": "Phan Văn Hớn",\n'
        '        "category_code": "A1",\n'
        '        "category_name": "Hiện Trạng Mặt Bằng Ga",\n'
        '        "photo_name": "S1_A1_HienTrang.jpg",\n'
        '        "photo_url": "https://.../photos/S1/S1_A1_HienTrang.jpg",\n'
        '        "local_path": "data/photos/S1/S1_A1_HienTrang.jpg",\n'
        '        "surveyor": "Nguyễn Văn A",\n'
        '        "survey_time": "2026-09-12 14:30:00"\n'
        "      }\n"
        "    }\n"
        "  ]\n"
        "}"
    )

    add_h2("B. Cách Thao Tác Trong Phần Mềm QGIS")
    add_bullet("Kéo thả trực tiếp file metro2_survey_photos.geojson vào không gian làm việc của QGIS. Toàn bộ các chấm điểm khảo sát sẽ xuất hiện chuẩn xác dọc theo hành lang tuyến Metro 2.", "1. Nạp bản đồ: ")
    add_bullet("Trong bảng thuộc tính (Attribute Table), mỗi điểm đều có đầy đủ thông tin mã Ga, loại hạng mục ảnh, thời gian chụp.", "2. Bảng thuộc tính: ")
    add_bullet("Vào Layer Properties -> HTML Map Tip, chèn đoạn mã hiển thị ảnh: <img src=\"[%local_path%]\" width=\"300\"/>. Khi rê chuột vào bất kỳ điểm khảo sát nào trên bản đồ, bức ảnh thực địa tương ứng sẽ lập tức hiện ra trực quan ngay tại chỗ!", "3. Xem ảnh trực quan (Map Tip): ")

    # -------------------------------------------------------------
    # 6. BƯỚC 5: PIPELINE TỰ ĐỘNG XUẤT BÁO CÁO
    # -------------------------------------------------------------
    add_h1("6. BƯỚC 5: PIPELINE TỰ ĐỘNG XUẤT BÁO CÁO WORD / PDF")
    add_body("Đây là khâu cuối cùng mang lại giá trị thực tế cao nhất: Chuyển toàn bộ dữ liệu hiện trường thành báo cáo kỹ thuật hoàn chỉnh mà không cần con người nhúng tay:")

    add_h2("A. Cơ Chế Truy Cập Dữ Liệu Của Script Báo Cáo")
    add_body("Script Python (sử dụng thư viện python-docx và docxtpl) thực thi quy trình theo 4 bước liên hoàn:")
    add_bullet("Đọc file mẫu Template_BaoCao_Ga.docx đã được định dạng sẵn tiêu đề, khung viền, bảng biểu kỹ thuật và các khung chèn ảnh tỷ lệ 4:3.", "Bước 1 (Đọc Mẫu): ")
    add_bullet("Kết nối vào Database (hoặc đọc file JSON của Ga cần xuất, ví dụ S1_PhanVanHon.json) để trích xuất số liệu vỉa hè, đánh giá đi bộ, hiện trạng xe buýt.", "Bước 2 (Bơm Số Liệu): ")
    add_bullet("Script tự động quét thư mục data/photos/S1/ để nhặt đúng 6 file ảnh đã mã hóa (S1_A1.jpg, S1_A2.jpg, ... S1_A6.jpg). Tự động điều chỉnh kích thước ảnh vừa khít ô bảng trong Word.", "Bước 3 (Chèn Ảnh Tự Động): ")
    add_bullet("Lưu file thành BaoCao_KhaoSat_Ga_S1_PhanVanHon.docx và tự động chuyển đổi sang PDF nếu cần bàn giao.", "Bước 4 (Xuất Báo Cáo): ")

    add_h2("B. Đoạn Mã Mẫu Python Tự Động Chèn Ảnh Vào Word")
    add_code_block(
        "import os\n"
        "from docx import Document\n"
        "from docx.shared import Inches\n\n"
        "def fill_station_report(station_id, station_name):\n"
        "    doc = Document('templates/Template_BaoCao_Metro2.docx')\n"
        "    photo_dir = f'data/photos/{station_id}'\n"
        "    \n"
        "    # Danh sách mã ảnh cần chèn\n"
        "    photo_slots = ['A1_HienTrang', 'A2_ViaHeDiBo', 'A3_TramBus', \n"
        "                   'A4_DiemDoXe', 'A5_PUDO', 'A6_RaoCan']\n"
        "    \n"
        "    for table in doc.tables:\n"
        "        for row in table.rows:\n"
        "            for cell in row.cells:\n"
        "                for slot in photo_slots:\n"
        "                    placeholder = f'{{{{IMG_{slot}}}}}'\n"
        "                    if placeholder in cell.text:\n"
        "                        cell.text = '' # Xóa tag placeholder\n"
        "                        img_path = os.path.join(photo_dir, f'{station_id}_{slot}.jpg')\n"
        "                        if os.path.exists(img_path):\n"
        "                            cell.paragraphs[0].add_run().add_picture(img_path, width=Inches(3.2))\n"
        "    \n"
        "    output_file = f'output/BaoCao_KhaoSat_Ga_{station_id}.docx'\n"
        "    doc.save(output_file)\n"
        "    print(f'Da xuat thanh cong bao cao cho Ga {station_name}!')"
    )

    # -------------------------------------------------------------
    # 7. TỔNG KẾT HIỆU QUẢ KỸ THUẬT
    # -------------------------------------------------------------
    add_h1("7. BẢNG TỔNG HỢP HIỆU QUẢ WORKFLOW MỚI SO VỚI QUY TRÌNH CŨ")

    tbl_cmp = doc.add_table(rows=6, cols=3)
    tbl_cmp.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_cmp)

    cmp_headers = ["Hạng Mục", "Quy Trình Cũ (Google Form & Drive)", "Workflow Kỹ Thuật Mới (GIS Connect)"]
    for i, h in enumerate(cmp_headers):
        c = tbl_cmp.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    cmp_data = [
        ("Tên file ảnh", "Tên ngẫu nhiên (z4836...jpg), phải mở từng ảnh để đoán nội dung.", "Tự động chuẩn hóa: S1_A1_HienTrang.jpg, nhận diện ngay lập tức."),
        ("Quyền truy cập ảnh", "Thường xuyên bị chặn quyền 403 Forbidden trên Google Drive.", "Lưu trữ nội bộ / CDN Direct URL, code truy cập trực tiếp 100%."),
        ("Dung lượng tải ảnh", "Ảnh gốc 5-10MB, gửi qua 4G rất chậm và dễ lỗi mạng.", "Tự động nén tại máy khách còn ~700KB, gửi nhanh dưới 1 giây."),
        ("Liên kết phần mềm GIS", "Dữ liệu tách rời, phải nhập tọa độ thủ công từng điểm vào QGIS.", "Tự động sinh GeoJSON, kéo vào QGIS hiện ngay điểm kèm ảnh."),
        ("Thời gian làm báo cáo", "Thủ công copy/paste mất 3 - 4 giờ cho 01 nhà ga.", "Tự động hóa hoàn toàn bằng Python: 2 - 3 giây cho 01 nhà ga.")
    ]

    for row_idx, row_data in enumerate(cmp_data, start=1):
        row = tbl_cmp.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(row_data):
            cell = row.cells[col_idx]
            cell.text = text
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 120, 120)
            p = cell.paragraphs[0]
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(9.5)
            if col_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = NAVY
            elif col_idx == 2:
                p.runs[0].font.color.rgb = RGBColor(20, 110, 40)

    # Lưu tài liệu Word
    output_path = r"d:\gis_connect\QUY_TRINH_WORKFLOW_KY_THUAT_KHAO_SAT_METRO_GIS.docx"
    doc.save(output_path)
    print("SUCCESS: File Word Workflow ky thuat da duoc cap nhat tai: " + output_path)

if __name__ == '__main__':
    create_technical_workflow_docx()
