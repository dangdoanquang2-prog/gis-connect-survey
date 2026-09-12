import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

def create_workflow_docx():
    doc = docx.Document()

    # Cấu hình Margins trang A4 chuẩn
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # Bảng màu chủ đạo (Executive Corporate Navy)
    NAVY = RGBColor(0, 51, 102)        # #003366 - Tiêu đề chính
    SLATE = RGBColor(70, 130, 180)     # #4682B4 - Tiêu đề phụ
    CHARCOAL = RGBColor(40, 40, 40)    # Văn bản thông thường
    MUTED_GRAY = RGBColor(100, 100, 100)
    HEX_HEADER_BG = "003366"
    HEX_SUBHEADER_BG = "E6EEF8"
    HEX_ROW_LIGHT = "F8FAFC"
    HEX_BORDER = "CCCCCC"

    def set_cell_background(cell, hex_color):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
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
        run.font.size = Pt(14)
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
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = SLATE
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

    def add_callout(text, title="LƯU Ý KỸ THUẬT QUAN TRỌNG"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.rows[0].cells[0]
        set_cell_background(cell, "F0F4F8")
        set_cell_margins(cell, top=140, bottom=140, left=200, right=180)
        
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
        p.paragraph_format.space_after = Pt(3)
        r_title = p.add_run(f"📌 {title}: ")
        r_title.font.name = 'Calibri'
        r_title.font.size = Pt(10.5)
        r_title.font.bold = True
        r_title.font.color.rgb = NAVY
        
        r_txt = p.add_run(text)
        r_txt.font.name = 'Calibri'
        r_txt.font.size = Pt(10)
        r_txt.font.color.rgb = CHARCOAL
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # -------------------------------------------------------------
    # 1. TRANG TIÊU ĐỀ (HEADER BLOCK)
    # -------------------------------------------------------------
    p_top = doc.add_paragraph()
    p_top.paragraph_format.space_before = Pt(10)
    p_top.paragraph_format.space_after = Pt(2)
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_top.add_run("DỰ ÁN ĐƯỜNG SẮT ĐÔ THỊ TP. HỒ CHÍ MINH (METRO TUYẾN SỐ 2)\n")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(11)
    r_sub.font.bold = True
    r_sub.font.color.rgb = SLATE

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(8)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_main = p_title.add_run("QUY TRÌNH KỸ THUẬT THIẾT LẬP HỆ THỐNG\nTHU THẬP HIỆN TRƯỜNG, LƯU TRỮ SERVER, TÍCH HỢP GIS\nVÀ TỰ ĐỘNG HÓA XUẤT BÁO CÁO")
    r_main.font.name = 'Calibri'
    r_main.font.size = Pt(16)
    r_main.font.bold = True
    r_main.font.color.rgb = NAVY

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(2)
    p_meta.paragraph_format.space_after = Pt(16)
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run("Tài liệu hướng dẫn phối hợp & Thiết lập hệ thống chuẩn hóa (Version 2.0 - Tháng 09/2026)")
    r_meta.font.name = 'Calibri'
    r_meta.font.size = Pt(9.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = MUTED_GRAY

    # Bảng phân công nhân sự đầu trang
    tbl_team = doc.add_table(rows=4, cols=3)
    tbl_team.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_team)
    
    headers_team = ["Nhân Sự Phụ Trách", "Vai Trò & Nhiệm Vụ Chính", "Phạm Vi Phối Hợp Kỹ Thuật"]
    for i, h in enumerate(headers_team):
        c = tbl_team.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 120, 120)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    team_data = [
        ("Dũng & Anh Đạt", "Thiết lập Hệ thống & Hạ tầng Server", "Cài đặt Web App PWA di động, kết nối Supabase Cloud Database, cấu hình cơ chế lưu trữ ảnh và phân quyền form."),
        ("Anh Sơn", "Chuyên gia GIS & Dữ liệu Không gian", "Chuẩn hóa hệ tọa độ WGS84/VN-2000, kiểm soát lớp bản đồ 26 ga, tích hợp file GeoJSON/KMZ vào QGIS và tạo bản đồ bán kính 200m/1000m."),
        ("Hệ thống Tự động (AI/Pipeline)", "Điều phối Dữ liệu & Báo cáo Tự động", "Mã hóa tự động tên ảnh thực địa, xuất file GeoJSON tương thích GIS, tự động bơm dữ liệu + hình ảnh vào template Word/PDF.")
    ]

    for row_idx, data in enumerate(team_data, start=1):
        row = tbl_team.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(data):
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

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # 2. PHẦN I: BỐI CẢNH & NGUYÊN NHÂN CẢI TIẾN
    # -------------------------------------------------------------
    add_h1("I. BỐI CẢNH & PHÂN TÍCH HIỆN TRẠNG (BÀI HỌC TỪ DỰ ÁN TRƯỚC)")
    add_body("Qua quá trình rà soát dữ liệu thực tế tại các thư mục triển khai trước đây (formkhaosat, baocaotudong, baocaometro2s), nhóm kỹ thuật nhận thấy các nút thắt cốt lõi làm tiêu tốn từ 70% đến 80% thời gian nhân sự văn phòng:")
    
    add_bullet("Ảnh bị dồn chung vào một thư mục Google Drive với tên file ngẫu nhiên (ví dụ: z4836087779256_...jpg hoặc mã hash của Google). Khi làm báo cáo, nhân viên phải mở từng ảnh để đoán xem đó là mặt bằng, trạm xe buýt hay vỉa hè của Ga nào.", "1. Tên ảnh vô nghĩa: ")
    add_bullet("Google Drive chặn quyền truy cập trực tiếp bằng API (lỗi 403 Forbidden). Các liên kết lưu trên Google Sheet chỉ là trang xem trước web, không thể dùng script lập trình tự động tải hoặc nhúng ảnh vào file báo cáo.", "2. Khóa quyền Google Drive: ")
    add_bullet("Dữ liệu khảo sát trên Sheet tách rời hoàn toàn với file bản đồ không gian (.kmz, .shp). Không có cơ chế tự động hiển thị vị trí bức ảnh lên phần mềm bản đồ GIS.", "3. Rời rạc với GIS: ")
    add_bullet("Người làm báo cáo phải làm thủ công 100%: sao chép số liệu, tải ảnh, cắt dán kích thước ảnh vào Word, căn chỉnh lề vô cùng mất thời gian và dễ nhầm lẫn.", "4. Báo cáo thủ công: ")

    add_callout("Mục tiêu của quy trình mới: Khảo sát viên tại hiện trường chỉ cần mở App trên điện thoại -> Chọn Ga -> Chụp các ô ảnh theo quy ước -> Bấm Gửi. Mọi công đoạn mã hóa tên file, lưu trữ phân cấp, cập nhật GIS và sinh file báo cáo Word/PDF đều được máy tính thực hiện tự động 100%.", "MỤC TIÊU CỐT LÕI")

    # -------------------------------------------------------------
    # 3. PHẦN II: SƠ ĐỒ QUY TRÌNH KHÉP KÍN 5 GIAI ĐOẠN
    # -------------------------------------------------------------
    add_h1("II. QUY TRÌNH KHÉP KÍN 5 GIAI ĐOẠN (END-TO-END WORKFLOW)")
    add_body("Quy trình kỹ thuật được chuẩn hóa thành 5 giai đoạn liên hoàn, mỗi bộ phận giữ vai trò then chốt:")

    tbl_flow = doc.add_table(rows=6, cols=4)
    tbl_flow.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_flow)

    flow_headers = ["Giai Đoạn", "Tên Bước", "Nội Dung Thực Hiện", "Đầu Ra (Deliverable)"]
    for i, h in enumerate(flow_headers):
        c = tbl_flow.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 100, 100)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    flow_steps = [
        ("Giai đoạn 1", "Thiết lập Form & Hạ tầng", "Dũng và anh Đạt tích hợp danh mục 26 ga vào App PWA; thiết lập cấu trúc bảng Database và Bucket lưu ảnh.", "App chạy trên web/mobile, form khảo sát sẵn sàng."),
        ("Giai đoạn 2", "Thu thập Thực địa", "Khảo sát viên đứng tại nhà ga mở App trên điện thoại, tự động bắt GPS, chụp 6 góc ảnh quy chuẩn và gửi bài.", "Dữ liệu khảo sát + Ảnh gốc được nén tại máy."),
        ("Giai đoạn 3", "Mã hóa & Lưu trữ Server", "Hệ thống tự động đổi tên ảnh theo công thức [Mã_Ga]_[Mã_Ảnh]_[GPS].jpg và xếp vào đúng thư mục của Ga đó trên Server.", "Ảnh lưu có thứ tự tại data/photos/{Ga}/, dữ liệu lưu JSON/Cloud."),
        ("Giai đoạn 4", "Liên kết Không gian GIS", "Anh Sơn nhận file GeoJSON/Shapefile tự động xuất ra từ hệ thống, tích hợp vào dự án QGIS hiển thị toàn tuyến Metro 2.", "Lớp bản đồ GIS chứa các điểm chụp ảnh có Popup xem ảnh."),
        ("Giai đoạn 5", "Tự động Xuất Báo Cáo", "Script Python quét dữ liệu và ảnh của từng Ga, đổ vào Template Word mẫu (.docx), sinh báo cáo hoàn chỉnh.", "Báo cáo Word/PDF đầy đủ biểu bảng và ảnh hiện trường.")
    ]

    for row_idx, step in enumerate(flow_steps, start=1):
        row = tbl_flow.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(step):
            cell = row.cells[col_idx]
            cell.text = text
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 100, 100)
            p = cell.paragraphs[0]
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(9.5)
            if col_idx in [0, 1]:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # 4. PHẦN III: KIẾN TRÚC LƯU TRỮ SERVER (SERVER STORAGE ARCHITECTURE)
    # -------------------------------------------------------------
    add_h1("III. THIẾT LẬP LƯU TRỮ MÁY CHỦ (SERVER STORAGE ARCHITECTURE)")
    add_body("Hệ thống lưu trữ được thiết kế phân tầng rõ rệt: Tách biệt hoàn toàn giữa dữ liệu thuộc tính (Database) và file đa phương tiện (Object Storage), nhưng được khóa chặt chẽ với nhau thông qua Định danh Duy nhất (Station ID & Photo UID):")

    add_h2("1. Mô Hình Lưu Trữ Đám Mây (Supabase Cloud + Local Storage Fallback)")
    add_bullet("Lưu trữ cấu trúc 26 nhà ga, câu trả lời thuộc tính, đánh giá hiện trạng, thông tin khảo sát viên và tọa độ GPS thực tế. Hệ thống hỗ trợ truy vấn nhanh theo từng Ga.", "Bảng Dữ liệu (PostgreSQL Database): ")
    add_bullet("Kho lưu trữ file ảnh chuyên dụng. Mỗi ảnh được cấp một đường dẫn Direct CDN URL vĩnh viễn, cho phép script Python ở văn phòng kéo về cực nhanh mà không bị lỗi phân quyền.", "Kho Tệp Đa Phương Tiện (Storage Bucket): ")
    add_bullet("Trường hợp khảo sát viên vào vùng mất sóng 4G (tầng hầm, nhà ga ngầm), App tự động lưu tạm dữ liệu vào bộ nhớ máy (IndexedDB). Khi có sóng trở lại, hệ thống tự động đẩy ngầm (Auto-sync) lên Server.", "Cơ chế Dự phòng Offline: ")

    add_h2("2. Cây Thư Mục Lưu Trữ Phân Cấp Trên Server / Máy Tính Văn Phòng")
    add_body("Toàn bộ dữ liệu và ảnh tải về được tự động tổ chức ngăn nắp theo cây thư mục sau:")

    code_tree = (
        "gis_connect/\n"
        "├── data/\n"
        "│   ├── stations_db.json              <-- Danh mục chuẩn 26 Ga (Tên, ID, Tọa độ gốc)\n"
        "│   ├── surveys/                      <-- Dữ liệu thuộc tính khảo sát từng ga\n"
        "│   │   ├── S1_PhanVanHon.json\n"
        "│   │   ├── S2_DongHungThuan.json\n"
        "│   │   └── ...\n"
        "│   ├── photos/                       <-- Kho ảnh thực địa đã được mã hóa chuẩn\n"
        "│   │   ├── S1/\n"
        "│   │   │   ├── S1_A1_HienTrang_10.8360_106.6182.jpg\n"
        "│   │   │   ├── S1_A2_ViaHeDiBo_10.8361_106.6185.jpg\n"
        "│   │   │   ├── S1_A3_TramBus_10.8365_106.6190.jpg\n"
        "│   │   │   ├── S1_A4_DiemDoXe_10.8358_106.6179.jpg\n"
        "│   │   │   ├── S1_A5_PUDO_10.8362_106.6183.jpg\n"
        "│   │   │   └── S1_A6_RaoCan_10.8364_106.6181.jpg\n"
        "│   │   ├── S2/\n"
        "│   │   └── ...\n"
        "│   └── gis_export/                   <-- Dữ liệu không gian cung cấp cho Anh Sơn (GIS)\n"
        "│       ├── metro2_stations_core.geojson\n"
        "│       └── metro2_survey_photos.geojson"
    )
    p_code = doc.add_paragraph()
    p_code.paragraph_format.left_indent = Inches(0.2)
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(6)
    r_code = p_code.add_run(code_tree)
    r_code.font.name = 'Consolas'
    r_code.font.size = Pt(8.5)
    r_code.font.color.rgb = RGBColor(30, 70, 32)

    # -------------------------------------------------------------
    # 5. PHẦN IV: QUY CHUẨN FORM & MÃ HÓA HÌNH ẢNH HIỆN TRƯỜNG
    # -------------------------------------------------------------
    add_h1("IV. QUY CHUẨN FORM KHẢO SÁT & MÃ HÓA TỰ ĐỘNG")
    add_body("Để triệt tiêu lỗi nhầm lẫn ảnh, biểu mẫu khảo sát di động được cấu hình cố định 6 hạng mục kỹ thuật bắt buộc tại mỗi nhà ga:")

    tbl_photo_specs = doc.add_table(rows=7, cols=4)
    tbl_photo_specs.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_photo_specs)

    spec_headers = ["Mã Hạng Mục", "Tên Quy Ước", "Mô Tả Kỹ Thuật Khi Chụp Ảnh", "Mẫu Tên File Tự Động Sinh"]
    for i, h in enumerate(spec_headers):
        c = tbl_photo_specs.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 100, 100)
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(10)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    specs = [
        ("A1", "Hiện Trạng Mặt Bằng", "Chụp toàn cảnh khu đất / tim đường vị trí quy hoạch xây dựng nhà ga.", "S1_A1_HienTrang.jpg"),
        ("A2", "Vỉa Hè Tiếp Cận Đi Bộ", "Chụp hiện trạng vỉa hè, bề rộng lối đi, độ dốc và tình trạng lát gạch.", "S1_A2_ViaHeDiBo.jpg"),
        ("A3", "Trạm Dừng Xe Buýt", "Chụp nhà chờ / biển dừng xe buýt gần nhất kết nối với cửa ga (bán kính 200m).", "S1_A3_TramBus.jpg"),
        ("A4", "Bãi Đỗ Xe Cá Nhân (P&R)", "Chụp khu đất trống, bãi xe hiện hữu hoặc bãi giữ xe máy/ô tô lân cận.", "S1_A4_DiemDoXe.jpg"),
        ("A5", "Điểm Đón Trả Khách (PUDO)", "Chụp các vị trí có khả năng bố trí vịnh dừng đón trả khách cho Taxi/Grab.", "S1_A5_PUDO.jpg"),
        ("A6", "Chướng Ngại Vật & Lấn Chiếm", "Chụp các điểm nghẽn: cột điện hạ cao độ, cây xanh lớn, chợ tạm, lấn chiếm vỉa hè.", "S1_A6_RaoCan.jpg")
    ]

    for row_idx, spec in enumerate(specs, start=1):
        row = tbl_photo_specs.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(spec):
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
            elif col_idx == 3:
                p.runs[0].font.name = 'Consolas'
                p.runs[0].font.size = Pt(8.5)
                p.runs[0].font.color.rgb = RGBColor(180, 50, 50)

    add_body("Thuật toán tự động đóng Watermark lên ảnh:", bold_prefix="Cơ chế bảo toàn tính pháp lý: ")
    add_bullet("Thời gian thực (Timestamp: YYYY-MM-DD HH:MM:SS)", "1. ")
    add_bullet("Tọa độ vệ tinh chính xác tại thời điểm bấm máy (Lat, Long, Accuracy ±m)", "2. ")
    add_bullet("Mã Ga và Mã hạng mục kỹ thuật (ví dụ: GA S1 - A1)", "3. ")

    # -------------------------------------------------------------
    # 6. PHẦN V: MA TRẬN PHỐI HỢP & PHÂN CÔNG CÔNG VIỆC
    # -------------------------------------------------------------
    add_h1("V. MA TRẬN PHỐI HỢP NHÂN SỰ & CHECKLIST TRIỂN KHAI")
    add_body("Để đảm bảo dự án vận hành mượt mà và không bị chồng chéo trách nhiệm, ma trận RACI được phân định rõ cho các thành sự chủ chốt:")

    tbl_raci = doc.add_table(rows=7, cols=5)
    tbl_raci.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl_raci)

    raci_headers = ["Hạng Mục Công Việc", "Dũng & Anh Đạt\n(Hệ Thống)", "Anh Sơn\n(Chuyên Gia GIS)", "Khảo Sát Viên\n(Hiện Trường)", "Hệ Thống Tự Động\n(AI Pipeline)"]
    for i, h in enumerate(raci_headers):
        c = tbl_raci.rows[0].cells[i]
        c.text = h
        set_cell_background(c, HEX_HEADER_BG)
        set_cell_margins(c, 100, 100, 80, 80)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraphs[0].runs[0].font.name = 'Calibri'
        c.paragraphs[0].runs[0].font.size = Pt(9.5)
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)

    raci_tasks = [
        ("1. Chuẩn bị danh mục 26 Ga & Tọa độ", "Phối hợp", "Chịu trách nhiệm (R)", "Hỗ trợ", "Tự động nạp DB"),
        ("2. Cấu hình Form PWA & Tự động mã hóa ảnh", "Chịu trách nhiệm (R)", "Tư vấn trường dữ liệu", "Thử nghiệm form", "Thực thi thuật toán"),
        ("3. Cấu hình Server Storage & Cloud Bucket", "Chịu trách nhiệm (R)", "Tham vấn cấu trúc", "Không tham gia", "Tự động phân luồng"),
        ("4. Thu thập số liệu và chụp ảnh thực địa", "Giám sát kỹ thuật", "Kiểm tra GPS hiện trường", "Chịu trách nhiệm (R)", "Nén & Watermark"),
        ("5. Xử lý dữ liệu không gian & Lớp bản đồ", "Hỗ trợ xuất GeoJSON", "Chịu trách nhiệm (R)", "Không tham gia", "Xuất file GIS tự động"),
        ("6. Xuất Báo Cáo Word/PDF tự động", "Phối hợp Template", "Duyệt bản đồ nhúng", "Cung cấp ghi chú", "Chịu trách nhiệm (R)")
    ]

    for row_idx, task in enumerate(raci_tasks, start=1):
        row = tbl_raci.rows[row_idx]
        bg = HEX_ROW_LIGHT if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, text in enumerate(task):
            cell = row.cells[col_idx]
            cell.text = text
            set_cell_background(cell, bg)
            set_cell_margins(cell, 80, 80, 80, 80)
            p = cell.paragraphs[0]
            p.runs[0].font.name = 'Calibri'
            p.runs[0].font.size = Pt(9)
            if col_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = NAVY
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if "Chịu trách nhiệm" in text:
                    p.runs[0].font.bold = True
                    p.runs[0].font.color.rgb = RGBColor(160, 40, 40)

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # 7. PHẦN VI: TÍCH HỢP GIS & TỰ ĐỘNG HÓA BÁO CÁO
    # -------------------------------------------------------------
    add_h1("VI. LIÊN KẾT PHẦN MỀM GIS & TỰ ĐỘNG XUẤT BÁO CÁO")
    
    add_h2("1. Tích Hợp Phần Mềm Bản Đồ GIS (Dành riêng cho Anh Sơn)")
    add_body("Hệ thống giải quyết triệt để yêu cầu liên kết GIS bằng cách tự động sinh file định dạng không gian chuẩn GeoJSON / Shapefile:")
    add_bullet("Tập tin metro2_survey_photos.geojson chứa toàn bộ các điểm khảo sát kèm tọa độ chuẩn WGS-84 (EPSG:4326). Anh Sơn chỉ cần kéo thả trực tiếp vào phần mềm QGIS hoặc ArcGIS.", "Định dạng Chuẩn: ")
    add_bullet("Mỗi điểm trên bản đồ GIS chứa đầy đủ thuộc tính: Mã Ga, Loại Hạng Mục (A1-A6), Tên Khảo Sát Viên, Thời Gian Khảo Sát và Đường Dẫn Ảnh Thực Địa.", "Bảng Thuộc Tính Đầy Đủ (Attributes): ")
    add_bullet("Trong QGIS, cấu hình tính năng HTML Map Tip hoặc Action: Khi người dùng rê chuột hoặc bấm vào điểm khảo sát trên bản đồ, bức ảnh thực địa tương ứng sẽ tự động hiển thị trực quan ngay trên giao diện bản đồ.", "Hiển Thị Ảnh Trực Tiếp Trên Bản Đồ: ")

    add_h2("2. Cơ Chế Tự Động Hóa Báo Cáo Word (.docx)")
    add_body("Kế thừa thành công từ module tạo báo cáo Ga S1 và Ga S23 (trong thư mục baocaotudong):")
    add_bullet("File Word mẫu được thiết kế sẵn khung viền chuẩn, tiêu đề bộ môn, vị trí đặt bảng số liệu và các khung ảnh chữ nhật tỷ lệ 4:3.", "Template Mẫu Chuẩn: ")
    add_bullet("Script Python quét thư mục data/surveys/{Ga}.json và data/photos/{Ga}/ -> Bơm trực tiếp các thông số kỹ thuật vào bảng biểu -> Tự động chèn 6 bức ảnh A1 đến A6 vào đúng các ô trong báo cáo.", "Nhúng Ảnh & Bảng Tự Động: ")
    add_bullet("Thời gian xuất 01 bản báo cáo hoàn chỉnh cho 01 nhà ga là 2 - 3 giây. Có thể xuất hàng loạt 26 Ga chỉ trong vòng chưa đầy 1 phút.", "Tốc Độ & Hiệu Suất: ")

    # -------------------------------------------------------------
    # 8. KẾT LUẬN & ĐỀ XUẤT BƯỚC TIẾP THEO
    # -------------------------------------------------------------
    add_h1("VII. KẾ HOẠCH HÀNH ĐỘNG NGAY BÂY GIỜ")
    add_bullet("Cùng Dũng và anh Đạt kiểm tra kết nối Supabase, đưa App lên link Vercel để các bên mở thử nghiệm.", "Bước 1: ")
    add_bullet("Chuyển giao file quy chuẩn này cho anh Sơn để thống nhất danh mục 26 ga và hệ tọa độ bản đồ GIS.", "Bước 2: ")
    add_bullet("Thử nghiệm khảo sát mẫu tại 01 Ga (Ga S1 Phan Văn Hớn): chụp 6 ảnh theo chuẩn -> Kiểm tra tên file mã hóa -> Kéo vào QGIS xem bản đồ -> Xuất thử file Báo Cáo Word.", "Bước 3: ")

    # Lưu file
    output_path = r"d:\gis_connect\QUY_TRINH_THIET_LAP_HE_THONG_KHAO_SAT_METRO_GIS.docx"
    doc.save(output_path)
    print("SUCCESS: File Word da duoc tao thanh cong tai: " + output_path)

if __name__ == '__main__':
    create_workflow_docx()
