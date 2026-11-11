import re
import unicodedata
from io import BytesIO
from pathlib import Path
from collections import defaultdict, Counter
from copy import copy

import streamlit as st
from docx import Document
from docx.oxml.ns import qn
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Font


# =====================================================
# CẤU HÌNH
# =====================================================

st.set_page_config(
    page_title="Chuyển Word sang Excel",
    page_icon="📚",
    layout="centered"
)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = BASE_DIR / "default_template.xlsx"


# =====================================================
# HÀM XỬ LÝ VĂN BẢN
# =====================================================

def clean_text(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize(value):
    value = clean_text(value).lower()
    value = unicodedata.normalize("NFD", value)
    value = "".join(
        char for char in value
        if unicodedata.category(char) != "Mn"
    )
    return value.replace("đ", "d")


def get_cell_text(cell):
    try:
        return clean_text(
            " ".join(paragraph.text for paragraph in cell.paragraphs)
        )
    except Exception:
        return ""


def get_period(value):
    """Đọc ô số tiết dạng 2, 2 tiết; không nhận chữ hoặc số trang."""
    value = clean_text(value)
    match = re.fullmatch(r"(\d+)\s*(?:tiết)?", value, re.IGNORECASE)
    return int(match.group(1)) if match else None


def title_case_vietnamese(value):
    """Ví dụ: TIẾNG VIỆT -> Tiếng Việt."""
    value = clean_text(value)
    return value.title() if value else ""


def extract_week_number(value):
    text = normalize(value)
    match = re.match(r"^\s*(?:tuan\s*)?(\d{1,2})\b", text)
    if not match:
        return None

    number = int(match.group(1))
    return number if 1 <= number <= 52 else None


def cell_is_vmerge_continue(cell):
    """
    Kiểm tra ô Word có phải phần tiếp tục của ô gộp dọc không.
    Trong Word, ô tiếp tục thường có <w:vMerge/> hoặc val='continue'.
    """
    try:
        tc_pr = cell._tc.tcPr
        if tc_pr is None:
            return False

        v_merge = tc_pr.find(qn("w:vMerge"))
        if v_merge is None:
            return False

        value = v_merge.get(qn("w:val"))
        return value in (None, "continue")
    except Exception:
        return False


def cell_is_vmerge_restart(cell):
    try:
        tc_pr = cell._tc.tcPr
        if tc_pr is None:
            return False
        v_merge = tc_pr.find(qn("w:vMerge"))
        return (
            v_merge is not None
            and v_merge.get(qn("w:val")) == "restart"
        )
    except Exception:
        return False


# =====================================================
# ĐỌC LỚP, MÔN, PHÂN MÔN TỪ WORD
# =====================================================

def extract_word_metadata(file_bytes):
    document = Document(BytesIO(file_bytes))
    text_parts = []

    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            text_parts.append(paragraph.text)

    for table in document.tables:
        for row in table.rows:
            # Dùng unique cell để tránh lặp nội dung của ô gộp.
            seen = set()
            for cell in row.cells:
                cell_id = id(cell._tc)
                if cell_id in seen:
                    continue
                seen.add(cell_id)
                cell_text = get_cell_text(cell)
                if cell_text:
                    text_parts.append(cell_text)

    full_text = "\n".join(text_parts)

    metadata = {
        "lop": "",
        "mon": "",
        "phan_mon": ""
    }

    # Trường hợp tiêu đề như: MÔN TIẾNG VIỆT LỚP 4.
    combined = re.search(
        r"\bmôn\s+(.+?)\s+lớp\s*(\d+)\b",
        full_text,
        re.IGNORECASE
    )

    if combined:
        metadata["mon"] = title_case_vietnamese(combined.group(1))
        metadata["lop"] = combined.group(2)

    # Trường hợp lớp và môn được ghi riêng trong bảng thông tin.
    if not metadata["lop"]:
        class_match = re.search(
            r"\blớp\s*[:：-]?\s*(\d+)\b",
            full_text,
            re.IGNORECASE
        )
        if class_match:
            metadata["lop"] = class_match.group(1)

    if not metadata["mon"]:
        subject_match = re.search(
            r"\bmôn\s*[:：-]?\s*(.+?)(?=\s+lớp\s*\d+|[\r\n]|$)",
            full_text,
            re.IGNORECASE
        )
        if subject_match:
            subject = clean_text(subject_match.group(1))
            # Không lấy nhầm phần còn lại của tiêu đề/bảng.
            metadata["mon"] = title_case_vietnamese(subject)

    sub_match = re.search(
        r"phân\s*môn\s*[:：-]?\s*([^\r\n]+)",
        full_text,
        re.IGNORECASE
    )
    if sub_match:
        metadata["phan_mon"] = clean_text(sub_match.group(1))

    return metadata


# =====================================================
# TÌM BẢNG VÀ CÁC CỘT TRONG WORD
# =====================================================

def find_word_program_table(document):
    for table in document.tables:
        for row_index, row in enumerate(table.rows):
            values = [normalize(get_cell_text(cell)) for cell in row.cells]

            week_col = next(
                (
                    i for i, value in enumerate(values)
                    if value == "tuan" or "tuan ppct" in value
                ),
                None
            )

            content_col = next(
                (
                    i for i, value in enumerate(values)
                    if "noi dung" in value or "ten bai" in value
                ),
                None
            )

            period_col = next(
                (
                    i for i, value in enumerate(values)
                    if "so tiet" in value
                ),
                None
            )

            if (
                week_col is not None
                and content_col is not None
                and period_col is not None
            ):
                return table, row_index, week_col, content_col, period_col

    raise ValueError(
        "Không tìm thấy bảng có đủ cột Tuần, Nội dung và Số tiết. "
        "Hãy kiểm tra hàng tiêu đề trong file Word."
    )



# =====================================================
# ĐỌC BẢNG WORD - HỖ TRỢ BẢNG BỊ NGẮT SANG TRANG MỚI
# =====================================================

def read_word(file_bytes):

    document = Document(BytesIO(file_bytes))

    if not document.tables:
        raise ValueError(
            "File Word không có bảng dữ liệu chương trình."
        )

    weekly_data = defaultdict(list)

    current_week = None
    started = False

    week_col = None
    content_col = None
    period_col = None

    # Đọc TẤT CẢ bảng trong Word, không chỉ bảng đầu tiên.
    for table in document.tables:

        header_row = None

        # Tìm dòng tiêu đề nếu bảng có tiêu đề.
        for row_index, row in enumerate(table.rows):

            values = [
                normalize(get_cell_text(cell))
                for cell in row.cells
            ]

            w = next(
                (
                    i for i, v in enumerate(values)
                    if v == "tuan" or "tuan ppct" in v
                ),
                None
            )

            c = next(
                (
                    i for i, v in enumerate(values)
                    if "noi dung" in v or "ten bai" in v
                ),
                None
            )

            p = next(
                (
                    i for i, v in enumerate(values)
                    if "so tiet" in v
                ),
                None
            )

            if w is not None and c is not None and p is not None:

                week_col = w
                content_col = c
                period_col = p

                header_row = row_index
                started = True

                break

        # Chưa tìm thấy bảng chương trình thì bỏ qua.
        if not started:
            continue

        # Nếu bảng có tiêu đề thì đọc từ dòng tiếp theo.
        # Nếu là bảng tiếp tục ở trang sau thì đọc từ dòng đầu.
        start_index = (
            header_row + 1
            if header_row is not None
            else 0
        )

        for row in table.rows[start_index:]:

            cells = [
                get_cell_text(cell)
                for cell in row.cells
            ]

            if len(cells) <= max(
                week_col,
                content_col,
                period_col
            ):
                continue

            week_text = cells[week_col]
            title = cells[content_col]
            period_text = cells[period_col]

            # Nhận diện số tuần.
            # Nếu ô tuần bị gộp và để trống ở trang sau,
            # giữ nguyên tuần đã nhận diện trước đó.
            week_match = re.match(
                r"^\s*(\d{1,2})\b",
                week_text
            )

            if week_match:

                number = int(week_match.group(1))

                if 1 <= number <= 52:
                    current_week = number

            # Không có tuần thì không ghi nội dung.
            if current_week is None:
                continue

            if not title:
                continue

            title_normalized = normalize(title)

            period_count = get_period(period_text)

            # Nhận diện ôn tập, kiểm tra, đánh giá.
            is_review = any(
                word in title_normalized
                for word in [
                    "on tap",
                    "danh gia",
                    "kiem tra"
                ]
            )

            # Nội dung không có số tiết:
            # Chỉ nhận ôn tập nếu dòng đó có số tuần cụ thể.
            # Nếu không có số tuần thì bỏ qua, không gán vào tuần trước.
            if period_count is None:

                if not is_review or not week_match:
                    continue
            item_type = (
                "review"
                if is_review
                else "lesson"
            )

            weekly_data[current_week].append({
                "title": title,
                "periods": period_count,
                "type": item_type
            })

    if not weekly_data:
        raise ValueError(
            "Không nhận diện được tuần và nội dung trong Word."
        )

    return dict(weekly_data)



# =====================================================
# TỰ NHẬN DIỆN ĐỊNH MỨC SỐ TIẾT/TUẦN
# =====================================================

def detect_period_standard(weekly_data):
    """
    Xác định định mức bằng số tiết xuất hiện phổ biến nhất (mode),
    KHÔNG lấy tuần có số tiết lớn nhất làm chuẩn.

    Chỉ dùng các tuần học bình thường:
    - Có bài học và số tiết rõ ràng.
    - Không có dòng ôn tập/kiểm tra riêng.
    - Không có nội dung thiếu số tiết.

    Nếu có nhiều mức cùng xuất hiện nhiều nhất, không tự đoán định mức
    mà báo lỗi để người dùng kiểm tra dữ liệu Word.
    """
    normal_week_totals = {}

    for week, entries in weekly_data.items():
        lessons = [
            item for item in entries
            if item["type"] == "lesson"
        ]
        reviews = [
            item for item in entries
            if item["type"] == "review"
        ]

        # Tuần có dòng ôn tập riêng không dùng để suy ra định mức.
        if not lessons or reviews:
            continue

        # Chỉ lấy tuần có đầy đủ số tiết của tất cả bài học.
        if any(item["periods"] is None for item in lessons):
            continue

        total = sum(item["periods"] for item in lessons)
        if total > 0:
            normal_week_totals[week] = total

    if not normal_week_totals:
        raise ValueError(
            "Không tìm được tuần học bình thường có số tiết rõ ràng "
            "để xác định định mức. Cần ít nhất một tuần có bài học "
            "và số tiết đầy đủ."
        )

    frequency = Counter(normal_week_totals.values())
    highest_frequency = max(frequency.values())
    most_common_totals = sorted(
        total for total, count in frequency.items()
        if count == highest_frequency
    )

    # Nếu nhiều mức đồng hạng, không lấy mức lớn nhất làm chuẩn.
    if len(most_common_totals) > 1:
        details = ", ".join(
            f"{total} tiết ({frequency[total]} tuần)"
            for total in most_common_totals
        )
        raise ValueError(
            "Không thể tự xác định định mức số tiết/tuần vì có nhiều "
            f"mức xuất hiện ngang nhau: {details}. "
            "Vui lòng kiểm tra lại bảng Word."
        )

    return most_common_totals[0]


# =====================================================
# PHÂN PHỐI BÀI HỌC/ÔN TẬP THEO SỐ TIẾT
# =====================================================

def distribute_lessons(weekly_data):
    """Phân phối tiết và báo lỗi cụ thể theo từng tuần, không hiện cảnh báo tự phân bổ."""
    standard = detect_period_standard(weekly_data)

    result = {}
    errors = []

    for week in sorted(weekly_data):
        entries = [item.copy() for item in weekly_data[week]]
        if not entries:
            continue

        missing = [item for item in entries if item["periods"] is None]
        known_total = sum(item["periods"] or 0 for item in entries)

        # Một nội dung chưa ghi số tiết có thể nhận phần tiết còn thiếu.
        # Việc này được thực hiện âm thầm, không tạo cảnh báo riêng.
        if len(missing) == 1:
            remaining = standard - known_total
            if remaining > 0:
                missing[0]["periods"] = remaining
            else:
                if known_total > standard:
                    errors.append(
                        f"Tuần {week}: nội dung '{missing[0]['title']}' chưa ghi số tiết "
                        f"và không còn tiết để phân bổ. Các nội dung đã ghi có {known_total} tiết, "
                        f"vượt định mức {standard} tiết."
                    )
                else:
                    errors.append(
                        f"Tuần {week}: nội dung '{missing[0]['title']}' chưa ghi số tiết; "
                        f"các nội dung còn lại đã đủ {standard}/{standard} tiết, "
                        "nên không còn tiết để phân bổ cho nội dung này."
                    )
                # Vẫn kiểm tra và báo các tiết thừa đã ghi (nếu có).
        elif len(missing) > 1:
            names = "; ".join(f"'{item['title']}'" for item in missing)
            errors.append(
                f"Tuần {week}: có {len(missing)} nội dung chưa ghi số tiết: {names}. "
                "Không thể xác định chính xác số tiết cho từng nội dung."
            )

        total = sum(item["periods"] or 0 for item in entries)

        if total > standard:
            excess = total - standard
            details = []
            running = 0
            for item in entries:
                count = item["periods"] or 0
                before = running
                running += count
                extra_here = max(0, running - standard) - max(0, before - standard)
                if extra_here > 0:
                    details.append(
                        f"'{item['title']}' (phần vượt {extra_here} tiết; nội dung ghi {count} tiết)"
                    )

            detail_text = "; ".join(details) if details else "không xác định được dòng gây thừa"
            errors.append(
                f"Tuần {week}: tổng cộng {total}/{standard} tiết, THỪA {excess} tiết. "
                f"Phần vượt nằm ở: {detail_text}."
            )
            continue

        if total < standard:
            deficit = standard - total
            if missing:
                missing_names = "; ".join(f"'{item['title']}'" for item in missing)
                reason = f"Nội dung chưa ghi số tiết: {missing_names}."
            else:
                breakdown = "; ".join(
                    f"'{item['title']}' ({item['periods']} tiết)"
                    for item in entries
                )
                reason = f"Các nội dung đã ghi: {breakdown}."

            errors.append(
                f"Tuần {week}: tổng cộng {total}/{standard} tiết, THIẾU {deficit} tiết. "
                f"{reason}"
            )
            continue

        # Chỉ đưa tuần hợp lệ vào kết quả xuất Excel.
        if any(item["periods"] is None for item in entries):
            continue

        expanded = []
        for item in entries:
            count = item["periods"] or 0
            if count > 0:
                expanded.extend([item["title"]] * count)

        result[week] = expanded

    # Giữ cấu trúc 4 giá trị để tương thích với phần giao diện hiện tại.
    return standard, result, [], errors


# =====================================================
# TẠO EXCEL CƠ BẢN NẾU KHÔNG CÓ FILE MẪU
# =====================================================

def create_basic_template():
    wb = Workbook()
    ws = wb.active
    ws.title = "Phân phối chương trình"

    ws.merge_cells("A1:D1")
    ws["A1"] = "PHÂN PHỐI CHƯƠNG TRÌNH"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws["A2"] = "Lớp"
    ws["A3"] = "Môn"
    ws["A4"] = "Phân môn"

    ws["A6"] = "Tuần PPCT"
    ws["B6"] = "Tiết"
    ws["C6"] = "Tên bài"
    ws["D6"] = "Ghi chú"

    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 65
    ws.column_dimensions["D"].width = 25

    for cell in ws[6]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    return wb


# =====================================================
# TÌM HÀNG TIÊU ĐỀ TRONG EXCEL MẪU
# =====================================================

def find_excel_header(ws):
    for row_index in range(1, min(ws.max_row, 60) + 1):
        labels = [
            normalize(ws.cell(row_index, col).value)
            for col in range(1, ws.max_column + 1)
        ]

        week_col = next(
            (
                i + 1 for i, value in enumerate(labels)
                if "tuan" in value
            ),
            None
        )
        period_col = next(
            (
                i + 1 for i, value in enumerate(labels)
                if value == "tiet"
            ),
            None
        )
        title_col = next(
            (
                i + 1 for i, value in enumerate(labels)
                if "ten bai" in value or "noi dung" in value
            ),
            None
        )
        note_col = next(
            (
                i + 1 for i, value in enumerate(labels)
                if "ghi chu" in value
            ),
            None
        )

        if week_col and period_col and title_col:
            return row_index, week_col, period_col, title_col, note_col

    raise ValueError(
        "Excel mẫu cần có các cột Tuần PPCT, Tiết và Tên bài."
    )


# =====================================================
# ĐIỀN LỚP, MÔN, PHÂN MÔN VÀO EXCEL
# =====================================================

def fill_excel_metadata(ws, metadata):
    labels = {
        "lop": "lop",
        "mon": "mon",
        "phan_mon": "phan mon"
    }

    filled = []

    for row in range(1, min(ws.max_row, 60) + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row, col)

            if isinstance(cell, MergedCell):
                continue

            label = normalize(cell.value).rstrip(":：- ").strip()

            key = next(
                (name for name, normalized_label in labels.items()
                 if label == normalized_label),
                None
            )
            if not key:
                continue

            value = metadata.get(key, "")
            if not value:
                continue

            target_col = col + 1

            while target_col <= ws.max_column:
                target = ws.cell(row, target_col)

                if isinstance(target, MergedCell):
                    target_col += 1
                    continue

                target.value = value

                alignment = copy(target.alignment)
                alignment.horizontal = "center"
                alignment.vertical = "center"
                target.alignment = alignment

                if key not in filled:
                    filled.append(key)
                break

    return filled


# =====================================================
# SAO CHÉP ĐỊNH DẠNG DÒNG EXCEL
# =====================================================

def copy_row_format(ws, source_row, target_row):
    if source_row == target_row:
        return

    for col in range(1, ws.max_column + 1):
        source = ws.cell(source_row, col)
        target = ws.cell(target_row, col)

        if source.has_style:
            target._style = copy(source._style)

        if source.number_format:
            target.number_format = source.number_format

        if source.alignment:
            target.alignment = copy(source.alignment)

        if source.protection:
            target.protection = copy(source.protection)

        if source.font:
            target.font = copy(source.font)

        if source.fill:
            target.fill = copy(source.fill)

        if source.border:
            target.border = copy(source.border)

    source_height = ws.row_dimensions[source_row].height
    if source_height:
        ws.row_dimensions[target_row].height = source_height


# =====================================================
# XUẤT DỮ LIỆU RA EXCEL
# =====================================================

def export_excel(template_bytes, weekly_lessons, metadata):
    if template_bytes:
        wb = load_workbook(BytesIO(template_bytes))
        ws = wb.active
    else:
        wb = create_basic_template()
        ws = wb.active

    fill_excel_metadata(ws, metadata)

    header_row, week_col, period_col, title_col, note_col = (
        find_excel_header(ws)
    )

    start_row = header_row + 1
    style_row = start_row

    data = []
    global_period = 1

    for week in sorted(weekly_lessons):
        for lesson in weekly_lessons[week]:
            data.append({
                "week": week,
                "period": global_period,
                "title": lesson
            })
            global_period += 1

    last_row = start_row + len(data) - 1
    old_last_row = max(ws.max_row, last_row)

    # Nhân định dạng dòng mẫu cho các dòng mới.
    for row in range(start_row, last_row + 1):
        if row > ws.max_row:
            copy_row_format(ws, style_row, row)

    # Xóa dữ liệu cũ trong toàn bộ vùng dữ liệu, tránh sót dòng
    # khi lần xuất mới ít tiết hơn lần xuất trước.
    for row in range(start_row, old_last_row + 1):
        for col in (week_col, period_col, title_col):
            cell = ws.cell(row, col)
            if not isinstance(cell, MergedCell):
                cell.value = None

        if note_col:
            cell = ws.cell(row, note_col)
            if not isinstance(cell, MergedCell):
                cell.value = None

    # Ghi dữ liệu.
    for index, item in enumerate(data):
        row = start_row + index

        ws.cell(row, week_col).value = item["week"]
        ws.cell(row, period_col).value = item["period"]
        ws.cell(row, title_col).value = item["title"]

    output = BytesIO()
    wb.save(output)
    return output.getvalue(), len(data)


# =====================================================
# GIAO DIỆN STREAMLIT
# =====================================================

st.title("📚 PPCT WORD → EXCEL")


word_file = st.file_uploader(
    "📄 Tải file Word chương trình",
    type=["docx"]
)

excel_file = st.file_uploader(
    "📊 Tải file Excel mẫu (không bắt buộc)",
    type=["xlsx"]
)

st.info(
    "Nếu không tải Excel mẫu, chương trình sẽ tạo file Excel mới"
)

if st.button(
    "🔄 CHUYỂN WORD SANG EXCEL",
    type="primary",
    use_container_width=True
):
    # Xóa kết quả cũ để tránh tải nhầm file cũ.
    st.session_state.pop("excel_output", None)
    st.session_state.pop("excel_count", None)

    if not word_file:
        st.error("Vui lòng tải file Word (.docx).")
    else:
        try:
            word_bytes = word_file.getvalue()

            with st.spinner("Đang đọc Word và kiểm tra số tiết..."):
                metadata = extract_word_metadata(word_bytes)
                st.session_state["metadata"] = metadata
                weekly_data = read_word(word_bytes)
                standard, weekly_lessons, warnings, errors = (
                    distribute_lessons(weekly_data)
                )

            st.subheader("Thông tin nhận diện từ Word")
            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Lớp",
                metadata.get("lop") or "Chưa nhận diện"
            )
            col2.metric(
                "Môn",
                metadata.get("mon") or "Chưa nhận diện"
            )
            col3.metric(
                "Số tuần",
                len(weekly_data)
            )

            st.subheader("Kết quả kiểm tra")
            st.write(
                f"Định mức phổ biến tự nhận diện: **{standard} tiết/tuần**"
            )

            with st.expander(
                "Xem dữ liệu đã đọc và phân phối",
                expanded=False
            ):
                for week in sorted(weekly_data):
                    st.markdown(f"**Tuần {week}**")

                    for item in weekly_data[week]:
                        periods = (
                            str(item["periods"])
                            if item["periods"] is not None
                            else "Tự lấy phần tiết còn thiếu/định mức"
                        )
                        st.write(
                            f"- {item['title']} — {periods} "
                            f"({item['type']})"
                        )

                    distributed = weekly_lessons.get(week, [])
                    st.caption(
                        f"Kết quả phân phối: {len(distributed)} tiết"
                    )

            if errors:
                st.error(
                    "Chưa thể xuất Excel vì có tuần thừa hoặc thiếu số tiết so với định mức."
                )
                for error in errors:
                    st.write("❌ " + error)
                st.stop()

            # Chọn Excel người dùng tải lên, nếu không thì dùng mẫu cùng thư mục.
            if excel_file:
                template_bytes = excel_file.getvalue()
            elif DEFAULT_TEMPLATE.exists():
                template_bytes = DEFAULT_TEMPLATE.read_bytes()
            else:
                template_bytes = None

            output, total_periods = export_excel(
                template_bytes,
                weekly_lessons,
                metadata
            )

            st.session_state["excel_output"] = output
            st.session_state["excel_count"] = total_periods

            st.success(
                f"Đã xử lý thành công {total_periods} tiết."
            )

        except Exception as error:
            st.error(f"Không thể xử lý file: {error}")


# =====================================================
# NÚT TẢI EXCEL
# =====================================================


if st.session_state.get("excel_output"):

    metadata = st.session_state.get("metadata", {})

    lop = metadata.get("lop", "")
    mon = metadata.get("mon", "")

    # Tạo tên file tự động.
    file_name = (
        f"PPCT_mon_{mon}lop_{lop}.xlsx"
    )

    st.download_button(
        label="⬇️ TẢI FILE EXCEL HOÀN THÀNH",
        data=st.session_state["excel_output"],
        file_name=file_name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    st.success(
        "File Excel đã được tạo. "
        "Hãy mở kiểm tra trước khi sử dụng chính thức."
    )

st.divider()
st.caption(
    "Hỗ trợ Word .docx và Excel .xlsx. "
    "File .xls cũ cần lưu lại thành .xlsx trước khi tải lên."
)
