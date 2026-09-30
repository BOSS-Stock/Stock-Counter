From __future__ import annotations

import io
import re
from typing import Any

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(
    page_title="Stock Diff | ระบบนับสต็อก",
    layout="wide",
    initial_sidebar_state="expanded",
)


SAMPLE_ROWS = [
    {"Product": "น้ำดื่ม 600 ml", "System qty": 48, "Actual qty": None},
    {"Product": "กาแฟกระป๋อง สูตรดั้งเดิม", "System qty": 36, "Actual qty": None},
    {"Product": "ขนมปังโฮลวีต", "System qty": 24, "Actual qty": None},
    {"Product": "นมสดพาสเจอร์ไรส์ 2 ลิตร", "System qty": 18, "Actual qty": None},
    {"Product": "น้ำผลไม้รวม", "System qty": 30, "Actual qty": None},
]

SYSTEM_ADD_COLUMNS = (6, 7, 8, 9, 10)  # Excel columns 7, 8, 9, 10, 11
SYSTEM_SUBTRACT_COLUMNS = (11, 13)  # Excel columns 12 and 14


def _thai_digits_to_arabic(value: str) -> str:
    translation = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")
    return value.translate(translation)


def _number_from_text(value: str) -> int | float | None:
    value = _thai_digits_to_arabic(value).replace(",", "").strip()
    match = re.fullmatch(r"-?\d+(?:\.\d+)?", value)
    if not match:
        return None
    parsed = float(value)
    return int(parsed) if parsed.is_integer() else parsed


def _cost_from_text(value: Any) -> int | float | None:
    if value is None or pd.isna(value):
        return None
    cleaned = _thai_digits_to_arabic(str(value)).strip()
    cleaned = re.sub(r"(?:฿|บาท|baht|thb)", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace("$", "").replace("€", "").replace("£", "")
    cleaned = cleaned.replace(" ", "").replace(",", "")
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = f"-{cleaned[1:-1]}"
    return _number_from_text(cleaned)


def _normalize_column_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9ก-๙]+", "", str(value).casefold())


def _find_column(columns: list[Any], candidates: tuple[str, ...]) -> Any | None:
    normalized_columns = {
        column: _normalize_column_name(column) for column in columns
    }
    normalized_candidates = [_normalize_column_name(candidate) for candidate in candidates]
    for candidate in normalized_candidates:
        for column, normalized in normalized_columns.items():
            if normalized == candidate:
                return column
    for candidate in normalized_candidates:
        for column, normalized in normalized_columns.items():
            if candidate and candidate in normalized:
                return column
    return None


def _clean_product(value: Any) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _combine_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    # Combine repeated product lines, which is common when a workbook has
    # several store or category sheets.
    combined: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = re.sub(r"\s+", " ", row["Product"]).casefold()
        if key in combined:
            combined[key]["System qty"] += row["System qty"]
        else:
            combined[key] = row
    return list(combined.values())


def _product_key(value: Any) -> str:
    return re.sub(r"\s+", " ", _clean_product(value)).casefold()


def parse_excel(file_bytes: bytes) -> list[dict[str, Any]]:
    workbook = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    rows: list[dict[str, Any]] = []
    product_candidates = (
        "ชื่อสินค้า",
        "สินค้า",
        "รายการสินค้า",
        "product name",
        "product",
        "item name",
        "item",
        "รายการ",
    )
    eligible_sheet_found = False

    for sheet_name, sheet in workbook.items():
        if sheet.empty:
            continue
        if len(sheet.columns) <= max(*SYSTEM_ADD_COLUMNS, *SYSTEM_SUBTRACT_COLUMNS):
            continue
        eligible_sheet_found = True
        columns = list(sheet.columns)
        product_column = _find_column(columns, product_candidates)

        # If the export uses custom headers, choose the first text-like column
        # as a practical fallback.
        if product_column is None:
            for column in columns:
                if sheet[column].map(_clean_product).str.len().gt(0).any():
                    product_column = column
                    break
        if product_column is None:
            continue

        for row_index, product_value in enumerate(sheet[product_column]):
            product = _clean_product(product_value)
            if not product or product.casefold() in {"nan", "none"}:
                continue
            additions = sum(
                _number_from_text(str(sheet.iloc[row_index, column_index])) or 0
                for column_index in SYSTEM_ADD_COLUMNS
            )
            deductions = sum(
                _number_from_text(str(sheet.iloc[row_index, column_index])) or 0
                for column_index in SYSTEM_SUBTRACT_COLUMNS
            )
            rows.append(
                {
                    "Product": product,
                    "System qty": additions - deductions,
                    "Actual qty": None,
                }
            )

    if not eligible_sheet_found:
        raise ValueError(
            "ไฟล์ Excel ต้องมีอย่างน้อย 14 คอลัมน์ เพื่อคำนวณยอดในระบบจากคอลัมน์ 7+8+9+10+11-12-14"
        )
    # Keep both positive and negative balances, but remove products whose
    # calculated system balance is exactly zero.
    return [row for row in _combine_rows(rows) if row["System qty"] != 0]


PDF_PRODUCT_CANDIDATES = (
    "ชื่อสินค้า",
    "สินค้า",
    "รายการสินค้า",
    "product name",
    "product",
    "item name",
    "item",
    "รายการ",
)

PDF_COST_CANDIDATES = (
    "std cost (excl tax)",
    "std cost",
    "standard cost",
    "ต้นทุนต่อหน่วย",
    "ต้นทุนสินค้า",
    "ราคาทุนต่อหน่วย",
    "ราคาทุน",
    "ต้นทุน",
    "unit cost",
    "cost price",
    "cost",
)

PDF_STD_COST_CANDIDATES = (
    "std cost (excl tax)",
    "std cost",
    "standard cost",
)


def _find_value_position(values: list[Any], candidates: tuple[str, ...]) -> int | None:
    normalized_values = [_normalize_column_name(value) for value in values]
    normalized_candidates = [_normalize_column_name(candidate) for candidate in candidates]
    for candidate in normalized_candidates:
        for index, normalized in enumerate(normalized_values):
            if normalized == candidate:
                return index
    for candidate in normalized_candidates:
        for index, normalized in enumerate(normalized_values):
            if candidate and candidate in normalized:
                return index
    return None


def _is_pdf_summary_or_header(value: Any) -> bool:
    normalized = _normalize_column_name(value)
    return normalized in {
        "รวม",
        "รวมทั้งสิ้น",
        "ยอดรวม",
        "total",
        "totals",
        "subtotal",
        "grandtotal",
    } or normalized.startswith("รวม")


def _infer_pdf_columns(rows: list[list[Any]]) -> tuple[int, int] | None:
    if not rows:
        return None
    column_count = max(len(row) for row in rows)
    numeric_scores: list[tuple[float, int]] = []
    text_scores: list[tuple[int, int]] = []

    for column_index in range(column_count):
        values = [
            row[column_index]
            for row in rows
            if column_index < len(row) and _clean_product(row[column_index])
        ]
        if not values:
            continue
        numeric_count = sum(_cost_from_text(value) is not None for value in values)
        non_numeric_count = sum(
            _cost_from_text(value) is None and not _is_pdf_summary_or_header(value)
            for value in values
        )
        numeric_scores.append((numeric_count / len(values), column_index))
        text_scores.append((non_numeric_count, column_index))

    if not numeric_scores or not text_scores:
        return None
    cost_score, cost_index = max(numeric_scores)
    _, product_index = max(text_scores)
    if cost_score <= 0 or cost_index == product_index:
        return None
    return product_index, cost_index


def _parse_pdf_stocktake_table(table: list[list[Any]]) -> list[dict[str, Any]]:
    rows = [
        ["" if cell is None else str(cell).strip() for cell in row]
        for row in table
        if row and any(cell is not None and str(cell).strip() for cell in row)
    ]
    if not rows:
        return []

    std_cost_index: int | None = None
    for row in rows[:8]:
        candidate_index = _find_value_position(row, PDF_STD_COST_CANDIDATES)
        if candidate_index is not None:
            std_cost_index = candidate_index
            break

    parsed: list[dict[str, Any]] = []
    for row in rows:
        if not row or not re.match(r"^\s*\d{8,}\b", row[0]):
            continue
        product = re.sub(r"^\s*\d{8,}\s*", "", row[0])
        product = re.sub(r"\s+", " ", product).strip()
        if not product:
            continue

        cost: int | float | None = None
        if std_cost_index is not None and std_cost_index < len(row):
            cost = _cost_from_text(row[std_cost_index])
        if cost is None:
            for cell in reversed(row[1:]):
                cost = _cost_from_text(cell)
                if cost is not None:
                    break
        if cost is not None:
            parsed.append({"Product": product, "Cost": cost})
    return parsed


def _parse_pdf_table(table: list[list[Any]]) -> list[dict[str, Any]]:
    stocktake_rows = _parse_pdf_stocktake_table(table)
    if stocktake_rows:
        return stocktake_rows

    rows = [
        ["" if cell is None else str(cell).strip() for cell in row]
        for row in table
        if row and any(cell is not None and str(cell).strip() for cell in row)
    ]
    if not rows:
        return []

    product_index: int | None = None
    cost_index: int | None = None
    header_index: int | None = None
    for row_index, row in enumerate(rows[:8]):
        candidate_product_index = _find_value_position(row, PDF_PRODUCT_CANDIDATES)
        candidate_cost_index = _find_value_position(row, PDF_COST_CANDIDATES)
        if (
            candidate_product_index is not None
            and candidate_cost_index is not None
            and candidate_product_index != candidate_cost_index
        ):
            product_index = candidate_product_index
            cost_index = candidate_cost_index
            header_index = row_index
            break

    if product_index is None or cost_index is None:
        inferred_columns = _infer_pdf_columns(rows)
        if inferred_columns is None:
            return []
        product_index, cost_index = inferred_columns

    data_rows = rows[header_index + 1 :] if header_index is not None else rows
    parsed: list[dict[str, Any]] = []
    for row in data_rows:
        if max(product_index, cost_index) >= len(row):
            continue
        product = _clean_product(row[product_index])
        cost = _cost_from_text(row[cost_index])
        if not product or _is_pdf_summary_or_header(product) or cost is None:
            continue
        parsed.append({"Product": product, "Cost": cost})
    return parsed


def _parse_pdf_text_lines(text: str) -> list[dict[str, Any]]:
    parsed: list[dict[str, Any]] = []
    cost_pattern = re.compile(
        r"(?:\s{2,}|[|:])\s*"
        r"(?P<cost>(?:฿\s*)?\(?-?\s*[๐-๙0-9][๐-๙0-9,]*(?:\.\d+)?\)?)"
        r"\s*(?:บาท|฿|THB)?\s*$",
        flags=re.IGNORECASE,
    )
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        match = cost_pattern.search(line)
        if not match:
            continue
        product = re.sub(r"\s+", " ", line[: match.start()]).strip(" |:-")
        product = re.sub(r"^\d+\s*[.)-]\s*", "", product)
        cost = _cost_from_text(match.group("cost"))
        if (
            product
            and not _is_pdf_summary_or_header(product)
            and cost is not None
            and _find_value_position([product], PDF_PRODUCT_CANDIDATES) is None
        ):
            parsed.append({"Product": product, "Cost": cost})
    return parsed


def _combine_cost_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    combined: dict[str, dict[str, Any]] = {}
    for row in rows:
        product = _clean_product(row.get("Product"))
        cost = _cost_from_text(row.get("Cost"))
        key = _product_key(product)
        if not key or cost is None:
            continue
        if key not in combined:
            combined[key] = {"Product": product, "Cost": cost}
    return list(combined.values())


def parse_pdf_costs(file_bytes: bytes) -> list[dict[str, Any]]:
    parsed_rows: list[dict[str, Any]] = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            page_rows: list[dict[str, Any]] = []
            for table in page.extract_tables() or []:
                page_rows.extend(_parse_pdf_table(table))
            if page_rows:
                parsed_rows.extend(page_rows)
            else:
                page_text = page.extract_text() or ""
                parsed_rows.extend(_parse_pdf_text_lines(page_text))

    combined_rows = _combine_cost_rows(parsed_rows)
    if not combined_rows:
        raise ValueError(
            "ไม่พบตารางต้นทุนใน PDF โปรดตรวจว่ามีคอลัมน์ชื่อสินค้าและ Cost/ต้นทุน "
            "และไฟล์ไม่ใช่ PDF แบบภาพสแกน"
        )
    return combined_rows


def make_result_table(edited: pd.DataFrame) -> pd.DataFrame:
    result = edited.copy()
    result = result.drop(columns=["ลำดับ"], errors="ignore")
    result.insert(0, "ลำดับ", range(1, len(result) + 1))
    result["System qty"] = pd.to_numeric(result["System qty"], errors="coerce").fillna(0)
    result["Actual qty"] = pd.to_numeric(result["Actual qty"], errors="coerce")
    result["Diff"] = result["Actual qty"].fillna(0) - result["System qty"]
    result["Status"] = result["Diff"].map(
        lambda diff: "ตรงกัน" if diff == 0 else ("เกิน" if diff > 0 else "ขาด")
    )
    return result[["ลำดับ", "Product", "System qty", "Actual qty", "Diff", "Status"]]


def make_cost_result_table(
    result: pd.DataFrame, cost_rows: list[dict[str, Any]]
) -> pd.DataFrame:
    if not cost_rows:
        return pd.DataFrame(
            columns=[
                "ลำดับ",
                "Product",
                "System qty",
                "Actual qty",
                "Diff",
                "Status",
                "Cost",
                "Shortage cost",
                "Overage cost",
            ]
        )

    costs = pd.DataFrame(cost_rows)
    costs["_Product key"] = costs["Product"].map(_product_key)
    matched = result[result["Actual qty"].notna()].copy()
    matched["_Product key"] = matched["Product"].map(_product_key)
    matched = matched.merge(costs[["_Product key", "Cost"]], on="_Product key", how="inner")
    if matched.empty:
        return pd.DataFrame(
            columns=[
                "ลำดับ",
                "Product",
                "System qty",
                "Actual qty",
                "Diff",
                "Status",
                "Cost",
                "Shortage cost",
                "Overage cost",
            ]
        )

    matched["Shortage cost"] = matched["Diff"].clip(upper=0) * matched["Cost"]
    matched["Overage cost"] = matched["Diff"].clip(lower=0) * matched["Cost"]
    return matched[
        [
            "ลำดับ",
            "Product",
            "System qty",
            "Actual qty",
            "Diff",
            "Status",
            "Cost",
            "Shortage cost",
            "Overage cost",
        ]
    ]


def status_cell_style(status: Any) -> str:
    colors = {
        "ตรงกัน": ("#DCFCE7", "#166534"),
        "ขาด": ("#FEE2E2", "#991B1B"),
        "เกิน": ("#FEF3C7", "#92400E"),
    }
    background, foreground = colors.get(status, ("", ""))
    if not background:
        return ""
    return (
        f"background-color: {background}; "
        f"color: {foreground}; "
        "font-weight: 700; "
        "text-align: center;"
    )


def load_rows(rows: list[dict[str, Any]], source_label: str) -> None:
    st.session_state["inventory_rows"] = rows
    st.session_state["source_label"] = source_label


if "inventory_rows" not in st.session_state:
    load_rows(SAMPLE_ROWS, "ตัวอย่างสำหรับทดลองใช้งาน")
if "cost_rows" not in st.session_state:
    st.session_state["cost_rows"] = []
if "cost_source_label" not in st.session_state:
    st.session_state["cost_source_label"] = ""


with st.sidebar:
    st.header("นำเข้าข้อมูล")
    uploaded_file = st.file_uploader(
        "อัปโหลดไฟล์ Excel สต็อก",
        type=["xlsx", "xls"],
        help="รองรับไฟล์ Excel ที่มีคอลัมน์ชื่อสินค้าและจำนวนคงเหลือในระบบ",
    )
    if uploaded_file is not None:
        file_key = f"{uploaded_file.name}:{uploaded_file.size}"
        if st.session_state.get("uploaded_file_key") != file_key:
            with st.spinner("กำลังอ่านรายการสินค้าจาก Excel..."):
                try:
                    parsed_rows = parse_excel(uploaded_file.getvalue())
                except Exception as error:
                    st.error(f"อ่านไฟล์ไม่สำเร็จ: {error}")
                    parsed_rows = []
            st.session_state["uploaded_file_key"] = file_key
            if parsed_rows:
                load_rows(parsed_rows, uploaded_file.name)
                st.success(f"พบ {len(parsed_rows)} รายการ")
            else:
                st.warning(
                    "ไม่พบรายการสินค้าอัตโนมัติ ลองตรวจว่าไฟล์มีคอลัมน์ชื่อสินค้าและจำนวนคงเหลือในระบบ"
                )

    st.divider()
    st.subheader("ไฟล์ต้นทุนสินค้า")
    cost_pdf = st.file_uploader(
        "อัปโหลดไฟล์ที่ 2 (PDF เท่านั้น)",
        type=["pdf"],
        help="PDF ควรมีคอลัมน์ชื่อสินค้าและ Cost/ต้นทุนต่อหน่วย",
    )
    if cost_pdf is not None:
        cost_file_key = f"{cost_pdf.name}:{cost_pdf.size}"
        if st.session_state.get("cost_pdf_file_key") != cost_file_key:
            with st.spinner("กำลังอ่าน Cost จากไฟล์ PDF..."):
                try:
                    parsed_cost_rows = parse_pdf_costs(cost_pdf.getvalue())
                except Exception as error:
                    st.session_state["cost_rows"] = []
                    st.session_state["cost_source_label"] = ""
                    st.error(f"อ่านไฟล์ต้นทุนไม่สำเร็จ: {error}")
                    parsed_cost_rows = []
            st.session_state["cost_pdf_file_key"] = cost_file_key
            if parsed_cost_rows:
                st.session_state["cost_rows"] = parsed_cost_rows
                st.session_state["cost_source_label"] = cost_pdf.name
                st.success(f"พบต้นทุน {len(parsed_cost_rows)} รายการ")
            elif not st.session_state.get("cost_source_label"):
                st.warning("ไม่พบรายการต้นทุนจากไฟล์ PDF")

    if st.button("เริ่มจากตัวอย่าง", width="stretch"):
        st.session_state.pop("uploaded_file_key", None)
        load_rows(SAMPLE_ROWS, "ตัวอย่างสำหรับทดลองใช้งาน")
        st.rerun()

    if st.session_state.get("cost_source_label"):
        st.caption(f"ไฟล์ต้นทุน: {st.session_state['cost_source_label']}")


st.title("ระบบนับสต็อก Con BLP")
st.write(
    "อัปโหลดรายการจาก Excel กรอกยอดที่นับได้จริง แล้วดูส่วนต่างระหว่างหน้าร้านกับระบบทันที"
)
st.caption(f"แหล่งข้อมูล: {st.session_state.get('source_label', 'ยังไม่ได้เลือกไฟล์')}")

rows = st.session_state["inventory_rows"]
if not rows:
    st.info("ยังไม่มีรายการสินค้า กรุณาอัปโหลดไฟล์ Excel ที่แถบด้านซ้าย")
    st.stop()

input_df = pd.DataFrame(rows)
input_df.insert(0, "ลำดับ", range(1, len(input_df) + 1))
input_df["Actual qty"] = input_df["Actual qty"].astype(object).where(
    input_df["Actual qty"].notna(), ""
)

st.subheader("กรอกยอดนับจริง")
st.caption("แก้ไขเฉพาะคอลัมน์ “นับจริงหน้าร้าน” จากนั้นดูผลสรุปด้านล่าง")
edited_df = st.data_editor(
    input_df,
    hide_index=True,
    width="stretch",
    num_rows="dynamic",
    c
