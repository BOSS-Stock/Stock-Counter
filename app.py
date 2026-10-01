from __future__ import annotations

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
    {"Category": "Beverage", "Product": "น้ำดื่ม 600 ml", "System qty": 48, "Actual qty": None},
    {"Category": "Beverage", "Product": "กาแฟกระป๋อง สูตรดั้งเดิม", "System qty": 36, "Actual qty": None},
    {"Category": "Candy", "Product": "ขนมปังโฮลวีต", "System qty": 24, "Actual qty": None},
    {"Category": "Candy", "Product": "นมสดพาสเจอร์ไรส์ 2 ลิตร", "System qty": 18, "Actual qty": None},
    {"Category": "Candy", "Product": "น้ำผลไม้รวม", "System qty": 30, "Actual qty": None},
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
    combined: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = (row.get("Category", ""), re.sub(r"\s+", " ", row["Product"]).casefold())
        if key in combined:
            combined[key]["System qty"] += row["System qty"]
        else:
            combined[key] = row.copy()
    return list(combined.values())


def _product_key(value: Any) -> str:
    return re.sub(r"\s+", " ", _clean_product(value)).casefold()


def parse_excel(file_bytes: bytes, fallback_category_name: str = "ทั่วไป") -> list[dict[str, Any]]:
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
            
            # ดึงหมวดสินค้าจากคอลัมน์ A (Index 0) หากมีข้อมูล ถ้าไม่มีให้ใช้ชื่อไฟล์
            col_a_val = sheet.iloc[row_index, 0]
            col_a_clean = _clean_product(col_a_val)
            if col_a_clean and col_a_clean.casefold() not in {"nan", "none", "ลำดับ", "no", "no."}:
                category_name = col_a_clean
            else:
                category_name = fallback_category_name

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
                    "Category": category_name,
                    "Product": product,
                    "System qty": additions - deductions,
                    "Actual qty": None,
                }
            )

    if not eligible_sheet_found:
        raise ValueError(
            f"ไฟล์ต้องมีอย่างน้อย 14 คอลัมน์ เพื่อคำนวณยอดในระบบ"
        )
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


def prepare_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """จัดระเบียบลำดับและแทรกแถบหัวข้อคั่นระหว่างหมวดสินค้า"""
    display_rows = []
    grouped = df.groupby("Category", sort=False)
    
    for cat_name, group in grouped:
        # แถบแบ่งหมวดสินค้า
        display_rows.append({
            "ลำดับ": None,
            "Category": f"📌 [{cat_name}]",
            "Product": f"--- หมวด: {cat_name} ---",
            "System qty": None,
            "Actual qty": None,
            "is_header": True
        })
        
        # รันลำดับ 1, 2, 3 ใหม่ในหมวดนั้นๆ
        for idx, (_, row) in enumerate(group.iterrows(), start=1):
            row_dict = row.to_dict()
            row_dict["ลำดับ"] = idx
            row_dict["is_header"] = False
            display_rows.append(row_dict)
            
    res_df = pd.DataFrame(display_rows)
    return res_df


def make_result_table(edited: pd.DataFrame) -> pd.DataFrame:
    # กรองเอาแถบหัวข้อคั่นหมวดหมู่ออกเพื่อการคำนวณที่ถูกต้อง
    result = edited[edited.get("is_header", False) == False].copy()
    result = result.drop(columns=["is_header"], errors="ignore")
    
    result["System qty"] = pd.to_numeric(result["System qty"], errors="coerce").fillna(0)
    result["Actual qty"] = pd.to_numeric(result["Actual qty"], errors="coerce")
    result["Diff"] = result["Actual qty"].fillna(0) - result["System qty"]
    result["Status"] = result["Diff"].map(
        lambda diff: "ตรงกัน" if diff == 0 else ("เกิน" if diff > 0 else "ขาด")
    )
    cols = ["ลำดับ", "Category", "Product", "System qty", "Actual qty", "Diff", "Status"]
    return result[[c for c in cols if c in result.columns]]


def make_cost_result_table(
    result: pd.DataFrame, cost_rows: list[dict[str, Any]]
) -> pd.DataFrame:
    if not cost_rows:
        return pd.DataFrame(
            columns=[
                "ลำดับ",
                "Category",
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
                "Category",
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
    cols = [
        "ลำดับ",
        "Category",
        "Product",
        "System qty",
        "Actual qty",
        "Diff",
        "Status",
        "Cost",
        "Shortage cost",
        "Overage cost",
    ]
    return matched[[c for c in cols if c in matched.columns]]


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
    uploaded_files = st.file_uploader(
        "อัปโหลดไฟล์ Excel สต็อก (สูงสุด 11 ไฟล์)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        help="รองรับการอัปโหลดได้สูงสุด 11 ไฟล์พร้อมกัน โดยระบบจะดึงหมวดสินค้าจากคอลัมน์ A",
    )
    if uploaded_files:
        if len(uploaded_files) > 11:
            st.error("สามารถอัปโหลดได้สูงสุด 11 ไฟล์เท่านั้นครับ")
            uploaded_files = uploaded_files[:11]

        files_key = "-".join([f"{f.name}:{f.size}" for f in uploaded_files])
        if st.session_state.get("uploaded_files_key") != files_key:
            all_parsed_rows: list[dict[str, Any]] = []
            success_files = []
            with st.spinner("กำลังอ่านรายการสินค้าจาก Excel..."):
                for file in uploaded_files:
                    fallback_cat = file.name.rsplit(".", 1)[0]
                    try:
                        parsed = parse_excel(file.getvalue(), fallback_category_name=fallback_cat)
                        all_parsed_rows.extend(parsed)
                        success_files.append(fallback_cat)
                    except Exception as error:
                        st.error(f"อ่านไฟล์ {file.name} ไม่สำเร็จ: {error}")

            st.session_state["uploaded_files_key"] = files_key
            if all_parsed_rows:
                load_rows(all_parsed_rows, f"อัปโหลด {len(success_files)} ไฟล์เรียบร้อย")
                st.success(f"นำเข้าข้อมูลเรียบร้อย {len(all_parsed_rows)} รายการ")
            else:
                st.warning("ไม่พบรายการสินค้าจากไฟล์ที่อัปโหลด")

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

    if st.button("เริ่มจากตัวอย่าง", use_container_width=True):
        st.session_state.pop("uploaded_files_key", None)
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

raw_input_df = pd.DataFrame(rows)
if "Category" not in raw_input_df.columns:
    raw_input_df["Category"] = "ทั่วไป"

raw_input_df["Actual qty"] = raw_input_df["Actual qty"].astype(object).where(
    raw_input_df["Actual qty"].notna(), ""
)

# จัดโครงสร้างตารางภาพรวม รันลำดับ 1 ใหม่ทุกหมวด และเพิ่มเส้นคั่นหมวด
input_df = prepare_display_dataframe(raw_input_df)

st.subheader("กรอกยอดนับจริง")
st.caption("แก้ไขเฉพาะคอลัมน์ “นับจริงหน้าร้าน” จากนั้นดูผลสรุปด้านล่าง")

# Tab navigation for categories
categories = list(raw_input_df["Category"].unique())
tab_titles = ["ภาพรวม (รวมทุกหมวด)"] + [f"📦 {cat}" for cat in categories]
tabs = st.tabs(tab_titles)

with tabs[0]:
    edited_df = st.data_editor(
        input_df,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "ลำดับ": st.column_config.NumberColumn("ลำดับ", format="%.0f", disabled=True),
            "Category": st.column_config.TextColumn("หมวดสินค้า", disabled=True),
            "Product": st.column_config.TextColumn("สินค้า", disabled=True),
            "System qty": st.column_config.NumberColumn("ในระบบ", disabled=True),
            "Actual qty": st.column_config.TextColumn(
                "นับจริงหน้าร้าน",
                validate=r"^\s*\d*(?:\.\d+)?\s*$",
                help="ใส่จำนวนที่นับได้จริงเป็นตัวเลข",
            ),
        },
        key="inventory_editor_all",
    )

for i, cat in enumerate(categories):
    with tabs[i + 1]:
        cat_df = raw_input_df[raw_input_df["Category"] == cat].copy()
        cat_df.insert(0, "ลำดับ", range(1, len(cat_df) + 1))
        st.dataframe(
            cat_df,
            hide_index=True,
            use_container_width=True,
            column_config={
                "ลำดับ": st.column_config.NumberColumn("ลำดับ", format="%.0f"),
                "Category": st.column_config.TextColumn("หมวดสินค้า"),
                "Product": st.column_config.TextColumn("สินค้า"),
                "System qty": st.column_config.NumberColumn("ในระบบ"),
                "Actual qty": st.column_config.TextColumn("นับจริงหน้าร้าน"),
            },
        )

result_df = make_result_table(edited_df)
actual_entered = result_df["Actual qty"].notna()
diff_series = result_df.loc[actual_entered, "Diff"]
short_count = int((diff_series < 0).sum())
over_count = int((diff_series > 0).sum())
match_count = int((diff_series == 0).sum())

st.subheader("ภาพรวม")
metric_columns = st.columns(4)
metric_columns[0].metric("สินค้าทั้งหมด", len(result_df))
metric_columns[1].metric("กรอกยอดแล้ว", int(actual_entered.sum()))
metric_columns[2].metric("ขาด", short_count)
metric_columns[3].metric("เกิน", over_count)

if actual_entered.any():
    st.subheader("ผลต่างสต็อก")
    filter_choice = st.selectbox(
        "แสดงรายการ",
        ["ทั้งหมด", "ขาด", "เกิน", "ตรงกัน", "ยังไม่ได้นับ"],
        label_visibility="collapsed",
    )
    filtered = result_df.copy()
    if filter_choice == "ขาด":
        filtered = filtered[filtered["Diff"] < 0]
    elif filter_choice == "เกิน":
        filtered = filtered[filtered["Diff"] > 0]
    elif filter_choice == "ตรงกัน":
        filtered = filtered[filtered["Diff"] == 0]
    elif filter_choice == "ยังไม่ได้นับ":
        filtered = filtered[~actual_entered]

    st.dataframe(
        filtered.style.map(status_cell_style, subset=["Status"]),
        hide_index=True,
        use_container_width=True,
        column_config={
            "ลำดับ": st.column_config.NumberColumn("ลำดับ", format="%.0f"),
            "Category": st.column_config.TextColumn("หมวดสินค้า"),
            "Product": st.column_config.TextColumn("สินค้า"),
            "System qty": st.column_config.NumberColumn("ในระบบ"),
            "Actual qty": st.column_config.NumberColumn("นับจริง", format="%.2f"),
            "Diff": st.column_config.NumberColumn("Diff", format="%.2f"),
            "Status": st.column_config.TextColumn("สถานะ"),
        },
    )

    export_df = result_df.copy()
    export_df["Actual qty"] = export_df["Actual qty"].fillna("")
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "ดาวน์โหลดผลลัพธ์ CSV",
        data=csv_bytes,
        file_name="stock-diff-result.csv",
        mime="text/csv",
        type="primary",
    )

    cost_rows = st.session_state.get("cost_rows", [])
    if cost_rows:
        st.subheader("ต้นทุนจากผลต่าง")
        cost_result_df = make_cost_result_table(result_df, cost_rows)
        if cost_result_df.empty:
            st.warning(
                "ยังไม่พบชื่อสินค้าที่ตรงกันระหว่าง Excel กับ PDF "
                "หรือรายการที่ตรงกันยังไม่ได้กรอกยอดนับจริง"
            )
        else:
            shortage_cost_signed_total = float(cost_result_df["Shortage cost"].sum())
            shortage_cost_total = abs(shortage_cost_signed_total)
            overage_cost_total = float(cost_result_df["Overage cost"].sum())
            st.caption(
                f"จับคู่ชื่อสินค้าได้ {len(cost_result_df)} รายการ "
                f"จากไฟล์ต้นทุน {len(cost_rows)} รายการ"
            )
            cost_metric_columns = st.columns(2)
            cost_metric_columns[0].metric(
                "ต้นทุนสินค้าขาด (บาท)", f"{shortage_cost_total:,.2f}"
            )
            cost_metric_columns[1].metric(
                "ต้นทุนสินค้าเกิน (บาท)", f"{overage_cost_total:,.2f}"
            )
            net_cost_total = overage_cost_total + shortage_cost_signed_total
            st.metric(
                "สรุปต้นทุนขาด/เกินสุทธิ (บาท)", f"{net_cost_total:,.2f}"
            )
            st.dataframe(
                cost_result_df.style.map(status_cell_style, subset=["Status"]),
                hide_index=True,
                use_container_width=True,
                column_config={
                    "ลำดับ": st.column_config.NumberColumn("ลำดับ", format="%.0f"),
                    "Category": st.column_config.TextColumn("หมวดสินค้า"),
                    "Product": st.column_config.TextColumn("สินค้า"),
                    "System qty": st.column_config.NumberColumn("ในระบบ"),
                    "Actual qty": st.column_config.NumberColumn(
                        "นับจริง", format="%.2f"
                    ),
                    "Diff": st.column_config.NumberColumn("Diff", format="%.2f"),
                    "Status": st.column_config.TextColumn("สถานะ"),
                    "Cost": st.column_config.NumberColumn(
                        "Cost ต่อหน่วย (บาท)", format="%.2f"
                    ),
                    "Shortage cost": st.column_config.NumberColumn(
                        "ต้นทุนขาด (บาท)", format="%.2f"
                    ),
                    "Overage cost": st.column_config.NumberColumn(
                        "ต้นทุนเกิน (บาท)", format="%.2f"
                    ),
                },
            )
            cost_csv_bytes = cost_result_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "ดาวน์โหลดผลต้นทุน CSV",
                data=cost_csv_bytes,
                file_name="stock-cost-result.csv",
                mime="text/csv",
            )
else:
    st.info("ใส่ยอดนับจริงอย่างน้อย 1 รายการ เพื่อดูผลต่างและดาวน์โหลดรายงาน")
