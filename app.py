import math
import re
import zipfile
from copy import copy
from datetime import datetime
from io import BytesIO

import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill

st.set_page_config(page_title="OPPO补货单模板生成", page_icon="📦", layout="wide")

TITLE = "OPPO补货单模板生成工具"

REGION_CONFIG = {
    "广州分销": {
        "source_keywords": ["广州"],
        "receiver_address": "广东省广州市花都区花东镇金港北一路3号J8栋301单元",
    },
    "粤北分销": {
        "source_keywords": ["粤北"],
        "receiver_address": "广东省广州市花都区花东镇金谷工业园永大路7号10平台",
    },
}

DATA_START_ROW = 8
YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")


def normalize_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", "", s)
    return s


def clean_model(model: str) -> str:
    s = str(model or "").strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("分销公开版", "").replace("公开版", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def model_key(model, color):
    return f"{normalize_text(clean_model(model))}|{normalize_text(color)}"


def extract_po_no(replenish_no: str) -> str:
    m = re.search(r"(\d{10})\d*$", str(replenish_no or ""))
    return m.group(1) if m else ""


def calc_boxes(qty):
    try:
        qty = float(qty)
    except Exception:
        return ""
    return int(math.ceil(qty / 10)) if qty > 0 else 0


def infer_brand_cn(brand):
    brand = str(brand or "").strip().upper()
    if brand == "OPPO":
        return "欧珀"
    if brand in {"一加", "ONEPLUS"}:
        return "一加"
    return brand or ""


def parse_model_description(brand, model, color):
    brand_cn = infer_brand_cn(brand)
    raw = clean_model(model)
    raw = raw.replace("(", "（").replace(")", "）")

    base = raw.split("（")[0].strip()
    code_part = ""
    if "（" in raw and "）" in raw:
        code_part = raw.split("（", 1)[1].split("）", 1)[0].strip()

    mem_match = re.search(r"(\d+G)\+(\d+G|1T)", code_part, re.I)
    memory = ""
    storage = ""
    if mem_match:
        memory = mem_match.group(1).upper().replace("G", "")
        storage = mem_match.group(2).upper()

    display_base = base
    if brand_cn == "欧珀":
        if re.fullmatch(r"Reno\d+[A-Za-z]?", base, re.I):
            display_base = base.upper()
        elif base.lower().startswith("find "):
            display_base = base
        elif re.fullmatch(r"A\d+[A-Za-z]?", base, re.I):
            display_base = base

    if memory and storage:
        title = f"{brand_cn}_{display_base} {memory}+{storage}"
    else:
        title = f"{brand_cn}_{display_base}"

    desc_parts = [title, str(color or "").strip()]
    if storage:
        desc_parts.append(storage)
    desc_parts.extend(["公开版", "销售用机"])
    return ",".join([x for x in desc_parts if x])


def is_merged_proxy(cell):
    return isinstance(cell, MergedCell)


def safe_clear_cell(ws, row, col):
    cell = ws.cell(row, col)
    if is_merged_proxy(cell):
        return
    cell.value = None


def safe_write(ws, row, col, value):
    cell = ws.cell(row, col)
    if is_merged_proxy(cell):
        return
    cell.value = value


def safe_fill(ws, row, col, fill):
    cell = ws.cell(row, col)
    if is_merged_proxy(cell):
        return
    cell.fill = fill


def copy_sheet(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            if is_merged_proxy(cell):
                continue
            new_cell = dst_ws[cell.coordinate]
            new_cell.value = cell.value
            if cell.has_style:
                new_cell._style = copy(cell._style)
            new_cell.number_format = cell.number_format
            new_cell.font = copy(cell.font)
            new_cell.fill = copy(cell.fill)
            new_cell.border = copy(cell.border)
            new_cell.alignment = copy(cell.alignment)
            new_cell.protection = copy(cell.protection)

    for col_letter, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = dim.width
        dst_ws.column_dimensions[col_letter].hidden = dim.hidden
    for row_idx, dim in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[row_idx].height = dim.height
        dst_ws.row_dimensions[row_idx].hidden = dim.hidden

    for merged_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merged_range))

    dst_ws.sheet_view.showGridLines = src_ws.sheet_view.showGridLines
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.page_margins = copy(src_ws.page_margins)
    dst_ws.page_setup = copy(src_ws.page_setup)
    dst_ws.print_options = copy(src_ws.print_options)
    dst_ws.sheet_properties = copy(src_ws.sheet_properties)
    dst_ws.print_title_rows = src_ws.print_title_rows
    dst_ws.print_title_cols = src_ws.print_title_cols


def build_lookup(source_ws, template_ws):
    template_by_replenish = {}
    for r in range(DATA_START_ROW, template_ws.max_row + 1):
        replenish_no = template_ws.cell(r, 9).value
        if replenish_no:
            template_by_replenish[str(replenish_no).strip()] = {
                "brand": template_ws.cell(r, 3).value,
                "code": template_ws.cell(r, 4).value,
                "desc": template_ws.cell(r, 5).value,
            }

    lookup = {}
    for r in range(3, source_ws.max_row + 1):
        replenish_no = source_ws.cell(r, 1).value
        model = source_ws.cell(r, 3).value
        color = source_ws.cell(r, 4).value
        if replenish_no and model and color and str(replenish_no).strip() in template_by_replenish:
            lookup[model_key(model, color)] = template_by_replenish[str(replenish_no).strip()]
    return lookup


def collect_source_rows(source_ws, region_name):
    cfg = REGION_CONFIG[region_name]
    out = []
    for r in range(3, source_ws.max_row + 1):
        row = {
            "补货单号": source_ws.cell(r, 1).value,
            "品牌": source_ws.cell(r, 2).value,
            "型号": source_ws.cell(r, 3).value,
            "颜色": source_ws.cell(r, 4).value,
            "仓1": source_ws.cell(r, 5).value,
            "仓2": source_ws.cell(r, 6).value,
            "配送单总台数": source_ws.cell(r, 7).value,
            "本次配送台数": source_ws.cell(r, 8).value,
            "备注": source_ws.cell(r, 9).value,
        }
        remark = str(row["备注"] or "")
        if any(k in remark for k in cfg["source_keywords"]):
            out.append(row)
    return out


def clear_data_rows(ws):
    for r in range(DATA_START_ROW, ws.max_row + 1):
        for c in range(1, 12):
            safe_clear_cell(ws, r, c)


def format_data_row_from_template(ws, template_row_idx, target_row_idx):
    for c in range(1, 12):
        src = ws.cell(template_row_idx, c)
        dst = ws.cell(target_row_idx, c)
        if is_merged_proxy(src) or is_merged_proxy(dst):
            continue
        if src.has_style:
            dst._style = copy(src._style)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)
    ws.row_dimensions[target_row_idx].height = ws.row_dimensions[template_row_idx].height


def build_region_workbook(uploaded_bytes: bytes, region_name: str, delivery_date: datetime):
    wb = load_workbook(BytesIO(uploaded_bytes))
    if "源文件" not in wb.sheetnames or "生成pdf模板" not in wb.sheetnames:
        raise ValueError("上传文件必须包含【源文件】和【生成pdf模板】两个sheet。")

    source_ws = wb["源文件"]
    template_ws = wb["生成pdf模板"]
    lookup = build_lookup(source_ws, template_ws)
    rows = collect_source_rows(source_ws, region_name)
    if not rows:
        raise ValueError(f"没有找到【{region_name}】对应的数据。请检查备注列是否包含该区域关键字。")

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = f"{region_name}模板"
    copy_sheet(template_ws, out_ws)
    clear_data_rows(out_ws)

    safe_write(out_ws, 2, 3, delivery_date)
    safe_write(out_ws, 4, 7, REGION_CONFIG[region_name]["receiver_address"])

    sample_data_row = DATA_START_ROW
    missing_codes = []
    total_total_qty = 0
    total_delivery_qty = 0
    total_boxes = 0

    for idx, row in enumerate(rows, start=1):
        target_row = DATA_START_ROW + idx - 1
        format_data_row_from_template(out_ws, sample_data_row, target_row)

        key = model_key(row["型号"], row["颜色"])
        matched = lookup.get(key, {})
        code = matched.get("code", "")
        desc = matched.get("desc") or parse_model_description(row["品牌"], row["型号"], row["颜色"])
        if not code:
            missing_codes.append(f"{clean_model(row['型号'])} | {row['颜色']}")

        total_qty = row["配送单总台数"] or 0
        delivery_qty = row["本次配送台数"] or 0
        boxes = calc_boxes(delivery_qty)

        safe_write(out_ws, target_row, 1, idx)
        safe_write(out_ws, target_row, 2, extract_po_no(row["补货单号"]))
        safe_write(out_ws, target_row, 3, row["品牌"])
        safe_write(out_ws, target_row, 4, code)
        safe_write(out_ws, target_row, 5, desc)
        safe_write(out_ws, target_row, 6, total_qty)
        safe_write(out_ws, target_row, 7, delivery_qty)
        safe_write(out_ws, target_row, 8, boxes)
        safe_write(out_ws, target_row, 9, row["补货单号"])
        safe_write(out_ws, target_row, 10, None)
        safe_write(out_ws, target_row, 11, row["备注"])

        if not code:
            safe_fill(out_ws, target_row, 4, YELLOW_FILL)

        total_total_qty += int(total_qty or 0)
        total_delivery_qty += int(delivery_qty or 0)
        total_boxes += int(boxes or 0)

    total_row = DATA_START_ROW + len(rows)
    format_data_row_from_template(out_ws, sample_data_row, total_row)
    safe_write(out_ws, total_row, 1, "合计：")
    safe_write(out_ws, total_row, 6, total_total_qty)
    safe_write(out_ws, total_row, 7, total_delivery_qty)
    safe_write(out_ws, total_row, 8, total_boxes)

    sign_rows = ["签收人：", "签收数量：", "签收日期、时间："]
    for offset, label in enumerate(sign_rows, start=1):
        r = total_row + offset
        format_data_row_from_template(out_ws, sample_data_row, r)
        safe_write(out_ws, r, 1, label)
        for c in range(2, 12):
            safe_write(out_ws, r, c, None)

    out_ws.print_area = f"A1:K{total_row + len(sign_rows)}"

    bio = BytesIO()
    out_wb.save(bio)
    bio.seek(0)
    return bio.getvalue(), rows, sorted(set(missing_codes))


def build_zip(file_map):
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for filename, content in file_map.items():
            zf.writestr(filename, content)
    buf.seek(0)
    return buf.getvalue()


st.title(TITLE)
st.caption("上传一个带【源文件】和【生成pdf模板】sheet的 Excel，自动拆分生成广州分销 / 粤北分销模板。")

with st.expander("使用说明", expanded=True):
    st.markdown(
        """
1. 上传原始 Excel 文件。
2. 工具会按【源文件】sheet 的备注列自动识别“广州分销”“粤北分销”。
3. 自动生成两个模板文件，可直接下载 ZIP。
4. 若某个型号+颜色在模板示例中找不到对应编码，会自动高亮黄色，方便你复核。
        """
    )

uploaded = st.file_uploader("上传源文件（.xlsx）", type=["xlsx"])
default_date = datetime.today()
delivery_date = st.date_input("预计送货日期", value=default_date)

if uploaded:
    file_bytes = uploaded.getvalue()

    if st.button("一键生成模板", type="primary"):
        outputs = {}
        summary_rows = []

        for region in REGION_CONFIG:
            try:
                content, rows, missing_codes = build_region_workbook(
                    file_bytes,
                    region,
                    datetime.combine(delivery_date, datetime.min.time()),
                )
                filename = f"{region}补货单模板.xlsx"
                outputs[filename] = content
                summary_rows.append({
                    "区域": region,
                    "明细行数": len(rows),
                    "缺失编码数": len(missing_codes),
                    "缺失编码明细": "；".join(missing_codes) if missing_codes else "无",
                })
            except Exception as e:
                summary_rows.append({
                    "区域": region,
                    "明细行数": 0,
                    "缺失编码数": 0,
                    "缺失编码明细": f"生成失败：{e}",
                })

        st.subheader("生成结果")
        st.dataframe(summary_rows, use_container_width=True, hide_index=True)

        if outputs:
            zip_bytes = build_zip(outputs)
            st.download_button(
                "下载全部模板 ZIP",
                data=zip_bytes,
                file_name="补货单模板结果.zip",
                mime="application/zip",
            )
            for filename, content in outputs.items():
                st.download_button(
                    f"单独下载：{filename}",
                    data=content,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error("没有成功生成任何模板，请检查源文件格式。")
else:
    st.info("先上传 Excel 文件，再点击“一键生成模板”。")
