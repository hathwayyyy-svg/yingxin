
import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="补货单生成器", page_icon="📄", layout="wide")

# =========================
# 配置区（按你的模板固定）
# =========================
SOURCE_SHEET = "源文件"
TEMPLATE_SHEET = "生成pdf模板"
HELPER_SHEET = "辅助表"

# 模板明细行范围（保留原格式，仅覆盖这一区域）
DETAIL_START_ROW = 8
DETAIL_END_ROW = 20
TOTAL_ROW = 21

# 模板列定义
COL_INDEX = 1       # A  #
COL_PO = 2          # B  终端公司采购单号
COL_BRAND = 3       # C  品牌
COL_CODE = 4        # D  编码
COL_DESC = 5        # E  产品描述
COL_TOTAL_QTY = 6   # F  配送单总台数
COL_SEND_QTY = 7    # G  本次配送台数
COL_BOXES = 8       # H  箱数
COL_ORDER_NO = 9    # I  补货单号
COL_CMD_NO = 10     # J  指令编号
COL_REMARK = 11     # K  备注

YELLOW_FILL = PatternFill(fill_type="solid", fgColor="FFF59D")


# =========================
# 通用函数
# =========================
def normalize_text(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("　", " ")
    s = re.sub(r"\s+", "", s)
    return s


def normalize_model(model: str) -> str:
    s = normalize_text(model)
    for token in ["分销公开版", "公开版", "分销权益版", "权益版", "销售用机"]:
        s = s.replace(token, "")
    return s


def safe_int(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0
    try:
        return int(float(val))
    except Exception:
        return 0


def make_po_short(order_no: str) -> str:
    """
    BH-ZD-450145083320 -> 4501450833
    若不符合规则，尽量提取中间10位数字
    """
    if not order_no:
        return ""
    s = str(order_no)
    m = re.search(r"(\d{10})\d{2,}$", s)
    if m:
        return m.group(1)
    digits = re.sub(r"\D", "", s)
    return digits[:10] if len(digits) >= 10 else digits


def is_merged_cell(ws, row, col) -> bool:
    coord = f"{get_column_letter(col)}{row}"
    for rng in ws.merged_cells.ranges:
        if coord in rng:
            # 只有左上角可写，其余 merged cell 只读
            return not (row == rng.min_row and col == rng.min_col)
    return False


def safe_set(ws, row, col, value=None, fill=None):
    if is_merged_cell(ws, row, col):
        return
    cell = ws.cell(row, col)
    if value is not None:
        cell.value = value
    if fill is not None:
        cell.fill = fill


def clear_detail_area(ws):
    for r in range(DETAIL_START_ROW, DETAIL_END_ROW + 1):
        for c in range(1, 12):
            if is_merged_cell(ws, r, c):
                continue
            cell = ws.cell(r, c)
            cell.value = None
            # 保留原样式，不清边框/字体/对齐
            # 恢复填充为无色，避免上次标黄残留
            if c in (COL_CODE, COL_DESC):
                cell.fill = copy(ws.cell(DETAIL_START_ROW, c).fill)


def clear_total_area(ws):
    # 只清 A21:H21，I:K 本身是合并单元格
    for c in range(1, 9):
        if not is_merged_cell(ws, TOTAL_ROW, c):
            ws.cell(TOTAL_ROW, c).value = None


def build_helper_map(df_helper: pd.DataFrame):
    helper_map = {}
    for _, row in df_helper.iterrows():
        model = normalize_model(row.get("CTMS机型"))
        color = normalize_text(row.get("颜色"))
        key = (model, color)
        helper_map[key] = {
            "code": "" if pd.isna(row.get("物料编码")) else str(row.get("物料编码")).strip(),
            "desc": "" if pd.isna(row.get("SCM物料描述")) else str(row.get("SCM物料描述")).strip(),
        }
    return helper_map


def detect_region(text: str) -> str:
    s = str(text or "")
    if "广州" in s:
        return "广州分销"
    if "粤北" in s:
        return "粤北分销"
    return ""


def find_date_from_template(ws):
    val = ws["C2"].value
    if isinstance(val, datetime):
        return val
    if val:
        try:
            return pd.to_datetime(val).to_pydatetime()
        except Exception:
            return datetime.now()
    return datetime.now()


def prepare_region_df(df_source: pd.DataFrame, region: str) -> pd.DataFrame:
    df = df_source.copy()
    df["__region__"] = df["备注"].apply(detect_region)
    df = df[df["__region__"] == region].copy()
    return df


def write_detail_rows(ws, df_region: pd.DataFrame, helper_map: dict):
    missing_items = []
    total_send_qty = 0
    total_boxes = 0

    max_rows = DETAIL_END_ROW - DETAIL_START_ROW + 1
    df_region = df_region.head(max_rows).copy()

    for i, (_, row) in enumerate(df_region.iterrows(), start=0):
        excel_row = DETAIL_START_ROW + i

        order_no = "" if pd.isna(row.get("补货单号")) else str(row.get("补货单号")).strip()
        brand = "" if pd.isna(row.get("品牌")) else str(row.get("品牌")).strip()
        model_raw = "" if pd.isna(row.get("型号")) else str(row.get("型号")).strip()
        color_raw = "" if pd.isna(row.get("颜色")) else str(row.get("颜色")).strip()
        total_qty = safe_int(row.get("配送单总台数"))
        send_qty = safe_int(row.get("本次配送台数"))
        remark = "" if pd.isna(row.get("备注")) else str(row.get("备注")).strip()

        key = (normalize_model(model_raw), normalize_text(color_raw))
        helper_info = helper_map.get(key, {})
        material_code = helper_info.get("code", "")
        scm_desc = helper_info.get("desc", "")

        if not material_code or not scm_desc:
            missing_items.append(f"{model_raw} | {color_raw}")

        boxes = send_qty // 10 if send_qty else 0
        po_short = make_po_short(order_no)

        safe_set(ws, excel_row, COL_INDEX, i + 1)
        safe_set(ws, excel_row, COL_PO, po_short)
        safe_set(ws, excel_row, COL_BRAND, brand)
        safe_set(ws, excel_row, COL_CODE, material_code if material_code else "")
        safe_set(ws, excel_row, COL_DESC, scm_desc if scm_desc else "")
        safe_set(ws, excel_row, COL_TOTAL_QTY, total_qty)
        safe_set(ws, excel_row, COL_SEND_QTY, send_qty)
        safe_set(ws, excel_row, COL_BOXES, boxes)
        safe_set(ws, excel_row, COL_ORDER_NO, order_no)
        safe_set(ws, excel_row, COL_CMD_NO, "")
        safe_set(ws, excel_row, COL_REMARK, remark)

        if not material_code:
            safe_set(ws, excel_row, COL_CODE, fill=YELLOW_FILL)
        if not scm_desc:
            safe_set(ws, excel_row, COL_DESC, fill=YELLOW_FILL)

        total_send_qty += send_qty
        total_boxes += boxes

    safe_set(ws, TOTAL_ROW, 1, "合计")
    safe_set(ws, TOTAL_ROW, 6, total_send_qty)
    safe_set(ws, TOTAL_ROW, 7, total_send_qty)
    safe_set(ws, TOTAL_ROW, 8, total_boxes)

    return {
        "detail_count": len(df_region),
        "missing_count": len(missing_items),
        "missing_items": missing_items,
        "overflow_count": max(0, len(df_region) - max_rows),
    }


def fill_template_workbook(uploaded_bytes: bytes, region: str):
    wb = load_workbook(io.BytesIO(uploaded_bytes))
    ws_template = wb[TEMPLATE_SHEET]

    df_source = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=SOURCE_SHEET)
    df_helper = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=HELPER_SHEET)

    # 去掉首行“求和项...”之类脏头
    df_source.columns = [str(c).strip() for c in df_source.columns]
    # 若首列不是“补货单号”，尝试重新读取 header=1
    if "补货单号" not in df_source.columns:
        df_source = pd.read_excel(io.BytesIO(uploaded_bytes), sheet_name=SOURCE_SHEET, header=1)

    helper_map = build_helper_map(df_helper)
    df_region = prepare_region_df(df_source, region)

    clear_detail_area(ws_template)
    clear_total_area(ws_template)
    summary = write_detail_rows(ws_template, df_region, helper_map)

    return wb, summary


def workbook_to_bytes(wb):
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def convert_excel_to_pdf(xlsx_path: str, out_dir: str):
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not libreoffice:
        return False, "当前环境未安装 LibreOffice，已生成 Excel，无法原样转 PDF。"

    try:
        cmd = [
            libreoffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", out_dir,
            xlsx_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            msg = (result.stderr or result.stdout or "LibreOffice 转换失败").strip()
            return False, msg

        pdf_path = str(Path(out_dir) / (Path(xlsx_path).stem + ".pdf"))
        if not os.path.exists(pdf_path):
            return False, "已执行转换命令，但未找到生成的 PDF。"
        return True, pdf_path
    except Exception as e:
        return False, str(e)


def make_zip(file_map: dict) -> bytes:
    """
    file_map: {filename: bytes}
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in file_map.items():
            zf.writestr(name, data)
    mem.seek(0)
    return mem.getvalue()


# =========================
# 页面
# =========================
st.title("补货单生成器（稳定版）")
st.caption("上传包含「源文件 / 生成pdf模板 / 辅助表」的 Excel，自动生成广州分销、粤北分销补货单，并在当前环境支持时原样转 PDF。")

with st.expander("使用说明", expanded=False):
    st.markdown(
        """
1. 上传一个 Excel 文件，必须包含 3 个 sheet：`源文件`、`生成pdf模板`、`辅助表`
2. 程序会按 `备注` 自动拆分：
   - 包含“广州” → 广州分销
   - 包含“粤北” → 粤北分销
3. 产品描述与编码，按 `型号 + 颜色` 去 `辅助表` 匹配：
   - `物料编码` → 模板“编码”
   - `SCM物料描述` → 模板“产品描述”
4. **不改模板格式**，仅填充数据；PDF 采用 **Excel 原样转 PDF**
        """
    )

uploaded = st.file_uploader("上传补货单源文件（xlsx）", type=["xlsx"])

if uploaded:
    uploaded_bytes = uploaded.getvalue()

    # 基础校验
    try:
        check_wb = load_workbook(io.BytesIO(uploaded_bytes), read_only=True)
        sheetnames = set(check_wb.sheetnames)
        required = {SOURCE_SHEET, TEMPLATE_SHEET, HELPER_SHEET}
        missing = required - sheetnames
        if missing:
            st.error(f"缺少必要 sheet：{', '.join(sorted(missing))}")
            st.stop()
    except Exception as e:
        st.error(f"文件无法读取，请确认是有效的 xlsx：{e}")
        st.stop()

    if st.button("一键生成", type="primary", use_container_width=True):
        results = []
        generated_files = {}

        for region in ["广州分销", "粤北分销"]:
            try:
                wb_region, summary = fill_template_workbook(uploaded_bytes, region)
                excel_name = f"{region}补货单.xlsx"
                excel_bytes = workbook_to_bytes(wb_region)
                generated_files[excel_name] = excel_bytes

                pdf_status = "未转换"
                pdf_name = f"{region}补货单.pdf"

                with tempfile.TemporaryDirectory() as tmpdir:
                    xlsx_path = os.path.join(tmpdir, excel_name)
                    with open(xlsx_path, "wb") as f:
                        f.write(excel_bytes)

                    ok, pdf_result = convert_excel_to_pdf(xlsx_path, tmpdir)
                    if ok:
                        with open(pdf_result, "rb") as f:
                            generated_files[pdf_name] = f.read()
                        pdf_status = "成功"
                    else:
                        pdf_status = f"失败：{pdf_result}"

                results.append({
                    "区域": region,
                    "明细行数": summary["detail_count"],
                    "缺失编码/描述数": summary["missing_count"],
                    "PDF转换": pdf_status,
                    "缺失明细": "；".join(summary["missing_items"][:10]) if summary["missing_items"] else ""
                })
            except Exception as e:
                results.append({
                    "区域": region,
                    "明细行数": 0,
                    "缺失编码/描述数": 0,
                    "PDF转换": f"失败：{e}",
                    "缺失明细": ""
                })

        st.subheader("生成结果")
        df_result = pd.DataFrame(results)
        st.dataframe(df_result, use_container_width=True, hide_index=True)

        if generated_files:
            st.success(f"已生成 {len(generated_files)} 个文件。")

            col1, col2 = st.columns(2)
            with col1:
                for name, data in generated_files.items():
                    if name.endswith(".xlsx") and "广州分销" in name:
                        st.download_button(f"下载 {name}", data=data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    if name.endswith(".pdf") and "广州分销" in name:
                        st.download_button(f"下载 {name}", data=data, file_name=name, mime="application/pdf", use_container_width=True)
            with col2:
                for name, data in generated_files.items():
                    if name.endswith(".xlsx") and "粤北分销" in name:
                        st.download_button(f"下载 {name}", data=data, file_name=name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                    if name.endswith(".pdf") and "粤北分销" in name:
                        st.download_button(f"下载 {name}", data=data, file_name=name, mime="application/pdf", use_container_width=True)

            zip_bytes = make_zip(generated_files)
            st.download_button(
                "一键下载全部（ZIP）",
                data=zip_bytes,
                file_name="补货单生成结果.zip",
                mime="application/zip",
                use_container_width=True
            )
        else:
            st.error("没有成功生成任何文件。")
