import io
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": NS_MAIN, "r": NS_REL, "pkg": NS_PKG_REL}


def normalize_header(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().replace(" ", "").lower()


def find_plan_index(row_values) -> int:
    for i, cell in enumerate(row_values):
        normalized = normalize_header(cell)
        if normalized == "plan" or normalized.startswith("plan"):
            return i
    return -1


def extract_structured_data(df: pd.DataFrame) -> list[dict]:
    structured_data = []
    current_plan_title = "未知Plan"
    current_test_title = "未知測試項目"
    header_row_index = -1
    headers = []
    header_indices = {}

    for index, row in df.iterrows():
        row_str = row.astype(str).tolist()

        plan_index = find_plan_index(row.tolist())
        if plan_index != -1:
            plan_text = row.iloc[8] if len(row) > 8 else None
            if pd.isna(plan_text):
                for j in range(plan_index + 1, len(row)):
                    if pd.notna(row.iloc[j]):
                        plan_text = row.iloc[j]
                        break
            current_plan_title = str(plan_text).strip() if pd.notna(plan_text) else "未知Plan"
            current_test_title = "未知測試項目"
            header_row_index = -1
            headers = []
            header_indices = {}

        if any(cell.strip().lower() == "test" for cell in row_str):
            try:
                test_index = next(i for i, cell in enumerate(row_str) if cell.strip().lower() == "test")
                title_text = row.iloc[9] if len(row) > 9 else None
                if pd.isna(title_text):
                    for j in range(test_index + 1, len(row)):
                        if pd.notna(row.iloc[j]):
                            title_text = row.iloc[j]
                            break
                current_test_title = str(title_text).strip() if pd.notna(title_text) else "未知測試項目"
                header_row_index = -1
                headers = []
                header_indices = {}
            except Exception:
                continue

        if any("label" in cell.lower() for cell in row_str):
            header_row_index = index
            headers = [str(c).strip() for c in row.tolist()]
            header_indices = {}
            for i, cell in enumerate(row.tolist()):
                normalized = normalize_header(cell)
                if normalized in ("label/value", "actual", "result", "%error"):
                    header_indices[normalized] = i
            continue

        if header_row_index != -1 and any(pd.notna(cell) for cell in row):
            if any(cell.strip().lower() == "test" for cell in row_str):
                header_row_index = -1
                headers = []
                header_indices = {}
                continue

            data_values = row.tolist()
            if len(headers) == len(data_values) and "label/value" in header_indices:
                data_dict = {"Plan": current_plan_title, "Test": current_test_title}
                data_dict["Label/Value"] = data_values[header_indices["label/value"]]
                if "actual" in header_indices:
                    data_dict["Actual"] = data_values[header_indices["actual"]]
                if "result" in header_indices:
                    data_dict["Result"] = data_values[header_indices["result"]]
                if "%error" in header_indices:
                    data_dict["% Error"] = data_values[header_indices["%error"]]
                structured_data.append(data_dict)
            elif len(headers) != len(data_values):
                header_row_index = -1

    return structured_data


def build_data_maps(cleaned_df: pd.DataFrame) -> dict:
    all_data_maps = {}
    plan_groups = cleaned_df.groupby("Plan", sort=False)

    for plan_name, plan_df in plan_groups:
        data_map = {}
        for _, row in plan_df.iterrows():
            try:
                test_name = str(row["Test"]).strip()
                if "復閉" in test_name or "始動" in test_name or "瞬跳" in test_name:
                    key = test_name
                else:
                    raw = str(row["Label/Value"]).strip().split()[0]
                    val = float(raw)
                    suffix = str(int(val)) if val == int(val) else f"{val:.2f}".rstrip("0")
                    key = f"{test_name}{suffix}"
                data_map[key] = row["Actual"]
            except Exception:
                continue

        sheet_name = str(plan_name).strip()
        all_data_maps[sheet_name] = data_map

    return all_data_maps


def to_excel_value(value):
    """Convert textual numbers to numeric values to avoid Excel warning icons."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    if text == "":
        return None

    # Remove thousands separators before numeric parsing.
    candidate = text.replace(",", "")
    try:
        number = float(candidate)
        return int(number) if number.is_integer() else number
    except ValueError:
        return text


def qn(local_name: str) -> str:
    return f"{{{NS_MAIN}}}{local_name}"


def normalize_zip_path(path: str) -> str:
    normalized = path.replace("\\", "/")
    if normalized.startswith("/"):
        normalized = normalized[1:]
    if not normalized.startswith("xl/"):
        normalized = f"xl/{normalized}"
    return normalized


def parse_cell_ref(ref: str):
    m = re.match(r"^([A-Z]+)(\d+)$", ref or "")
    if not m:
        return None
    return m.group(1), int(m.group(2))


def column_to_index(col_letters: str) -> int:
    index = 0
    for ch in col_letters:
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index


def load_shared_strings(zip_entries: dict[str, bytes]) -> list[str]:
    shared = zip_entries.get("xl/sharedStrings.xml")
    if shared is None:
        return []
    root = ET.fromstring(shared)
    result = []
    for si in root.findall("main:si", NS):
        parts = []
        for t in si.findall(".//main:t", NS):
            parts.append(t.text or "")
        result.append("".join(parts))
    return result


def get_cell_value(cell, shared_strings: list[str]):
    if cell is None:
        return None

    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        t_node = cell.find("main:is/main:t", NS)
        return (t_node.text if t_node is not None else None)

    v_node = cell.find("main:v", NS)
    if v_node is None or v_node.text is None:
        return None
    raw = v_node.text

    if cell_type == "s":
        try:
            idx = int(float(raw))
            return shared_strings[idx] if 0 <= idx < len(shared_strings) else None
        except Exception:
            return None
    return raw


def set_cell_value(cell, value) -> None:
    for node_name in ("f", "v", "is"):
        for child in cell.findall(f"main:{node_name}", NS):
            cell.remove(child)

    if value is None:
        cell.attrib.pop("t", None)
        return

    if isinstance(value, bool):
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, qn("v"))
        v.text = "1" if value else "0"
        return

    if isinstance(value, (int, float)):
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, qn("v"))
        v.text = str(value)
        return

    text = str(value)
    if text == "":
        cell.attrib.pop("t", None)
        return

    cell.set("t", "inlineStr")
    is_node = ET.SubElement(cell, qn("is"))
    t_node = ET.SubElement(is_node, qn("t"))
    if text != text.strip():
        t_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_node.text = text


def get_or_create_row(sheet_data, row_map: dict[int, ET.Element], row_idx: int):
    row = row_map.get(row_idx)
    if row is not None:
        return row

    row = ET.Element(qn("row"), {"r": str(row_idx)})
    rows = list(sheet_data.findall("main:row", NS))
    inserted = False
    for i, existing in enumerate(rows):
        try:
            existing_idx = int(existing.get("r", "0"))
        except ValueError:
            existing_idx = 0
        if existing_idx > row_idx:
            sheet_data.insert(i, row)
            inserted = True
            break
    if not inserted:
        sheet_data.append(row)
    row_map[row_idx] = row
    return row


def get_or_create_cell(row_elem, cell_ref: str):
    for c in row_elem.findall("main:c", NS):
        if c.get("r") == cell_ref:
            return c

    target = ET.Element(qn("c"), {"r": cell_ref})
    parsed = parse_cell_ref(cell_ref)
    target_idx = column_to_index(parsed[0]) if parsed else 0

    cells = list(row_elem.findall("main:c", NS))
    inserted = False
    for i, existing in enumerate(cells):
        parsed_existing = parse_cell_ref(existing.get("r", ""))
        if not parsed_existing:
            continue
        if column_to_index(parsed_existing[0]) > target_idx:
            row_elem.insert(i, target)
            inserted = True
            break
    if not inserted:
        row_elem.append(target)
    return target


def to_xml_bytes(root: ET.Element) -> bytes:
    buffer = io.BytesIO()
    ET.ElementTree(root).write(buffer, encoding="utf-8", xml_declaration=True)
    return buffer.getvalue()


def process_workbook(source_files, target_file):
    all_structured_data = []
    parse_logs = []

    for src in source_files:
        try:
            src.seek(0)
            df = pd.read_excel(src, header=None)
            structured_data = extract_structured_data(df)
            if not structured_data:
                parse_logs.append(f"⚠️ {src.name}: 未擷取到資料，已略過")
                continue
            all_structured_data.extend(structured_data)
            parse_logs.append(f"✅ {src.name}: 擷取 {len(structured_data)} 筆")
        except Exception as exc:
            parse_logs.append(f"❌ {src.name}: 讀取失敗（{exc}）")

    final_df = pd.DataFrame(all_structured_data)
    if final_df.empty:
        raise ValueError("未能成功擷取任何來源資料，請檢查 Excel 格式")

    keep_cols = ["Plan", "Test", "Label/Value", "Actual", "Result", "% Error"]
    existing_cols = [c for c in keep_cols if c in final_df.columns]
    cleaned_df = final_df[existing_cols].copy()

    if "Actual" in cleaned_df.columns:
        cleaned_df.loc[:, "Actual"] = cleaned_df.loc[:, "Actual"].apply(
            lambda x: re.sub(r"[sva]", "", str(x), flags=re.IGNORECASE) if pd.notna(x) else x
        )

    all_data_maps = build_data_maps(cleaned_df)
    if not all_data_maps:
        raise ValueError("沒有可用的 Plan 對應資料")

    target_file.seek(0)
    target_bytes = target_file.read()
    with zipfile.ZipFile(io.BytesIO(target_bytes), "r") as zin:
        zip_entries = {name: zin.read(name) for name in zin.namelist()}

    workbook_xml = zip_entries.get("xl/workbook.xml")
    workbook_rels = zip_entries.get("xl/_rels/workbook.xml.rels")
    if workbook_xml is None or workbook_rels is None:
        raise ValueError("目標檔案不是有效的 Excel 結構，無法處理。")

    wb_root = ET.fromstring(workbook_xml)
    rels_root = ET.fromstring(workbook_rels)
    shared_strings = load_shared_strings(zip_entries)

    rid_to_sheet_path = {}
    for rel in rels_root.findall("pkg:Relationship", NS):
        if rel.get("Type", "").endswith("/worksheet"):
            rel_id = rel.get("Id")
            target = rel.get("Target")
            if rel_id and target:
                rid_to_sheet_path[rel_id] = normalize_zip_path(target)

    matched_count = 0
    skipped_count = 0
    write_logs = []

    for sheet in wb_root.findall("main:sheets/main:sheet", NS):
        ws_name = sheet.get("name", "")
        if not ws_name.startswith("特性_"):
            continue

        rid = sheet.get(f"{{{NS_REL}}}id")
        sheet_path = rid_to_sheet_path.get(rid or "")
        if not sheet_path or sheet_path not in zip_entries:
            skipped_count += 1
            write_logs.append(f"--- 跳過 [{ws_name}]：找不到工作表資料")
            continue

        ws_root = ET.fromstring(zip_entries[sheet_path])
        sheet_data = ws_root.find("main:sheetData", NS)
        if sheet_data is None:
            skipped_count += 1
            write_logs.append(f"--- 跳過 [{ws_name}]：工作表缺少 sheetData")
            continue

        row_map: dict[int, ET.Element] = {}
        cell_map: dict[str, ET.Element] = {}
        for row_elem in sheet_data.findall("main:row", NS):
            try:
                row_idx = int(row_elem.get("r", "0"))
            except ValueError:
                continue
            row_map[row_idx] = row_elem
            for cell in row_elem.findall("main:c", NS):
                ref = cell.get("r")
                if ref:
                    cell_map[ref] = cell

        o1_value = get_cell_value(cell_map.get("O1"), shared_strings)

        if o1_value is None:
            skipped_count += 1
            write_logs.append(f"--- 跳過 [{ws_name}]：O1 為空")
            continue

        panel_name = str(o1_value).strip()
        if panel_name not in all_data_maps:
            skipped_count += 1
            write_logs.append(f"--- 跳過 [{ws_name}]：找不到盤名 [{panel_name}]")
            continue

        current_map = all_data_maps[panel_name]
        write_count = 0

        for cell_ref, cell in list(cell_map.items()):
            parsed = parse_cell_ref(cell_ref)
            if not parsed or parsed[0] != "CN":
                continue

            row_idx = parsed[1]
            label = get_cell_value(cell, shared_strings)
            if label is not None and label in current_map:
                target_ref = f"CO{row_idx}"
                row_elem = get_or_create_row(sheet_data, row_map, row_idx)
                target_cell = get_or_create_cell(row_elem, target_ref)
                set_cell_value(target_cell, to_excel_value(current_map[label]))
                write_count += 1

        matched_count += 1
        write_logs.append(f">>> [{ws_name}] 盤名 [{panel_name}]：寫入 {write_count} 筆")
        zip_entries[sheet_path] = to_xml_bytes(ws_root)

    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in zip_entries.items():
            zout.writestr(name, data)
    output.seek(0)

    return {
        "file_bytes": output.getvalue(),
        "cleaned_df": cleaned_df,
        "parse_logs": parse_logs,
        "write_logs": write_logs,
        "matched_count": matched_count,
        "skipped_count": skipped_count,
        "plan_count": len(all_data_maps),
    }

def render_template_downloads() -> None:
    st.subheader("檔案下載")

    assets = [
        (
            Path("assets") / "公版_IED試驗報告.xlsm",
            "下載試驗報告",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ),
        (
            Path("assets") / "公版_變電所_測試程序.psx",
            "下載測試程序",
            "application/octet-stream",
        ),
    ]

    for path, label, mime in assets:
        if path.exists():
            with path.open("rb") as f:
                st.download_button(
                    label=label,
                    data=f.read(),
                    file_name=path.name,
                    mime=mime,
                    use_container_width=True,
                )
        else:
            st.info(f"尚未找到 `{path}`")







st.set_page_config(page_title="Protection Substation Tool", layout="wide")
st.title("Protection suite 綜研所變電所報告")

render_template_downloads()
st.divider()

source_uploads = st.file_uploader(
    "來源資料檔案（可多選）",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
)
target_upload = st.file_uploader(
    "目標檔案（建議 xlsm / xlsx）",
    type=["xlsx", "xlsm"],
    accept_multiple_files=False,
)

if st.button("開始轉換", type="primary", use_container_width=True):
    if not source_uploads:
        st.error("請至少上傳 1 個來源檔案。")
    elif target_upload is None:
        st.error("請上傳目標檔案。")
    else:
        with st.spinner("處理中，請稍候..."):
            try:
                result = process_workbook(source_uploads, target_upload)
            except Exception as exc:
                st.exception(exc)
            else:
                st.success("處理完成！")
                c1, c2, c3 = st.columns(3)
                c1.metric("成功匹配分頁", result["matched_count"])
                c2.metric("跳過分頁", result["skipped_count"])
                c3.metric("盤名數", result["plan_count"])

                st.subheader("來源解析紀錄")
                st.code("\n".join(result["parse_logs"]) or "無")

                st.subheader("寫入紀錄")
                st.code("\n".join(result["write_logs"]) or "無")

                out_name = f"updated_{Path(target_upload.name).name}"
                st.download_button(
                    "下載更新後檔案",
                    data=result["file_bytes"],
                    file_name=out_name,
                    mime="application/vnd.ms-excel",
                    use_container_width=True,
                )

