"""Resolve an explicitly selected object slot; never ask for column letters.

Supports ordinary paired time/note columns and two-level merged headers. The
agent can supply inspected explicit mappings for other layouts.
"""
from __future__ import annotations

import copy
import re

from xlsx_io import CELL, col_name, col_number


def compact(value):
    return re.sub(r"\s+", "", value or "")


def object_groups(sheet):
    groups = {}
    cells = sheet["cells"]
    for address, cell in cells.items():
        text = compact(cell["value"])
        matched = re.search(r"测试对象\s*([1-9][0-9]*)", text)
        if not matched:
            continue
        col, row = CELL.fullmatch(address).groups()
        index, start = int(matched[1]), col_number(col)
        time_cell, note_cell = None, None
        for r in range(int(row), int(row) + 4):
            left, right = col + str(r), col_name(start + 1) + str(r)
            ltext = compact(cells.get(left, {}).get("value"))
            rtext = compact(cells.get(right, {}).get("value"))
            if ("用时" in ltext or "耗时" in ltext) and ("说明" in rtext or "完成情况" in rtext):
                time_cell, note_cell = left, right
                break
        if time_cell:
            candidate = {"slot": index, "anchor": address, "time_cell": time_cell,
                         "note_cell": note_cell, "time_column": col,
                         "note_column": col_name(start + 1)}
            if index in groups and groups[index] != candidate:
                raise ValueError("同一工作表对象编号重复，布局不唯一")
            groups[index] = candidate
    # Single-object sheets may omit the numeric suffix in their header.
    if not groups:
        candidates = []
        for address, cell in cells.items():
            if "测试对象" in compact(cell["value"]):
                col, row = CELL.fullmatch(address).groups()
                note_col = col_name(col_number(col) + 1)
                for r in range(int(row), int(row) + 4):
                    label = compact(cells.get(col + str(r), {}).get("value"))
                    note = compact(cells.get(note_col + str(r), {}).get("value"))
                    if ("用时" in label or "耗时" in label) and ("说明" in note or "完成情况" in note):
                        candidates.append({"slot": 1, "anchor": address, "time_cell": col + str(r),
                            "note_cell": note_col + str(r), "time_column": col, "note_column": note_col})
                        break
        if len(candidates) == 1:
            groups[1] = candidates[0]
    return groups


def metadata_text(info, slot, combined):
    name = info.get("display_name") or info["name"]
    version = info.get("version", "")
    if not info.get("display_name") and version and version not in name:
        name += "-" + version
    return f"测试对象{slot}（名称/版本：{name}；大模型名称：{info['model']}）" + ("用时" if combined else "")


def resolve_slot(excel, sheets, info):
    """Return resolved configuration and style-only extension instructions."""
    result = copy.deepcopy(excel)
    slot = result.get("object_slot")
    if isinstance(slot, bool) or not isinstance(slot, int) or not 1 <= slot <= 3:
        raise ValueError("object_slot应为调用明确指定的测试对象1、2或3")
    candidates = []
    for name, sheet in sheets.items():
        if result.get("sheet") and name != result["sheet"]:
            continue
        groups = object_groups(sheet)
        header_rows = {CELL.fullmatch(g[key])[2] for g in groups.values() for key in ("anchor", "time_cell", "note_cell")}
        # Instructions/legends below the records can also contain "任务编号".
        # Use the actual object-header band, not a workbook-wide text match.
        id_cells = [a for a, c in sheet["cells"].items()
                    if compact(c["value"]) in {"任务编号", "题号", "问题编号"} and CELL.fullmatch(a)[2] in header_rows]
        if len(id_cells) != 1:
            continue
        if groups:
            candidates.append((name, sheet, id_cells[0], groups))
    if len(candidates) != 1:
        raise ValueError("记录工作表或对象列组无法唯一定位；无人值守时记录异常，不猜测写入")
    name, sheet, id_address, groups = candidates[0]
    if sheet["protected"]:
        raise ValueError("记录表受保护，不能绕过保护扩展或写入")
    result["sheet"] = name
    result["task_id_column"] = CELL.fullmatch(id_address)[1]
    result["target_evidence"] = excel.get("target_evidence") or f"本次调用明确指定测试对象{slot}，按实际表头定位"
    result.setdefault("overwrite_existing", True)
    extensions, extra_headers = [], []
    if slot not in groups:
        if sorted(groups) != list(range(1, max(groups) + 1)) or slot < max(groups):
            raise ValueError("现有对象编号不连续，不能自动在右侧创建中间槽位")
        source = groups[max(groups)]
        max_col = max(col_number(CELL.fullmatch(a)[1]) for a in sheet["cells"])
        for region in sheet["merged"]:
            max_col = max(max_col, col_number(CELL.fullmatch(region.split(":")[-1])[1]))
        for index in range(max(groups) + 1, slot + 1):
            new_time, new_note = col_name(max_col + 1), col_name(max_col + 2)
            max_col += 2
            anchor_row = CELL.fullmatch(source["anchor"])[2]
            time_row = CELL.fullmatch(source["time_cell"])[2]
            note_row = CELL.fullmatch(source["note_cell"])[2]
            group = {"slot": index, "anchor": new_time + anchor_row, "time_cell": new_time + time_row,
                     "note_cell": new_note + note_row, "time_column": new_time, "note_column": new_note}
            groups[index] = group
            extensions.append({"sheet": name, "source_columns": [source["time_column"], source["note_column"]],
                               "target_columns": [new_time, new_note]})
            combined = group["anchor"] == group["time_cell"]
            placeholder = f"测试对象{index}（名称/版本：；大模型名称：）" + ("用时" if combined else "")
            extra_headers.append({"cell": group["anchor"], "expected": None, "value": placeholder})
            if not combined:
                extra_headers.append({"cell": group["time_cell"], "expected": None, "value": "用时"})
            extra_headers.append({"cell": group["note_cell"], "expected": None, "value": "任务完成情况、说明"})
    group = groups[slot]
    result.update(time_column=group["time_column"], note_column=group["note_column"])
    headers = {h["cell"]: h for h in extra_headers}
    anchor = group["anchor"]
    headers[anchor] = {"cell": anchor, "expected": sheet["cells"].get(anchor, {}).get("value"),
                       "value": metadata_text(info, slot, anchor == group["time_cell"])}
    result["headers"] = list(headers.values())
    # Blank intermediate columns are layout only, not another collected object.
    result["allowed_header_columns"] = sorted({CELL.fullmatch(h["cell"])[1] for h in result["headers"]})
    result["extensions"] = extensions
    return result
