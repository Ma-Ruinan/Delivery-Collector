"""Evidence-driven collection executor; the agent supplies verified judgments.

No product-specific guessing, prompt submission, content scoring, or source
modification. See references/plan.md before using execute.
"""
from __future__ import annotations

import argparse
from contextlib import closing
import datetime as dt
import json
import os
import re
import shutil
import sqlite3
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from xlsx_io import CELL, digest, extend_sheet, merged_non_anchor, read_workbook, update_workbook, write_audit
from deadline import snapshot_task, timestamp
from excel_layout import resolve_slot


def emit(value):
    print(json.dumps(value, ensure_ascii=False, indent=2))


def norm_prompt(text):
    text = unicodedata.normalize("NFC", text).replace("\r\n", "\n").replace("\r", "\n").strip("\ufeff")
    return "\n".join(line.rstrip() for line in text.split("\n")).rstrip()


def valid_component(value):
    if not isinstance(value, str) or not value or value in (".", ".."):
        raise ValueError("文件名组件必须为非空文本")
    if re.search(r'[<>:"/\\|?*\x00-\x1f]', value) or value[-1] in ". ":
        raise ValueError("不安全的文件名组件: " + value)
    if value.split(".")[0].upper() in {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}:
        raise ValueError("保留文件名: " + value)
    return value


def inside(path, root):
    path, root = Path(path).resolve(), Path(root).resolve()
    if not path.is_relative_to(root):
        raise ValueError("路径不在指定范围内: " + str(path))
    return path


def absolute(value):
    path = Path(value)
    if not path.is_absolute():
        raise ValueError("计划内路径必须为当前环境绝对路径: " + str(value))
    return path.resolve()


def roster(root):
    import hashlib
    root = absolute(root)
    tasks = []
    for file in sorted(root.rglob("*")):
        if file.is_file() and file.name in ("问题描述.txt", "问题描述.md"):
            file = inside(file, root)
            text = file.read_text(encoding="utf-8-sig")
            tasks.append({"task_id": None, "question_dir": str(file.parent),
                          "question_path": str(file), "prompt": text,
                          "prompt_sha256": hashlib.sha256(norm_prompt(text).encode()).hexdigest(),
                          "nearby_files_not_classified": [str(p) for p in sorted(file.parent.rglob("*")) if p.is_file() and p != file]})
    return {"dataset_root": str(root), "tasks": tasks,
            "workbook_candidates": [str(p) for p in root.glob("*.xlsx") if "记录" in p.name]}


def sqlite_read(args):
    db = absolute(args.db)
    with closing(sqlite3.connect(db.as_uri() + "?mode=ro", uri=True, timeout=5)) as con:
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA query_only=ON")
        tables = {row["name"]: row["sql"] for row in con.execute("SELECT name, sql FROM sqlite_master WHERE type='table'")}
        if not args.table:
            return {"tables": tables}
        if args.table not in tables:
            raise ValueError("表不存在，应先查看当前 schema")
        quote = lambda s: '"' + s.replace('"', '""') + '"'
        columns = {r["name"] for r in con.execute("PRAGMA table_info(" + quote(args.table) + ")")}
        where, params = "", []
        if args.equals:
            column, value = args.equals
            if column not in columns:
                raise ValueError("筛选字段不存在")
            where, params = " WHERE " + quote(column) + "=?", [value]
        query = "SELECT * FROM " + quote(args.table) + where + " LIMIT ? OFFSET ?"
        rows = [dict(r) for r in con.execute(query, params + [args.limit, args.offset])]
        for row in rows:
            for key, value in row.items():
                if isinstance(value, bytes):
                    row[key] = {"blob_bytes": len(value), "not_decoded": True}
        return {"table": args.table, "offset": args.offset, "rows": rows,
                "possibly_more": len(rows) == args.limit}


def jsonl_read(args):
    rows = []
    with absolute(args.file).open(encoding="utf-8-sig") as stream:
        for i, line in enumerate(stream, 1):
            if i < args.start:
                continue
            if len(rows) >= args.limit:
                return {"rows": rows, "next_line": i}
            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                data = {"parse_error": True, "raw": line.rstrip()}
            rows.append({"line": i, "data": data})
    return {"rows": rows, "next_line": None}


def select_run(task, discovery_complete):
    if task.get("pending"):
        return "pending", None, str(task["pending"])
    attempts = task.get("attempts", [])
    if not attempts:
        return ("not_executed", None, "未发现匹配运行") if discovery_complete else ("pending", None, "发现范围不完整")
    if len(attempts) > 2:
        return "pending", None, "超过两次独立运行，需消歧"
    if any(a.get("status") not in {"success", "failed", "running", "unknown"} for a in attempts):
        raise ValueError("非法运行状态")
    # Array order must have been established from real submission evidence.
    if any(not a.get("session_id") or not a.get("submission_evidence") or not a.get("match_evidence") for a in attempts):
        return "pending", None, "缺少会话、提交顺序或题目匹配证据"
    try:
        submitted = [dt.datetime.fromisoformat(a["submitted_at"].replace("Z", "+00:00")) for a in attempts]
        if any(t.tzinfo is None for t in submitted) or submitted != sorted(submitted):
            return "pending", None, "提交时间缺时区或运行顺序未按实际提交排序"
        if len(submitted) == 2 and submitted[0] == submitted[1]:
            return "pending", None, "两次提交时间相同，需进一步确认顺序"
    except (KeyError, ValueError, TypeError):
        return "pending", None, "缺少可确认的实际提交时间"
    if len(attempts) == 2 and attempts[0]["session_id"] == attempts[1]["session_id"]:
        return "pending", None, "同会话无法直接作为两次独立重试，需核实"
    first = attempts[0]
    if first["status"] == "success":
        return "first_success", first, "第一次实际成功；第二次仅审计"
    if first["status"] in ("unknown", "running"):
        return "pending", None, "第一次未能作最终判断"
    if len(attempts) == 1:
        return "pending", None, "仅发现一次失败，未发现第二次记录"
    second = attempts[1]
    if second["status"] == "success":
        return "second_success", second, "第一次失败，第二次实际成功"
    if second["status"] == "failed":
        return "both_failed", second, "两次均未成功，以第二次正式输出DOCX作为交付"
    return "pending", None, "第二次执行中或无法确认"


def timing(run):
    if run.get("duration") and run.get("duration_evidence"):
        return str(run["duration"])
    return ""


def excel_values(decision, task, selected):
    runs = task["attempts"]
    first_time = timing(runs[0]) or "用时无法确认"
    first_reason = runs[0].get("reason", "")
    if decision == "first_success":
        note = "一次运行成功" + ("；仅发现一次运行" if len(runs) == 1 else "")
        if not timing(selected):
            note += "；用时无法确认"
        return timing(selected), note
    if decision == "second_success":
        note = f"第一次用时{first_time}，{first_reason}；第二次运行成功"
        if not timing(selected):
            note += "；第二次用时无法确认"
        return timing(selected), note
    return "两次测试均未成功", f"第一次用时{first_time}，{first_reason}；第二次用时{timing(selected) or '用时无法确认'}，{selected.get('reason', '')}"


def structural_check(path):
    if path.stat().st_size == 0:
        raise ValueError("文件为空")
    ext = path.suffix.lower()
    if ext in {".docx", ".pptx", ".xlsx"}:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad:
                raise ValueError("OOXML损坏: " + bad)
            required = {".docx": "word/document.xml", ".pptx": "ppt/presentation.xml", ".xlsx": "xl/workbook.xml"}[ext]
            ET.fromstring(z.read(required))
    if ext == ".pdf" and not path.read_bytes()[:1024].lstrip().startswith(b"%PDF-"):
        raise ValueError("非有效 PDF 头")


def prepare_files(prefix, task, selected, read_roots):
    if selected.get("completeness") != "complete" or not selected.get("validation_evidence"):
        raise ValueError("成功判断缺少完整性或实际读取证据")
    artifacts = selected.get("artifacts", [])
    if not artifacts:
        raise ValueError("成功运行没有正式交付清单")
    operations, targets = [], set()
    base = task["_question_dir"]
    for file in artifacts:
        if file.get("session_id") != selected["session_id"] or not file.get("evidence"):
            raise ValueError("文件来源会话或正式交付证据缺失")
        if selected.get("cutoff_at") and file.get("delivered_at") and timestamp(file["delivered_at"]) > timestamp(selected["cutoff_at"]):
            raise ValueError("文件在截止后才正式交付，不纳入本次")
        source = absolute(file["source"])
        if not any(source.is_relative_to(root) for root in read_roots):
            raise ValueError("文件超出本对象读取范围")
        if not source.is_file():
            raise ValueError("未找到实际交付文件: " + str(source))
        structural_check(source)
        sha = digest(source)
        if file.get("sha256") != sha:
            raise ValueError("源文件变化或哈希缺失")
        if file.get("project_name"):
            group = valid_component(file["project_name"])
            rel = Path(file["relative_path"])
            if rel.is_absolute() or any(p in ("..", ".") for p in rel.parts) or not rel.parts:
                raise ValueError("不安全项目相对路径")
            for p in rel.parts:
                valid_component(p)
            if rel.name != source.name:
                raise ValueError("项目内部文件名不能重命名")
            dest = base / f"{prefix}-{task['task_id']}-{group}" / rel
        else:
            name = valid_component(source.name)
            dest = base / f"{prefix}-{task['task_id']}-{name}"
        dest = inside(dest, base)
        if dest in targets:
            raise ValueError("同题交付目标重名，需区分正式版本")
        targets.add(dest)
        if dest.exists() and (not dest.is_file() or digest(dest) != sha):
            raise FileExistsError("同名不同内容，不覆盖: " + str(dest))
        operations.append({"source": str(source), "destination": str(dest), "sha256": sha,
                           "size": source.stat().st_size, "evidence": file["evidence"],
                           "delivery_channel": file.get("delivery_channel", "source_evidence"),
                           "session_id": selected["session_id"], "copy_status": "planned"})
    return operations


def formal_blocks(selected):
    cutoff_snapshot = (selected.get("timed_out") and selected.get("cutoff_state_evidence") and
                       selected.get("delivery_search_complete") is True)
    if selected.get("official_output_complete") is not True and not cutoff_snapshot:
        raise ValueError("第二次正式界面输出未完整取得")
    blocks = []
    for block in selected.get("official_output", []):
        if block.get("kind") in ("thinking", "reasoning"):
            continue
        if block.get("visible") is not True or not block.get("evidence"):
            raise ValueError("正式输出缺少用户可见性/来源证据")
        if block.get("kind") not in {"text", "image", "table"}:
            raise ValueError("无法保留的正式输出类型")
        blocks.append(block)
    if not blocks and not cutoff_snapshot:
        raise ValueError("未取得第二次正式输出，不能编造DOCX内容")
    return blocks


def make_failure(path, blocks, read_roots, observations=None, snapshot_complete=True):
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Microsoft YaHei"
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    for block in blocks:
        if block["kind"] == "text":
            # Keep source order and exact text; do not summarize or reword.
            for line in block["text"].split("\n"):
                doc.add_paragraph(line)
        elif block["kind"] == "image":
            source = absolute(block["source"])
            if not any(source.is_relative_to(r) for r in read_roots) or digest(source) != block["sha256"]:
                raise ValueError("正式图片来源/哈希不符")
            doc.add_picture(str(source), width=Inches(6))
        else:
            rows = block["rows"]
            if not rows or not rows[0] or len({len(r) for r in rows}) != 1:
                raise ValueError("正式表格结构不完整")
            table = doc.add_table(rows=0, cols=len(rows[0]))
            table.style = "Table Grid"
            for values in rows:
                cells = table.add_row().cells
                for cell, value in zip(cells, values):
                    cell.text = str(value)
            # Prevent any inherited fixed row height.
            for row in table.rows:
                for cell in row.cells:
                    borders = OxmlElement("w:tcBorders")
                    for edge in ("top", "left", "bottom", "right"):
                        node = OxmlElement("w:" + edge)
                        node.set(qn("w:val"), "single")
                        node.set(qn("w:sz"), "4")
                        node.set(qn("w:color"), "D9D9D9")
                        borders.append(node)
                    cell._tc.get_or_add_tcPr().append(borders)
    if observations:
        # This is the collector's evidence record, never an invented answer.
        doc.add_paragraph("收集核查记录")
        if not blocks:
            doc.add_paragraph("截至本轮收集截止时点，已核实未见正式界面输出。" if snapshot_complete else
                              "正式界面输出快照未完整取得，不能据此认定没有正式输出。")
        for observation in observations:
            doc.add_paragraph(observation)
    # Fixed metadata avoids treating a repeated identical export as a conflict.
    doc.core_properties.created = dt.datetime(2000, 1, 1)
    doc.core_properties.modified = dt.datetime(2000, 1, 1)
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""
    doc.save(path)
    with zipfile.ZipFile(path) as z:
        parts = [(info, z.read(info.filename)) for info in z.infolist()]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, payload in parts:
            info.date_time = (2000, 1, 1, 0, 0, 0)
            z.writestr(info, payload)
    structural_check(path)


def copy_one(operation):
    source, dest = Path(operation["source"]), Path(operation["destination"])
    if digest(source) != operation["sha256"]:
        raise ValueError("复制前源文件已变化")
    if dest.exists():
        if not dest.is_file() or digest(dest) != operation["sha256"]:
            raise FileExistsError("目标冲突")
        operation["copy_status"] = "already_present"
        return
    dest.parent.mkdir(parents=True, exist_ok=True)
    # Exclusive create rather than replacing another collector's output.
    with source.open("rb") as src, dest.open("xb") as out:
        shutil.copyfileobj(src, out)
    if digest(dest) != operation["sha256"] or digest(source) != operation["sha256"]:
        operation["copy_status"] = "verification_failed"
        raise ValueError("复制后哈希校验失败，未登记成功")
    operation["copy_status"] = "copied"


def task_cells(excel, sheets, task, values):
    sheet = sheets[excel["sheet"]]
    id_col = excel["task_id_column"]
    wanted = str(task.get("excel_task_id", task["task_id"]))
    rows = [CELL.fullmatch(addr)[2] for addr, c in sheet["cells"].items()
            if CELL.fullmatch(addr) and CELL.fullmatch(addr)[1] == id_col and c["value"] == wanted]
    if len(rows) != 1:
        raise ValueError("Excel题号行不唯一/未找到: " + wanted)
    result = []
    for column, value in zip((excel["time_column"], excel["note_column"]), values):
        address = column + rows[0]
        if merged_non_anchor(address, sheet["merged"]):
            raise ValueError("目标单元格是合并区域非锚点")
        old = sheet["cells"].get(address, {}).get("value")
        if old not in (None, "", value) and not excel.get("overwrite_existing"):
            raise ValueError("已有历史结果，调用未明确覆盖")
        if sheet["cells"].get(address, {}).get("formula") is not None:
            raise ValueError("目标单元格是公式")
        result.append({"sheet": excel["sheet"], "cell": address, "expected": old, "value": value})
    return result


def save_json(path, value):
    data = json.dumps(value, ensure_ascii=False, indent=2)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(data, encoding="utf-8")
    os.replace(tmp, path)


def execute(plan, apply=False):
    if plan.get("schema_version") not in (1, 2):
        raise ValueError("未知计划 schema_version")
    cutoff_at = plan.get("cutoff_at")
    deadline_mode = plan.get("deadline_mode", bool(cutoff_at))
    if deadline_mode:
        if not cutoff_at:
            raise ValueError("定时收集须在开始发现前固定cutoff_at，不能在遍历结束后另取时间")
        timestamp(cutoff_at)
    root = absolute(plan["dataset_root"])
    if not root.is_dir():
        raise ValueError("数据集不存在")
    object_info = plan["object"]
    if not all(object_info.get(k) for k in ("name", "version", "model", "prefix")):
        raise ValueError("必须明确一个对象的名称、版本、模型和前缀")
    prefix = valid_component(object_info["prefix"])
    roots = [absolute(p) for p in plan["read_roots"]]
    if not roots:
        raise ValueError("必须明确本对象可读取项目范围")
    batch = valid_component(plan["batch_id"])
    if plan.get("discovery_complete") is True and not plan.get("discovery_evidence"):
        raise ValueError("完整发现范围须有证据，不能仅凭无文件判未执行")
    stamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
    report = inside(root / "收集核查记录" / f"{prefix}-{batch}-{stamp}", root)
    report.mkdir(parents=True, exist_ok=False)
    audit = {"schema_version": 2, "object": object_info, "batch_id": batch,
             "cutoff_at": cutoff_at, "deadline_mode": deadline_mode,
             "unattended": plan.get("unattended", True),
             "applied": apply, "plan": plan, "tasks": [], "files": [], "issues": [],
             "excel": {"status": "not_updated"}, "human_verified": False}
    record = report / "收集发现与执行记录.json"
    save_json(record, audit)
    excel, sheets, excel_problem = plan.get("excel"), None, None
    if excel:
        try:
            excel_path = inside(absolute(excel["path"]), root)
            if digest(excel_path) != excel["sha256"]:
                raise ValueError("结果表自读取后发生变化")
            sheets = read_workbook(excel_path)
            if "object_slot" in excel:
                excel = resolve_slot(excel, sheets, object_info)
                # Preflight extension structure before copying any files.
                with zipfile.ZipFile(excel_path) as z:
                    preview_parts = {}
                    for extension in excel.get("extensions", []):
                        part = sheets[extension["sheet"]]["part"]
                        preview_parts[part] = extend_sheet(preview_parts.get(part, z.read(part)), extension["source_columns"], extension["target_columns"])
                audit["resolved_excel"] = excel
            if not excel.get("target_evidence"):
                raise ValueError("未确认本对象 Excel 槽位")
            cols = [excel[k] for k in ("task_id_column", "time_column", "note_column")]
            if len(set(cols)) != 3 or any(not re.fullmatch("[A-Z]+", c) for c in cols):
                raise ValueError("Excel题号/用时/说明列非法或重叠")
            if excel["sheet"] not in sheets:
                raise ValueError("目标工作表不存在")
            for header in excel.get("headers", []):
                allowed = excel["allowed_header_columns"] if "object_slot" in excel else cols[1:]
                if not CELL.fullmatch(header["cell"]) or CELL.fullmatch(header["cell"])[1] not in allowed:
                    raise ValueError("表头不属于指定对象列")
                old = sheets[excel["sheet"]]["cells"].get(header["cell"], {}).get("value")
                if (old or "") != (header["expected"] or ""):
                    raise ValueError("表头自读取后变化")
        except Exception as e:
            excel_problem = str(e)
            audit["issues"].append({"task_id": "", "issue": "Excel: " + excel_problem})
    else:
        audit["issues"].append({"task_id": "", "issue": "未定位结果表，文件收集继续；Excel未填写"})
    changes, seen = [], set()

    def schedule_excel(task, entry, decision, selected):
        if sheets is None or excel_problem:
            return
        try:
            cells = task_cells(excel, sheets, task, excel_values(decision, task, selected))
            if {(c['sheet'], c['cell']) for c in cells} & {(c['sheet'], c['cell']) for c in changes}:
                raise ValueError("多个题号映射到同一Excel记录行，需确认编号映射")
            entry["planned_cells"] = cells
            changes.extend(cells)
            entry["excel_status"] = "planned"
        except Exception as e:
            audit["issues"].append({"task_id": task["task_id"], "issue": "Excel: " + str(e)})

    for original in plan["tasks"]:
        task = snapshot_task(original, cutoff_at) if deadline_mode else dict(original)
        entry = {"task_id": task.get("task_id"), "attempts": task.get("attempts", []),
                 "requirements": task.get("requirements", []), "excel_status": "not_updated", "human_check": ""}
        audit["tasks"].append(entry)
        try:
            task_id = valid_component(task["task_id"])
            if task_id in seen:
                raise ValueError("计划存在重复题号")
            seen.add(task_id)
            base = inside(absolute(task["question_dir"]), root)
            if base == root or not base.is_dir():
                raise ValueError("应使用已有问题文件夹，不是数据集根")
            task["_question_dir"] = base
            decision, selected, reason = select_run(task, plan.get("discovery_complete") is True)
            entry.update(decision=decision, selection_reason=reason,
                         selected_session=selected.get("session_id") if selected else None)
            if decision == "not_executed":
                continue
            if decision == "pending":
                raise ValueError(reason)
            if decision != "first_success" and any(r.get("status") == "failed" and not r.get("reason") for r in task["attempts"]):
                raise ValueError("失败运行缺少实际原因")
            if decision == "both_failed":
                # Record the confirmed execution outcome independently of DOCX
                # extraction, rendering and copying; success keeps its old gate.
                schedule_excel(task, entry, decision, selected)
                blocks = formal_blocks(selected)
                snapshot_complete = selected.get("output_snapshot_complete", selected.get("official_output_complete")) is True
                observations = list(selected.get("collection_observations") or [])
                if selected.get("timed_out") and not snapshot_complete:
                    warning = "正式界面输出快照未完整取得；仅保存已核实的可见内容及收集事实，待人工核查"
                    observations.append(warning)
                    entry["output_snapshot_status"] = "partial"
                    audit["issues"].append({"task_id": task_id, "issue": warning})
                    if entry.get("planned_cells"):
                        entry["planned_cells"][1]["value"] += "；界面输出快照未完整取得（见核查清单）"
                name = f"{prefix}-{task_id}-未成功任务界面输出.docx"
                target = inside(base / name, base)
                # Staged inside audit directory; never becomes an invented answer.
                staged = report / name
                make_failure(staged, blocks, roots, observations, snapshot_complete)
                sha = digest(staged)
                if target.exists() and (not target.is_file() or digest(target) != sha):
                    raise FileExistsError("失败输出目标同名不同内容")
                operations = [{"source": str(staged), "destination": str(target), "sha256": sha,
                    "size": staged.stat().st_size, "session_id": selected["session_id"],
                    "evidence": [b["evidence"] for b in blocks] + ([selected["cutoff_state_evidence"]] if selected.get("timed_out") else []), "copy_status": "planned",
                    "generated_from_official_output": True}]
            else:
                operations = prepare_files(prefix, task, selected, roots)
            for op in operations:
                op["task_id"] = task_id
                audit["files"].append(op)
            if apply:
                for op in operations:
                    copy_one(op)
                    save_json(record, audit)
            entry["collection_status"] = "collected" if apply else "planned"
            if decision != "both_failed":
                schedule_excel(task, entry, decision, selected)
        except Exception as e:
            entry["collection_status"] = "pending"
            entry["issue"] = str(e)
            if entry.get("decision") == "both_failed" and entry.get("planned_cells"):
                entry["planned_cells"][1]["value"] += "；输出快照保存失败（见核查清单）"
            audit["issues"].append({"task_id": task.get("task_id", ""), "issue": str(e)})
        finally:
            save_json(record, audit)
    if apply and changes and not excel_problem:
        try:
            if digest(excel_path) != excel["sha256"]:
                raise ValueError("收集期间结果表被修改，不覆盖")
            header_changes = [{"sheet": excel["sheet"], **h} for h in excel.get("headers", [])]
            audit["excel"] = update_workbook(excel_path, changes + header_changes, report / (excel_path.stem + "-更新前.xlsx"), excel.get("extensions"))
            for entry in audit["tasks"]:
                if entry["excel_status"] == "planned":
                    entry["excel_status"] = "updated"
        except Exception as e:
            audit["issues"].append({"task_id": "", "issue": "Excel写入: " + str(e)})
            audit["excel"] = {"status": "pending", "issue": str(e)}
    save_json(record, audit)
    task_rows = [["题号", "选取状态", "采用会话", "选择依据", "交付要求", "运行发现（含额外运行）", "文件收集", "Excel登记", "人工核查"]]
    for e in audit["tasks"]:
        task_rows.append([e["task_id"], e.get("decision", "pending"), e.get("selected_session"),
            e.get("selection_reason", e.get("issue", "")), json.dumps(e["requirements"], ensure_ascii=False),
            json.dumps(e["attempts"], ensure_ascii=False), e.get("collection_status", "未执行"), e["excel_status"], ""])
    file_rows = [["题号", "来源会话", "源文件", "目标文件", "字节数", "SHA256", "正式来源证据", "复制状态", "人工核查"]]
    for f in audit["files"]:
        file_rows.append([f["task_id"], f["session_id"], f["source"], f["destination"], f["size"], f["sha256"], str(f["evidence"]), f["copy_status"], ""])
    issues = [["题号", "待确认事项", "人工处理结果", "核查人", "核查日期"]]
    issues.extend([i["task_id"], i["issue"], "", "", ""] for i in audit["issues"])
    issues.append(["全部", "自动核查不替代人工核查；请打开交付并核对来源、格式、完整性与Excel。", "", "", ""])
    checklist = report / "交付收集核查清单.xlsx"
    write_audit(checklist, {"任务与运行发现": task_rows, "交付文件明细": file_rows, "待确认与人工核查": issues})
    return {"report_dir": str(report), "audit_json": str(record), "checklist": str(checklist),
            "applied": apply, "tasks": len(audit["tasks"]), "copied": sum(f["copy_status"] == "copied" for f in audit["files"]),
            "already_present": sum(f["copy_status"] == "already_present" for f in audit["files"]),
            "excel": audit["excel"], "issues": audit["issues"], "human_verified": False}


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    p = commands.add_parser("roster"); p.add_argument("dataset")
    p = commands.add_parser("sqlite"); p.add_argument("db"); p.add_argument("--table")
    p.add_argument("--equals", nargs=2, metavar=("COLUMN", "VALUE")); p.add_argument("--limit", type=int, default=100); p.add_argument("--offset", type=int, default=0)
    p = commands.add_parser("jsonl"); p.add_argument("file"); p.add_argument("--start", type=int, default=1); p.add_argument("--limit", type=int, default=50)
    p = commands.add_parser("workbook"); p.add_argument("file")
    commands.add_parser("cutoff", help="在收集启动时固定UTC截止时间，预检/执行复用该值")
    p = commands.add_parser("layout", help="按指定对象序号只读解析或规划Excel列组")
    p.add_argument("file"); p.add_argument("--slot", type=int, required=True)
    p.add_argument("--name", required=True); p.add_argument("--version", required=True)
    p.add_argument("--model", required=True); p.add_argument("--sheet")
    p = commands.add_parser("execute"); p.add_argument("plan"); p.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    if hasattr(args, "limit") and not 1 <= args.limit <= 1000:
        parser.error("limit应为1–1000")
    if getattr(args, "offset", 0) < 0 or getattr(args, "start", 1) < 1:
        parser.error("非法分页范围")
    try:
        if args.command == "roster":
            result = roster(args.dataset)
        elif args.command == "sqlite":
            result = sqlite_read(args)
        elif args.command == "jsonl":
            result = jsonl_read(args)
        elif args.command == "workbook":
            result = read_workbook(args.file)
        elif args.command == "cutoff":
            result = {"cutoff_at": dt.datetime.now(dt.timezone.utc).isoformat()}
        elif args.command == "layout":
            config = {"path": str(absolute(args.file)), "sha256": digest(args.file), "object_slot": args.slot}
            if args.sheet:
                config["sheet"] = args.sheet
            result = resolve_slot(config, read_workbook(args.file), {"name": args.name, "version": args.version, "model": args.model})
        else:
            result = execute(json.loads(Path(args.plan).read_text(encoding="utf-8-sig")), args.apply)
        emit(result)
    except Exception as e:
        parser.exit(2, str(e) + "\n")


if __name__ == "__main__":
    main()
