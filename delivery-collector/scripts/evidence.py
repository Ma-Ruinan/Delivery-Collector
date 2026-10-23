"""Validate current-session membership, complete delivery lists and time bases."""
from __future__ import annotations

import copy
from contextlib import closing
import datetime as dt
import json
from pathlib import Path
import sqlite3

from deadline import elapsed_label, timestamp


def eligible_task(original, strict=False):
    """Remove deleted/non-visible sessions before counting independent runs."""
    task = copy.deepcopy(original)
    attempts, excluded = [], task.setdefault("excluded_candidates", [])
    for run in task.get("attempts", []):
        presence = run.get("client_presence")
        if run.get("deleted_at") or presence in {"deleted", "not_visible"}:
            excluded.append({"session_id": run.get("session_id"), "reason": "客户端已删除或不在本次有效任务列表，不计运行次数",
                             "evidence": run.get("presence_evidence") or {"deleted_at": run.get("deleted_at")}})
            continue
        if (strict or presence is not None) and not (presence == "visible" and run.get("presence_evidence")):
            run["observed_status"] = run.get("status")
            run["status"] = "unknown"
            run["reason"] = "客户端有效任务列表归属未核实"
            # Deadline processing must not revive an unverified candidate.
            run["status_at_cutoff"] = "unknown"
            run["client_presence_verified"] = False
        elif presence == "visible" and run.get("presence_evidence"):
            run["client_presence_verified"] = True
        attempts.append(run)
    task["attempts"] = attempts
    return task


def normalize_duration(run, strict=False):
    """Normalize one attempt, retaining segment evidence and honest fallbacks."""
    if run.get("timed_out"):
        return
    source = run.get("duration_source")
    if not strict and source is None:
        return  # Existing schema 1/2 plans remain readable.
    try:
        if source == "client_total":
            if not run.get("duration") or not run.get("duration_evidence"):
                raise ValueError("客户端完整用时缺少数值或来源")
            run.pop("duration_issue", None)
            run.pop("duration_note", None)
            return  # A cumulative client value must not be summed a second time.
        if source in {"client_segments", "execution_segments"}:
            segments = run.get("duration_segments", [])
            if not isinstance(segments, list) or not segments or run.get("duration_segments_complete") is not True:
                raise ValueError("分段用时未覆盖完整同次运行")
            submitted, completed = timestamp(run["submitted_at"]), timestamp(run["completed_at"])
            previous, total, labels = None, 0, []
            for part in segments:
                if not isinstance(part, dict):
                    raise ValueError("分段用时记录结构无效")
                if part.get("session_id") != run["session_id"] or not part.get("evidence"):
                    raise ValueError("分段用时缺少同会话来源证据")
                start, end = timestamp(part["started_at"]), timestamp(part["completed_at"])
                if start < submitted or end > completed or end < start or (previous is not None and start < previous):
                    raise ValueError("分段用时越界、重叠或顺序错误，不能重复累计")
                seconds = (end - start).total_seconds()
                total += seconds
                labels.append(elapsed_label(seconds))
                previous = end
            run["duration"] = elapsed_label(total)
            run["duration_evidence"] = {"source": source, "segments": segments}
            run["duration_note"] = f"按完整同次运行{len(segments)}段累计（{'＋'.join(labels)}）"
        elif source == "elapsed_fallback":
            if not run.get("duration_evidence"):
                raise ValueError("经过时长缺少实际提交/执行与结束证据")
            start = timestamp(run.get("started_at") or run["submitted_at"])
            end = timestamp(run["completed_at"])
            if start < timestamp(run["submitted_at"]) or end < start:
                raise ValueError("经过时长起止顺序错误")
            seconds = (end - start).total_seconds()
            pauses = run.get("pause_seconds", 0)
            if isinstance(pauses, bool) or not isinstance(pauses, (int, float)) or not 0 <= pauses <= seconds or (pauses and not run.get("pause_evidence")):
                raise ValueError("等待/暂停扣除范围或证据不符")
            run["duration"] = elapsed_label(seconds - pauses) + "（提交/执行至完成经过时长）"
            run["duration_note"] = "客户端完整用时未取得，采用有证据的经过时长"
        else:
            raise ValueError("未明确用时来源或完整统计范围")
        run.pop("duration_issue", None)
    except (KeyError, TypeError, ValueError) as error:
        run.pop("duration", None)
        run.pop("duration_evidence", None)
        run.pop("duration_note", None)
        run["duration_issue"] = str(error)


def check_deliveries(run, strict=False):
    """Require every formally offered file, regardless of its filename."""
    manifest = run.get("formal_deliveries")
    if manifest is None and not strict:
        return
    if run.get("delivery_search_complete") is not True or not isinstance(manifest, list) or not manifest:
        raise ValueError("缺少全部正式交付核查清单，不能确认无漏收")
    actual = {Path(f["source"]).resolve() for f in run.get("artifacts", [])}
    for item in manifest:
        if item.get("session_id") != run["session_id"] or not item.get("evidence"):
            raise ValueError("正式交付清单缺少同会话来源证据")
        source = Path(item["source"])
        if not source.is_absolute():
            raise ValueError("正式交付清单须给出解析后的绝对路径")
        if run.get("cutoff_at") and item.get("delivered_at") and timestamp(item["delivered_at"]) > timestamp(run["cutoff_at"]):
            raise ValueError("正式清单包含截止后才交付的文件：" + source.name)
        if source.resolve() not in actual:
            raise ValueError("正式提供文件未纳入收集清单：" + source.name)


def read_deliveries(product, path, session_id, cutoff_at=None):
    """Read ALL successful component results, not just the first tool call.

    This does not find tasks, verify file content, or extract prose paths.
    Unknown result shapes fail visibly rather than returning an empty manifest.
    """
    path = Path(path).resolve()
    events = []
    if product == "workbuddy":
        for line, raw in enumerate(path.read_text(encoding="utf-8-sig").splitlines(), 1):
            if not raw.strip():
                continue
            row = json.loads(raw)
            if row.get("sessionId") not in (None, session_id):
                continue
            if row.get("type") != "function_call_result" or row.get("name") != "present_files" or row.get("status") != "completed":
                continue
            for block in row.get("output", []):
                if not block.get("text"):
                    continue
                value = json.loads(block["text"])
                if value.get("type") != "present_files_result" or not isinstance(value.get("files"), list):
                    raise ValueError("present_files结果结构未识别，须核查原始记录")
                events.append((value["files"], row.get("cwd"), row.get("timestamp"), f"{path}:{line}；结果{row.get('id')}"))
    elif product == "mobilework":
        with closing(sqlite3.connect(path.as_uri() + "?mode=ro", uri=True)) as db:
            session = db.execute("SELECT directory FROM session WHERE id=?", (session_id,)).fetchone()
            if session is None:
                raise ValueError("会话不在当前应用库中；不能从残留目录恢复为有效任务")
            for part_id, raw in db.execute("SELECT id,data FROM part WHERE session_id=? ORDER BY time_created,id", (session_id,)):
                part = json.loads(raw)
                state = part.get("state", {})
                if part.get("tool") != "result_file" or state.get("status") != "completed":
                    continue
                value = state.get("output")
                value = json.loads(value) if isinstance(value, str) else value
                if isinstance(value, dict) and value.get("ok") is False:
                    continue
                if not isinstance(value, dict) or value.get("ok") is not True or not isinstance(value.get("delivered"), list):
                    raise ValueError("result_file结果结构或交付成功状态未核实")
                events.append((value["delivered"], session[0], state.get("time", {}).get("end"), f"{path}；part:{part_id}"))
    else:
        raise ValueError("未知交付记录格式")
    manifest = {}
    for files, directory, delivered, evidence in events:
        if isinstance(delivered, (int, float)):
            delivered = dt.datetime.fromtimestamp(delivered / 1000, dt.timezone.utc).isoformat()
        if cutoff_at:
            if not delivered:
                raise ValueError("正式交付缺少截止归属时间证据")
            if timestamp(delivered) > timestamp(cutoff_at):
                continue
        for source in files:
            source = Path(source)
            if not source.is_absolute():
                if not directory or not Path(directory).is_absolute():
                    raise ValueError("相对交付路径缺少实际会话目录")
                source = Path(directory) / source
            key = str(source.resolve())
            item = manifest.setdefault(key, {"source": key, "session_id": session_id, "evidence": [],
                                             "delivered_at": delivered})
            item["evidence"].append(evidence)
    return {"session_id": session_id, "successful_delivery_events": len(events), "formal_deliveries": list(manifest.values()),
            "scope": "全部成功组件结果；仍须结合正式回答中的最终路径核查", "cutoff_at": cutoff_at}
