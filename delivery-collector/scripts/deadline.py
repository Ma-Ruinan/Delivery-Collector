"""Classify deadline snapshots without stopping or waiting on source sessions."""
from __future__ import annotations

import copy
import datetime as dt


def timestamp(value):
    if not isinstance(value, str):
        raise ValueError("时间必须是包含时区的文本")
    parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        raise ValueError("时间必须包含时区")
    return parsed


def elapsed_label(seconds):
    seconds = int(seconds)
    hours, rem = divmod(seconds, 3600)
    minutes, seconds = divmod(rem, 60)
    return (f"{hours}h" if hours else "") + (f"{minutes}m" if minutes or hours else "") + f"{seconds}s"


def at_deadline(original, cutoff_at):
    """Use source-backed state AT cutoff, never the executor's later wall clock.

    A running session can have final deliverables without a UI download widget.
    Missing discovery/state evidence stays unknown, not an invented timeout.
    """
    run = copy.deepcopy(original)
    cutoff = timestamp(cutoff_at)
    run["cutoff_at"] = cutoff_at
    run["observed_status"] = run.get("status")
    if run.get("client_presence_verified") is False:
        run["status"] = "unknown"
        run["reason"] = "客户端有效任务列表归属未核实"
        return run
    try:
        submitted = timestamp(run["submitted_at"])
        if submitted > cutoff:
            raise ValueError("提交晚于本批截止时间")
        state = run.get("status_at_cutoff", run["status"])
        if run.get("completed_at") and timestamp(run["completed_at"]) > cutoff:
            state = "running"
        run["status"] = state
        if state != "running":
            if state == "success" and run.get("delivery_at") and timestamp(run["delivery_at"]) > cutoff:
                raise ValueError("正式交付产生于截止后，不属于本轮成功结果")
            return run
        if not run.get("cutoff_state_evidence") or run.get("delivery_search_complete") is not True:
            raise ValueError("截止时状态或正式路径/交付核查证据不足")
        if run.get("completeness") == "complete":
            if not (run.get("artifacts") and run.get("validation_evidence") and run.get("delivery_at")):
                raise ValueError("完整交付缺少实际文件、核查或交付时间证据")
            if timestamp(run["delivery_at"]) > cutoff:
                raise ValueError("完整交付产生于截止后，不纳入本次")
            run["status"] = "success"
            run["reason"] = "截止前全部正式交付已实际取得；界面仍运行不影响已完成交付"
            return run
        if run.get("completeness") not in ("none", "incomplete"):
            raise ValueError("交付完整性尚未查清，不能直接判超时")
        # The failure is already established. Timing/export gaps must not turn
        # a confirmed timeout back into an unknown execution outcome.
        run["timed_out"] = True
        run["status"] = "failed"
        suffix = "截止时未见实际正式交付" if run["completeness"] == "none" else "截止时仅有部分交付，必需成果未齐"
        run["reason"] = "运行时间过长，判定为执行失败；" + suffix
        run.pop("duration_issue", None)
        run.pop("elapsed_seconds", None)
        try:
            fresh_client_time = bool(run.get("duration_evidence") and run.get("duration") and
                                     timestamp(run["duration_as_of"]) == cutoff)
        except (KeyError, ValueError, TypeError):
            fresh_client_time = False
        if not fresh_client_time:
            run.pop("duration", None)
            run.pop("duration_evidence", None)
            try:
                started = timestamp(run.get("started_at") or run["submitted_at"])
                if started < submitted or started > cutoff:
                    raise ValueError("实际执行开始时间不在提交与截止之间")
                seconds = int((cutoff - started).total_seconds())
                basis = "截至收集时经过时长"
                pauses = run.get("pause_seconds", 0)
                if pauses:
                    if not isinstance(pauses, (int, float)) or not 0 <= pauses <= seconds or not run.get("pause_evidence"):
                        raise ValueError("暂停时间缺少可信证据或范围不符")
                    seconds -= int(pauses)
                    basis = "截至收集时经过时长，已扣除可确认暂停"
                run["duration"] = f"{elapsed_label(seconds)}（{basis}）"
                run["duration_evidence"] = {"start": started.isoformat(), "cutoff": cutoff_at,
                                            "basis": basis, "submission_evidence": run.get("submission_evidence")}
                run["elapsed_seconds"] = seconds
            except (KeyError, ValueError, TypeError) as error:
                run["duration_issue"] = str(error)
                run["reason"] += "；用时无法确认：" + str(error)
        run["collection_observations"] = [
            f"收集截止时间：{cutoff_at}。会话：{run['session_id']}。",
            f"截止时会话仍在运行，{suffix}。本轮测评按运行超时判定失败，未终止原会话。",
            f"截至收集时耗时：{run.get('duration') or '用时无法确认'}。"
        ]
        return run
    except (KeyError, ValueError, TypeError) as error:
        run["status"] = "unknown"
        run["reason"] = str(error)
        return run


def snapshot_task(task, cutoff_at):
    result = copy.deepcopy(task)
    result["attempts"] = [at_deadline(run, cutoff_at) for run in task.get("attempts", [])]
    return result
