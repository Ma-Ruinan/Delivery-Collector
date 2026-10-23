"""Coverage, current-client membership, segmented timing and centered writes."""
import copy
import json
from pathlib import Path
import sqlite3
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

import test_collector as fixtures
from evidence import eligible_task, normalize_duration, read_deliveries

c, x = fixtures.c, fixtures.x


class EvidenceTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.FileAndWorkbookTests()
        self.fixture.setUp()

    def tearDown(self):
        self.fixture.tearDown()

    def plan(self):
        plan = self.fixture.plan()
        plan["schema_version"] = 3
        for run in plan["tasks"][0]["attempts"]:
            run.update(client_presence="visible", presence_evidence="当前精准工作空间的完整有效会话列表", duration_source="client_total")
        first = plan["tasks"][0]["attempts"][0]
        first.update(delivery_search_complete=True, formal_deliveries=copy.deepcopy(first["artifacts"]))
        return plan

    def audit(self, result):
        return json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))

    def test_formally_offered_plan_attachment_is_collected(self):
        plan = self.plan()
        run = plan["tasks"][0]["attempts"][0]
        source = Path(plan["read_roots"][0]) / "research_plan_attachment.md"
        source.write_text("正式提供的研究计划", encoding="utf-8")
        item = {"source": str(source), "session_id": "s1", "sha256": x.digest(source), "evidence": "成功present_files结果中的第三个附件"}
        run["artifacts"].append(item)
        run["formal_deliveries"].append(copy.deepcopy(item))
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 2, result)
        self.assertTrue((Path(plan["tasks"][0]["question_dir"]) / "workbuddy5.5.6-1.1-research_plan_attachment.md").is_file())

    def test_missing_formal_attachment_prevents_silent_success(self):
        plan = self.plan()
        run = plan["tasks"][0]["attempts"][0]
        run["formal_deliveries"].append({"source": str(Path(plan["read_roots"][0]) / "research_plan.md"), "session_id": "s1", "evidence": "正式交付组件第三个文件"})
        before = x.digest(plan["excel"]["path"])
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 0)
        self.assertEqual(x.digest(plan["excel"]["path"]), before)
        self.assertTrue(any("research_plan.md" in i["issue"] for i in result["issues"]))

    def test_schema3_requires_a_complete_manifest(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"][0].pop("formal_deliveries")
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 0)
        self.assertTrue(any("核查清单" in i["issue"] for i in result["issues"]))

    def test_deleted_session_removed_before_run_count_and_selection(self):
        plan = self.plan()
        old = copy.deepcopy(plan["tasks"][0]["attempts"][0])
        old.update(session_id="deleted-old", submitted_at="2026-10-14T08:00:00+08:00", deleted_at=1791936060000)
        plan["tasks"][0]["attempts"].insert(0, old)
        result = c.execute(plan, True)
        task = self.audit(result)["tasks"][0]
        self.assertEqual(task["decision"], "first_success")
        self.assertEqual(task["selected_session"], "s1")
        self.assertEqual(len(task["attempts"]), 2)
        self.assertEqual(task["excluded_candidates"][0]["session_id"], "deleted-old")

    def test_only_deleted_sessions_preserve_history(self):
        plan = self.plan()
        for run in plan["tasks"][0]["attempts"]:
            run["client_presence"] = "deleted"
        before = x.digest(plan["excel"]["path"])
        result = c.execute(plan, True)
        self.assertEqual(x.digest(plan["excel"]["path"]), before)
        self.assertEqual(self.audit(result)["tasks"][0]["decision"], "not_executed")

    def test_unknown_presence_is_not_revived_by_deadline(self):
        plan = self.plan()
        first = plan["tasks"][0]["attempts"][0]
        first.pop("presence_evidence")
        first.update(completed_at="2026-10-14T12:00:00+08:00", delivery_at="2026-10-14T10:00:00+08:00",
                     cutoff_state_evidence="截止时仍运行", status_at_cutoff="running")
        plan.update(deadline_mode=True, cutoff_at="2026-10-14T11:00:00+08:00")
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 0)
        self.assertEqual(self.audit(result)["tasks"][0]["decision"], "pending")

    def test_first_success_survives_unverified_extra_session(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"][1]["client_presence"] = "unknown"
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1)
        self.assertEqual(self.audit(result)["tasks"][0]["decision"], "first_success")

    def test_schema3_timeout_still_exports_word_and_records_both_failures(self):
        plan = self.plan()
        first, second = plan["tasks"][0]["attempts"]
        first.update(status="failed", reason="第一次未生成必需报告")
        second.update(status="running", completeness="none", cutoff_state_evidence="截止根会话仍在运行",
                      delivery_search_complete=True, output_snapshot_complete=True, official_output=[])
        plan.update(deadline_mode=True, cutoff_at="2026-10-14T11:00:00+08:00")
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        cells = x.read_workbook(plan["excel"]["path"])["记录"]["cells"]
        self.assertEqual(cells["F2"]["value"], "两次测试均未成功")
        self.assertIn("运行时间过长，判定为执行失败", cells["G2"]["value"])

    def test_center_alignment_even_when_cell_value_does_not_change(self):
        plan = self.plan()
        path = Path(plan["excel"]["path"])
        changes = [{"sheet": "记录", "cell": "F2", "expected": "历史B时间", "value": "历史B时间"}]
        result = x.update_workbook(path, changes, self.fixture.root / "center-backup.xlsx", center=True)
        self.assertEqual(result["status"], "updated")
        cell = x.read_workbook(path)["记录"]["cells"]["F2"]
        self.assertEqual(cell["value"], "历史B时间")
        with zipfile.ZipFile(path) as z:
            xf = ET.fromstring(z.read("xl/styles.xml")).find(x.Q + "cellXfs")[int(cell["style"])]
            self.assertEqual(xf.find(x.Q + "alignment").get("horizontal"), "center")

    def test_segments_sum_whole_attempt_and_keep_evidence_in_excel(self):
        plan = self.plan()
        first = plan["tasks"][0]["attempts"][0]
        first.update(duration_source="client_segments", duration_segments_complete=True,
                     completed_at="2026-10-14T09:07:00+08:00", duration_segments=[
                         {"session_id": "s1", "started_at": "2026-10-14T09:01:00+08:00", "completed_at": "2026-10-14T09:03:00+08:00", "evidence": "压缩前显示段"},
                         {"session_id": "s1", "started_at": "2026-10-14T09:04:00+08:00", "completed_at": "2026-10-14T09:07:00+08:00", "evidence": "压缩后继续执行段；中间1分钟是已核实暂停"}])
        result = c.execute(plan, True)
        cells = x.read_workbook(plan["excel"]["path"])["记录"]["cells"]
        self.assertEqual(cells["F2"]["value"], "5m0s")
        self.assertIn("2段累计", cells["G2"]["value"])
        self.assertEqual(len(self.audit(result)["tasks"][0]["attempts"][0]["duration_segments"]), 2)

    def test_overlapping_or_cross_session_segments_do_not_invent_time(self):
        for kind in ("overlap", "other_session", "incomplete"):
            run = fixtures.run("success")
            run.update(duration_source="execution_segments", duration_segments_complete=kind != "incomplete", completed_at="2026-10-14T09:08:00+08:00", duration_segments=[
                {"session_id": "s1", "started_at": "2026-10-14T09:01:00+08:00", "completed_at": "2026-10-14T09:05:00+08:00", "evidence": "段一"},
                {"session_id": "other" if kind == "other_session" else "s1", "started_at": "2026-10-14T09:04:00+08:00", "completed_at": "2026-10-14T09:07:00+08:00", "evidence": "段二"}])
            normalize_duration(run, True)
            self.assertNotIn("duration", run)
            self.assertIn("duration_issue", run)
            self.assertEqual(run["status"], "success")

    def test_fallback_time_is_explicit_and_deducts_confirmed_pause(self):
        run = fixtures.run("success")
        run.update(duration_source="elapsed_fallback", completed_at="2026-10-14T09:11:00+08:00", pause_seconds=60, pause_evidence="已核实暂停一分钟")
        normalize_duration(run, True)
        self.assertEqual(run["duration"], "9m0s（提交/执行至完成经过时长）")

    def test_client_total_is_not_added_to_segments_again(self):
        run = fixtures.run("success")
        run.update(duration_source="client_total", duration_segments=[{"invalid": "deliberately unused"}])
        normalize_duration(run, True)
        self.assertEqual(run["duration"], "2m3s")

    def test_schema3_missing_time_basis_does_not_block_deliveries(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"][0].pop("duration_source")
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1)
        cells = x.read_workbook(plan["excel"]["path"])["记录"]["cells"]
        self.assertIn("用时无法确认", cells["G2"]["value"])
        self.assertTrue(any("用时无法确认" in i["issue"] for i in result["issues"]))

    def test_center_written_cells_only_preserves_other_styles_and_is_idempotent(self):
        plan = self.plan()
        path = Path(plan["excel"]["path"])
        with zipfile.ZipFile(path) as z:
            original = {n: z.read(n) for n in z.namelist()}
        old_cells = x.read_workbook(path)["记录"]["cells"]
        c.execute(plan, True)
        cells = x.read_workbook(path)["记录"]["cells"]
        with zipfile.ZipFile(path) as z:
            old_styles = ET.fromstring(original["xl/styles.xml"]).find(x.Q + "cellXfs")
            new_styles = ET.fromstring(z.read("xl/styles.xml")).find(x.Q + "cellXfs")
            for i, old in enumerate(old_styles):
                self.assertEqual(ET.tostring(old), ET.tostring(new_styles[i]))
            for address in ["F1", "F2", "G2"]:
                xf = new_styles[int(cells[address]["style"])]
                self.assertEqual(xf.find(x.Q + "alignment").get("horizontal"), "center")
                self.assertEqual(xf.find(x.Q + "alignment").get("vertical"), "center")
            for name, value in original.items():
                if name not in {"xl/styles.xml", "xl/worksheets/sheet1.xml"}:
                    self.assertEqual(z.read(name), value)
        for address in ["D2", "E2"]:
            self.assertEqual(cells[address], old_cells[address])
        before = x.digest(path)
        plan["excel"]["sha256"] = before
        plan["excel"]["headers"][0]["expected"] = cells["F1"]["value"]
        result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "unchanged")
        self.assertEqual(x.digest(path), before)


class DeliveryReaderTests(unittest.TestCase):
    def test_all_mobilework_delivery_calls_and_success_results_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "app.db"
            with sqlite3.connect(path) as db:
                db.executescript("CREATE TABLE session(id TEXT,directory TEXT); CREATE TABLE part(id TEXT,session_id TEXT,time_created INTEGER,data TEXT);")
                db.execute("INSERT INTO session VALUES (?,?)", ("s1", tmp))
                for i, name in enumerate(["report.docx", "research_plan.md", "report.pdf"], 1):
                    data = {"tool": "result_file", "state": {"status": "completed", "time": {"end": 1791936000000 + i},
                            "input": {"output_paths": ["WRONG-input-only.txt"]}, "output": json.dumps({"ok": True, "delivered": [name]})}}
                    db.execute("INSERT INTO part VALUES (?,?,?,?)", (str(i), "s1", i, json.dumps(data)))
                db.execute("INSERT INTO part VALUES (?,?,?,?)", ("bad", "s1", 4, json.dumps({"tool": "result_file", "state": {"status": "error"}})))
            db.close()
            before = x.digest(path)
            result = read_deliveries("mobilework", path, "s1")
            self.assertEqual(len(result["formal_deliveries"]), 3)
            self.assertEqual(result["successful_delivery_events"], 3)
            self.assertEqual(x.digest(path), before)
            self.assertNotIn("WRONG", str(result))

    def test_all_workbuddy_results_keep_plan_attachment_and_cutoff(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "session.jsonl"
            events = []
            for i, name in enumerate(["report.docx", "research_plan.md", "late.pdf"]):
                events.append({"type": "function_call_result", "name": "present_files", "status": "completed", "sessionId": "s1", "cwd": tmp,
                    "timestamp": f"2026-10-14T09:0{i}:00+08:00", "id": str(i), "output": [{"text": json.dumps({"type": "present_files_result", "files": [name]})}]})
            path.write_text("\n".join(json.dumps(e) for e in events), encoding="utf-8")
            result = read_deliveries("workbuddy", path, "s1", "2026-10-14T09:01:00+08:00")
            self.assertEqual([Path(f["source"]).name for f in result["formal_deliveries"]], ["report.docx", "research_plan.md"])


if __name__ == "__main__":
    unittest.main()
