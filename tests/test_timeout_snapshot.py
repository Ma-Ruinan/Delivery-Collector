"""Regression cases for cutoff output snapshots and independent failure records."""
import json
import unittest
from pathlib import Path
from unittest.mock import patch

import test_scheduled as fixtures

c = fixtures.c
x = fixtures.x


class TimeoutSnapshotTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.ScheduledIntegrationTests()
        self.fixture.setUp()

    def tearDown(self):
        self.fixture.tearDown()

    def plan(self):
        plan = self.fixture.plan()
        self.fixture.label_groups(plan)
        plan["tasks"][0]["attempts"] = [fixtures.running(1), fixtures.running(2)]
        return plan

    def cells(self, plan):
        return x.read_workbook(plan["excel"]["path"])["记录"]["cells"]

    def output(self, plan):
        from docx import Document
        path = Path(plan["tasks"][0]["question_dir"]) / "workbuddy5.5.6-1.1-未成功任务界面输出.docx"
        return "\n".join(p.text for p in Document(path).paragraphs)

    def assert_failed_record(self, plan):
        cells = self.cells(plan)
        self.assertEqual(cells["F2"]["value"], "两次测试均未成功")
        self.assertEqual(cells["G2"]["value"].count("运行时间过长，判定为执行失败"), 2)
        self.assertIn("1h59m0s", cells["G2"]["value"])
        self.assertIn("1h58m0s", cells["G2"]["value"])
        self.assertEqual(cells["D2"]["value"], "对象1历史用时")

    def test_unfinished_answer_saves_current_output_not_thinking(self):
        plan = self.plan()
        selected = plan["tasks"][0]["attempts"][1]
        selected.update(official_output_complete=False, output_snapshot_complete=True,
                        official_output=[
                            {"kind": "text", "text": "截止时已显示，尚未写完的正式回答", "visible": True, "evidence": "message-part-2"},
                            {"kind": "reasoning", "text": "不得归档的内部推理"},
                        ])
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        text = self.output(plan)
        self.assertIn("尚未写完的正式回答", text)
        self.assertNotIn("内部推理", text)
        self.assertIn(fixtures.CUTOFF, text)
        self.assert_failed_record(plan)

    def test_partial_snapshot_preserves_available_output_and_marks_gap(self):
        plan = self.plan()
        selected = plan["tasks"][0]["attempts"][1]
        selected.update(official_output_complete=False, official_output=[
            {"kind": "text", "text": "已取得的正式片段", "visible": True, "evidence": "visible-part"}])
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        self.assertIn("已取得的正式片段", self.output(plan))
        self.assertIn("快照未完整取得", self.output(plan))
        self.assertTrue(any("快照未完整取得" in i["issue"] for i in result["issues"]))
        self.assert_failed_record(plan)
        self.assertIn("快照未完整取得", self.cells(plan)["G2"]["value"])

    def test_unread_snapshot_does_not_claim_verified_absence(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"][1]["official_output_complete"] = False
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        text = self.output(plan)
        self.assertNotIn("已核实未见正式界面输出", text)
        self.assertIn("快照未完整取得", text)
        self.assert_failed_record(plan)

    def test_docx_error_cannot_suppress_confirmed_failure_registration(self):
        plan = self.plan()
        with patch.object(c, "make_failure", side_effect=OSError("DOCX写入失败")):
            result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "updated", result)
        self.assert_failed_record(plan)
        self.assertIn("输出快照保存失败", self.cells(plan)["G2"]["value"])
        audit = json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))
        self.assertEqual(audit["tasks"][0]["decision"], "both_failed")
        self.assertEqual(audit["tasks"][0]["excel_status"], "updated")
        self.assertEqual(audit["tasks"][0]["collection_status"], "pending")
        self.assertEqual(result["copied"], 0)

    def test_copy_error_cannot_suppress_confirmed_failure_registration(self):
        plan = self.plan()
        with patch.object(c, "copy_one", side_effect=OSError("目标文件无法写入")):
            result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "updated", result)
        self.assert_failed_record(plan)
        self.assertIn("输出快照保存失败", self.cells(plan)["G2"]["value"])

    def test_existing_failed_sessions_still_register_when_output_unavailable(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"] = [fixtures.run("failed", 1), fixtures.run("failed", 2)]
        result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "updated", result)
        self.assertEqual(self.cells(plan)["F2"]["value"], "两次测试均未成功")
        self.assertIn("缺少必需文件", self.cells(plan)["G2"]["value"])
        self.assertEqual(result["copied"], 0)

    def test_cutoff_enables_deadline_if_flag_omitted(self):
        plan = self.plan()
        del plan["deadline_mode"]
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        self.assert_failed_record(plan)

    def test_explicit_non_deadline_mode_does_not_timeout_running_sessions(self):
        plan = self.plan()
        plan["deadline_mode"] = False
        before = x.digest(plan["excel"]["path"])
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 0)
        self.assertEqual(x.digest(plan["excel"]["path"]), before)

    def test_invalid_timing_does_not_demote_confirmed_timeout_to_unknown(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"][1]["started_at"] = "not-a-time"
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        self.assertEqual(self.cells(plan)["F2"]["value"], "两次测试均未成功")
        note = self.cells(plan)["G2"]["value"]
        self.assertIn("第二次用时用时无法确认", note)
        self.assertEqual(note.count("运行时间过长，判定为执行失败"), 2)
        audit = json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))
        self.assertIn("duration_issue", audit["tasks"][0]["attempts"][1])

    def test_snapshot_repeat_is_idempotent_and_never_uses_first_output(self):
        plan = self.plan()
        first, second = plan["tasks"][0]["attempts"]
        first["official_output"] = [
            {"kind": "text", "text": "第一次输出不应归档", "visible": True, "evidence": "first-part"}]
        second.update(official_output_complete=False, output_snapshot_complete=True,
                      official_output=[
                          {"kind": "text", "text": "第二次截至截止的正式内容", "visible": True, "evidence": "second-part"}])
        c.execute(plan, True)
        before = x.digest(plan["excel"]["path"])
        plan["excel"]["sha256"] = before
        result = c.execute(plan, True)
        self.assertEqual(result["already_present"], 1, result)
        self.assertEqual(result["excel"]["status"], "unchanged")
        self.assertEqual(x.digest(plan["excel"]["path"]), before)
        self.assertNotIn("第一次输出", self.output(plan))
        self.assertIn("第二次截至截止的正式内容", self.output(plan))

    def test_target_conflict_keeps_file_but_records_failure(self):
        plan = self.plan()
        path = Path(plan["tasks"][0]["question_dir"]) / "workbuddy5.5.6-1.1-未成功任务界面输出.docx"
        # Synthetic pre-existing delivery; never overwrite different contents.
        c.make_failure(path, [{"kind": "text", "text": "已有历史输出"}], [])
        before = c.digest(path)
        result = c.execute(plan, True)
        self.assertEqual(c.digest(path), before)
        self.assertEqual(result["copied"], 0)
        self.assert_failed_record(plan)
        self.assertIn("输出快照保存失败", self.cells(plan)["G2"]["value"])

    def test_unverified_visible_content_is_not_exported_but_failure_is_recorded(self):
        plan = self.plan()
        plan["tasks"][0]["attempts"][1].update(
            official_output_complete=False, output_snapshot_complete=True,
            official_output=[{"kind": "text", "text": "可见性不明的内容", "evidence": "some-part"}])
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 0)
        self.assertTrue(any("可见性" in i["issue"] for i in result["issues"]))
        self.assert_failed_record(plan)

    def test_fresh_client_duration_does_not_depend_on_invalid_start_metadata(self):
        value = fixtures.running(2)
        value.update(started_at="not-a-time", duration="1h2m3s", duration_as_of=fixtures.CUTOFF)
        result = fixtures.at_deadline(value, fixtures.CUTOFF)
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["duration"], "1h2m3s")
        self.assertNotIn("duration_issue", result)

    def test_dry_run_docx_error_does_not_change_workbook(self):
        plan = self.plan()
        before = x.digest(plan["excel"]["path"])
        with patch.object(c, "make_failure", side_effect=OSError("DOCX写入失败")):
            result = c.execute(plan, False)
        self.assertEqual(x.digest(plan["excel"]["path"]), before)
        audit = json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))
        self.assertEqual(audit["tasks"][0]["excel_status"], "planned")


if __name__ == "__main__":
    unittest.main()
