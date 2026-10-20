"""Unattended slot targeting, cutoff selection and path-only delivery cases."""
import copy
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "delivery-collector" / "scripts"))
import collector as c
import xlsx_io as x
from deadline import at_deadline, snapshot_task
from excel_layout import resolve_slot
import test_collector as fixtures
from test_collector import run

CUTOFF = "2026-10-14T11:00:00+08:00"


def running(number=1, completeness="none"):
    value = run("running", number)
    value.update(completeness=completeness, cutoff_state_evidence="截止快照/根会话状态记录",
                 delivery_search_complete=True, official_output_complete=True, official_output=[])
    return value


class DeadlineTests(unittest.TestCase):
    def test_two_timeouts_are_two_failures_and_durations_are_explicit(self):
        task = snapshot_task({"attempts": [running(1), running(2)]}, CUTOFF)
        choice, selected, _ = c.select_run(task, True)
        self.assertEqual(choice, "both_failed")
        self.assertEqual(task["attempts"][0]["duration"], "1h59m0s（截至收集时经过时长）")
        time, note = c.excel_values(choice, task, selected)
        self.assertEqual(time, "两次测试均未成功")
        self.assertEqual(note.count("运行时间过长，判定为执行失败"), 2)
        self.assertIn("第一次", note); self.assertIn("第二次", note)

    def test_first_success_ignores_extra_timeout(self):
        task = snapshot_task({"attempts": [run("success"), running(2)]}, CUTOFF)
        choice, selected, _ = c.select_run(task, True)
        self.assertEqual(choice, "first_success")
        self.assertEqual(c.excel_values(choice, task, selected), ("2m3s", "一次运行成功"))

    def test_timeout_then_success(self):
        task = snapshot_task({"attempts": [running(1), run("success", 2)]}, CUTOFF)
        choice, selected, _ = c.select_run(task, True)
        self.assertEqual(choice, "second_success")
        self.assertIn("运行时间过长", c.excel_values(choice, task, selected)[1])

    def test_failed_then_timeout(self):
        task = snapshot_task({"attempts": [run("failed"), running(2)]}, CUTOFF)
        self.assertEqual(c.select_run(task, True)[0], "both_failed")

    def test_unknown_is_not_guessed_as_timeout(self):
        value = running(); value["status"] = "unknown"
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "unknown")
        value = running(); value.pop("cutoff_state_evidence")
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "unknown")
        value = running(); value["delivery_search_complete"] = False
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "unknown")

    def test_incomplete_running_is_failure_not_success(self):
        value = running(completeness="incomplete")
        value["artifacts"] = [{"source": "partial-result"}]
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "failed")

    def test_full_final_delivery_before_cutoff_can_succeed_while_ui_running(self):
        value = running(completeness="complete")
        value.update(artifacts=[{"source": "formal-file"}], validation_evidence="全部文件已实际打开",
                     delivery_at="2026-10-14T10:00:00+08:00")
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "success")
        value["delivery_at"] = "2026-10-14T12:00:00+08:00"
        self.assertNotEqual(at_deadline(value, CUTOFF)["status"], "success")

    def test_completion_after_cutoff_not_taken_as_pre_cutoff_success(self):
        value = running(); value["status"] = "success"
        value["completed_at"] = "2026-10-14T12:00:00+08:00"
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "failed")

    def test_missing_start_or_future_submission_stays_unknown(self):
        value = running(); value["submitted_at"] = "2026-10-14T12:00:00+08:00"
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "unknown")
        value.pop("submitted_at")
        self.assertEqual(at_deadline(value, CUTOFF)["status"], "unknown")

    def test_pause_requires_evidence_and_does_not_count_twice(self):
        value = running(); value.update(pause_seconds=60, pause_evidence="明确暂停段")
        result = at_deadline(value, CUTOFF)
        self.assertEqual(result["elapsed_seconds"], 7080)
        value.pop("pause_evidence")
        result = at_deadline(value, CUTOFF)
        self.assertEqual(result["status"], "failed")
        self.assertIn("暂停时间", result["duration_issue"])
        self.assertNotIn("duration", result)

    def test_input_snapshot_not_mutated(self):
        value = running(); before = copy.deepcopy(value)
        at_deadline(value, CUTOFF)
        self.assertEqual(value, before)

    def test_client_duration_at_equivalent_timezone_cutoff_is_preserved(self):
        value = running(); value.update(duration="1h2m3s", duration_as_of="2026-10-14T03:00:00+00:00")
        self.assertEqual(at_deadline(value, CUTOFF)["duration"], "1h2m3s")


class ScheduledIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.fixture = fixtures.FileAndWorkbookTests()
        self.fixture.setUp()

    def tearDown(self):
        self.fixture.tearDown()

    def plan(self, slot=2):
        plan = self.fixture.plan()
        plan.update(schema_version=2, unattended=True, deadline_mode=True, cutoff_at=CUTOFF)
        plan["excel"] = {"path": plan["excel"]["path"], "sha256": plan["excel"]["sha256"], "object_slot": slot}
        # The first fixture includes a valid first run and an unneeded extra run.
        return plan

    def replace_xlsx_parts(self, path, replacements):
        with zipfile.ZipFile(path) as z:
            parts = {i.filename: z.read(i.filename) for i in z.infolist()}
        parts.update(replacements)
        with zipfile.ZipFile(path, "w") as z:
            for name, payload in parts.items(): z.writestr(name, payload)

    def label_groups(self, plan, count=2, merged=False):
        path = Path(plan["excel"]["path"])
        headers = ["维度", "任务编号", "任务名称"]
        values = ["维度一", "1.1", "问题1"]
        for index in range(1, count + 1):
            headers += [f"测试对象{index}（名称/版本：旧对象{index}；大模型名称：旧模型）用时", "任务完成情况、说明"]
            values += [f"对象{index}历史用时", f"对象{index}历史说明"]
        x.write_audit(path, {"记录": [headers, values], "其他表": [["其他工作表不动"]]})
        if merged:
            with zipfile.ZipFile(path) as z:
                raw = z.read("xl/worksheets/sheet1.xml").decode()
            root = ET.fromstring(raw)
            data = root.find(x.Q + "sheetData")
            # Put group labels in row1 and time/note subheaders in row2;
            # records are row3. Preserve existing styles and column widths.
            for row in list(data):
                row.set("r", str(int(row.get("r")) + 1))
                for cell in row:
                    ref = x.CELL.fullmatch(cell.get("r")); cell.set("r", ref[1] + str(int(ref[2]) + 1))
            row1 = ET.Element(x.Q + "row", {"r": "1"})
            merge_root = ET.Element(x.Q + "mergeCells", {"count": str(count)})
            for index in range(1, count + 1):
                col = x.col_name(2 * index + 2)
                c1 = ET.SubElement(row1, x.Q + "c", {"r": col + "1", "t": "inlineStr", "s": "1"})
                ET.SubElement(ET.SubElement(c1, x.Q + "is"), x.Q + "t").text = f"测试对象{index}"
                ET.SubElement(merge_root, x.Q + "mergeCell", {"ref": col + "1:" + x.col_name(2 * index + 3) + "1"})
                for cell in list(data)[0]:
                    if cell.get("r") == col + "2":
                        for t in cell.iter(x.Q + "t"): t.text = "用时"
            data.insert(0, row1)
            root.append(merge_root)
            self.replace_xlsx_parts(path, {"xl/worksheets/sheet1.xml": ET.tostring(root, encoding="utf-8", xml_declaration=True)})
        plan["excel"]["sha256"] = x.digest(path)
        return path

    def test_path_only_delivery_without_download_component_is_success(self):
        plan = self.plan(); self.label_groups(plan)
        file = plan["tasks"][0]["attempts"][0]["artifacts"][0]
        file.update(delivery_channel="formal_path", evidence="正式回答给出报告.md路径；文件真实存在并已打开")
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1)
        cells = x.read_workbook(plan["excel"]["path"])["记录"]["cells"]
        self.assertEqual(cells["G2"]["value"], "一次运行成功")
        self.assertIn("WorkBuddy-5.5.6", cells["F1"]["value"])

    def test_one_two_three_existing_slots(self):
        for count in (1, 2, 3):
            with self.subTest(count=count):
                temp = tempfile.TemporaryDirectory()
                try:
                    path = Path(temp.name) / "record.xlsx"
                    header = ["任务编号"]
                    for index in range(1, count + 1):
                        header += [f"测试对象{index}用时", "任务完成情况、说明"]
                    x.write_audit(path, {"记录": [header, ["W1"] + [""] * (count * 2)]})
                    for slot in range(1, count + 1):
                        result = resolve_slot({"object_slot": slot}, x.read_workbook(path), {"name": "MobileWork1009", "version": "1009", "model": "DSV4"})
                        self.assertEqual(result["time_column"], x.col_name(slot * 2))
                        self.assertFalse(result["extensions"])
                        self.assertIn("MobileWork1009", result["headers"][0]["value"])
                        self.assertNotIn("1009-1009", result["headers"][0]["value"])
                finally: temp.cleanup()

    def test_add_third_slot_preserves_existing_results_styles_widths_and_parts(self):
        plan = self.plan(slot=3); path = self.label_groups(plan)
        before = x.read_workbook(path)
        with zipfile.ZipFile(path) as z: original = {n: z.read(n) for n in z.namelist()}
        result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "updated", result)
        after = x.read_workbook(path)
        for address, old in before["记录"]["cells"].items():
            self.assertEqual(after["记录"]["cells"][address], old)
        cells = after["记录"]["cells"]
        self.assertIn("测试对象3", cells["H1"]["value"])
        self.assertEqual(cells["H2"]["value"], "2m3s")
        self.assertEqual(cells["I2"]["value"], "一次运行成功")
        self.assertEqual(cells["H1"]["style"], cells["F1"]["style"])
        with zipfile.ZipFile(path) as z:
            for name in original:
                if name != "xl/worksheets/sheet1.xml": self.assertEqual(z.read(name), original[name])
            root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
            cols = root.find(x.Q + "cols")
            widths = {int(node.get("min")): node.get("width") for node in cols}
            self.assertEqual(widths[8], widths[6]); self.assertEqual(widths[9], widths[7])

    def test_add_third_slot_to_merged_header_preserves_merges(self):
        plan = self.plan(slot=3); path = self.label_groups(plan, merged=True)
        before = x.read_workbook(path)
        result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "updated", result)
        after = x.read_workbook(path)["记录"]
        self.assertIn("H1:I1", after["merged"])
        self.assertEqual(after["cells"]["H3"]["value"], "2m3s")
        for address, old in before["记录"]["cells"].items(): self.assertEqual(after["cells"][address], old)

    def test_one_slot_can_expand_to_three_without_copying_old_results(self):
        plan = self.plan(slot=3); path = self.label_groups(plan, count=1)
        result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "updated", result)
        cells = x.read_workbook(path)["记录"]["cells"]
        self.assertEqual(cells["D2"]["value"], "对象1历史用时")
        self.assertEqual(cells["F2"]["value"], None)
        self.assertEqual(cells["G2"]["value"], None)
        self.assertEqual(cells["H2"]["value"], "2m3s")
        self.assertIn("测试对象2", cells["F1"]["value"])

    def test_repeated_collection_does_not_add_another_third_slot(self):
        plan = self.plan(slot=3); path = self.label_groups(plan)
        c.execute(plan, True)
        before = x.digest(path)
        plan["excel"]["sha256"] = before
        result = c.execute(plan, True)
        self.assertEqual(result["excel"]["status"], "unchanged")
        self.assertEqual(x.digest(path), before)

    def test_two_timeouts_generate_observation_docx_and_fill_excel_without_waiting(self):
        from docx import Document
        plan = self.plan(); self.label_groups(plan)
        plan["tasks"][0]["attempts"] = [running(1), running(2)]
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1, result)
        path = Path(plan["tasks"][0]["question_dir"]) / "workbuddy5.5.6-1.1-未成功任务界面输出.docx"
        text = '\n'.join(p.text for p in Document(path).paragraphs)
        self.assertIn("收集核查记录", text)
        self.assertIn("未见正式界面输出", text)
        self.assertIn(CUTOFF, text)
        cells = x.read_workbook(plan["excel"]["path"])["记录"]["cells"]
        self.assertEqual(cells["F2"]["value"], "两次测试均未成功")
        self.assertEqual(cells["G2"]["value"].count("运行时间过长，判定为执行失败"), 2)

    def test_extra_timeout_only_in_audit_not_first_success_excel(self):
        plan = self.plan(); self.label_groups(plan)
        plan["tasks"][0]["attempts"][1] = running(2)
        result = c.execute(plan, True)
        audit = json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))
        self.assertTrue(audit["tasks"][0]["attempts"][1]["timed_out"])
        self.assertEqual(x.read_workbook(plan["excel"]["path"])["记录"]["cells"]["G2"]["value"], "一次运行成功")
        self.assertFalse(any(p.suffix == ".docx" for p in Path(plan["tasks"][0]["question_dir"]).iterdir()))

    def test_unattended_ambiguous_layout_does_not_guess_or_block_files(self):
        plan = self.plan(); self.label_groups(plan)
        plan["excel"]["object_slot"] = 0
        result = c.execute(plan, True)
        self.assertEqual(result["copied"], 1)
        self.assertEqual(result["excel"]["status"], "not_updated")
        self.assertTrue(result["issues"])

    def test_dry_run_does_not_extend_workbook(self):
        plan = self.plan(slot=3); path = self.label_groups(plan)
        before = x.digest(path)
        result = c.execute(plan, False)
        self.assertEqual(x.digest(path), before)
        self.assertEqual(result["copied"], 0)

    def test_instruction_legend_task_label_not_confused_with_real_header(self):
        plan = self.plan(); path = self.label_groups(plan)
        with zipfile.ZipFile(path) as z:
            raw = z.read("xl/worksheets/sheet1.xml")
        marker = b'</sheetData>'
        legend = ('<row r="10"><c r="B10" t="inlineStr"><is><t>任务编号</t></is></c></row>').encode()
        self.replace_xlsx_parts(path, {"xl/worksheets/sheet1.xml": raw.replace(marker, legend + marker)})
        result = resolve_slot({"object_slot": 2}, x.read_workbook(path), plan["object"])
        self.assertEqual(result["task_id_column"], "B")


if __name__ == "__main__":
    unittest.main()
