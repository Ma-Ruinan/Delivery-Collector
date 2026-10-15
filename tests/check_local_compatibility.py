"""Optional local acceptance check; input dataset is strictly read-only.

python tests/check_local_compatibility.py CORPUS OUTPUT_DIRECTORY
Copies only result workbooks to a new output directory and validates narrow
updates there. No sessions or authentic deliverables are archived or uploaded.
"""
import argparse
import json
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "delivery-collector" / "scripts"))
import collector as c
import xlsx_io as x
from excel_layout import resolve_slot


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("corpus"); parser.add_argument("output")
    args = parser.parse_args()
    corpus, output = Path(args.corpus).resolve(), Path(args.output).resolve()
    if output == corpus or output.is_relative_to(corpus) or corpus.is_relative_to(output):
        raise ValueError("验收输出必须与源数据集目录分离")
    output.mkdir(parents=True, exist_ok=False)
    result = []
    for directory in sorted(corpus.iterdir()):
        if not directory.is_dir() or not directory.name.lower().startswith(("alphadata", "deepinsight", "常规能力")):
            continue
        roster = c.roster(str(directory))
        item = {"dataset": directory.name, "questions": len(roster["tasks"]), "workbooks": []}
        for original in directory.glob("*结果记录.xlsx"):
            before_hash = x.digest(original)
            path = output / original.name
            shutil.copy2(original, path)
            sheets = x.read_workbook(path)
            initial_layout = resolve_slot({"object_slot": 2}, sheets,
                                          {"name": "验收对象", "version": "v2", "model": "验收模型"})
            sheet_name = initial_layout["sheet"]; cells = sheets[sheet_name]["cells"]
            task_rows = [(addr, cell) for addr, cell in cells.items()
                         if x.CELL.fullmatch(addr)[1] == initial_layout["task_id_column"] and
                         cell["value"] and cell["value"] not in ("任务编号", "题号", "问题编号")]
            address = task_rows[0][0]; row = x.CELL.fullmatch(address)[2]
            change_cell = initial_layout["time_column"] + row
            with zipfile.ZipFile(path) as z:
                before = {n: z.read(n) for n in z.namelist()}
            update = x.update_workbook(path, [{"sheet": sheet_name, "cell": change_cell,
                                               "expected": cells.get(change_cell, {}).get("value"),
                                               "value": "验收测试值，仅在副本"}], output / (original.stem + "-备份.xlsx"))
            with zipfile.ZipFile(path) as z:
                for name, payload in before.items():
                    if name != sheets[sheet_name]["part"]:
                        assert z.read(name) == payload, name
            after_cells = x.read_workbook(path)[sheet_name]["cells"]
            for coord, old in cells.items():
                if coord != change_cell:
                    assert after_cells[coord] == old, coord
            assert x.digest(original) == before_hash, "源记录表不得修改"
            slot = resolve_slot({"object_slot": 3}, x.read_workbook(path),
                                {"name": "验收对象", "version": "v2", "model": "验收模型"})
            narrow_before = x.read_workbook(path)[sheet_name]["cells"]
            headers = [{"sheet": slot["sheet"], **h} for h in slot["headers"]]
            time_cell = slot["time_column"] + row
            x.update_workbook(path, headers + [{"sheet": sheet_name, "cell": time_cell, "expected": None,
                                                "value": "新增第三对象验收值，仅在副本"}],
                              output / (original.stem + "-扩展前.xlsx"), slot["extensions"])
            expanded = x.read_workbook(path)[sheet_name]["cells"]
            for coord, old in narrow_before.items(): assert expanded[coord] == old, coord
            assert expanded[time_cell]["value"] == "新增第三对象验收值，仅在副本"
            assert x.digest(original) == before_hash
            item["workbooks"].append({"name": original.name, "edited_cell": change_cell,
                                       "other_cells_and_parts_preserved": True, "status": update["status"],
                                       "third_object_added": True, "new_time_cell": time_cell})
        result.append(item)
    # A reusable failure transcript sample for visual/structural acceptance.
    failed = output / "验收样例-1.1-未成功任务界面输出.docx"
    c.make_failure(failed, [
        {"kind": "text", "visible": True, "text": "正式回答\n已完成数据读取，但未能生成要求的交付文件。\n界面错误：文件导出失败。", "evidence": "synthetic-formal-message"},
        {"kind": "table", "visible": True, "rows": [["项目", "状态"], ["报告导出", "失败"]], "evidence": "synthetic-formal-table"}], [output])
    print(json.dumps({"checks": result, "failure_sample": str(failed)}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
