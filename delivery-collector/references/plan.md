# 执行计划 JSON 约定

计划由 agent 在完成发现与证据判断后生成。不是给使用者额外填写的表。脚本不推测题号、会话归属或内容质量；`success` 是 agent 有证据支持的判断，不是文件扩展名判断。

## 执行步骤

1. 使用只读命令检查题目、实际记录和 Excel。将本次输入与实际定位信息写入计划，路径为当前环境绝对路径。
2. 每题列完整候选、必需交付和证据。`attempts` 已按核实的提交顺序排列。其它批次/子代理等写入 `excluded_candidates`，保留原始证据引用。完全未执行题 `attempts: []`。疑似会话归属不明设置 `pending`，不能伪造两次运行。
3. 成功候选须核查全部实质文件，包括正式附件和项目依赖，列入 `artifacts`；每个文件必须有源SHA256、正式提供依据及源会话。每次读取源后冻结哈希；变化后重新判断。
4. `workbook` 输出按现有编号定位对象列。Excel路径默认在数据集根。只有明确本次调用身份/槽位时允许更新表头，只有明确以新批次信息为准时设 `overwrite_existing: true`。题号列、用时列、说明列必须不同；仅目标对象列能写。
5. 运行 `execute PLAN.json` 预检。输出计划核查表和JSON，但不复制到问题目录，不更新结果表。预检的失败DOCX样本只生成于审计目录，供检查。明确项继续运行 `execute PLAN.json --apply`，不要把预检有问题的项悄悄改为成功。
6. 检查执行JSON里的实际复制、目标哈希和 Excel 状态。预检清单不等于完成清单。单题阻碍不阻碍其他题；Excel定位失败仍可收文件但必须告知未登记。

## 示例（单对象，第一次失败、第二次成功）

以下路径和哈希为占位示例，agent 必须替换为真实值。`task_id` 字符串，不用浮点数。未指定 `excel` 时脚本仍收文件并报告记录表缺失，不声称流程全部完成。

```json
{
  "schema_version": 1,
  "dataset_root": "/data/deepinsight",
  "batch_id": "batch-20261014",
  "object": {
    "name": "WorkBuddy", "version": "5.5.6", "model": "glm-5.2",
    "prefix": "workbuddy5.5.6"
  },
  "read_roots": ["/projects/workbuddy-batch"],
  "discovery_complete": true,
  "discovery_evidence": "已遍历指定工作空间该批次全部根会话及消息；记录查询/导出来源",
  "excel": {
    "path": "/data/deepinsight/DeepInsight测试结果记录.xlsx",
    "sha256": "读取时的实际SHA256",
    "sheet": "Sheet1",
    "task_id_column": "B", "time_column": "F", "note_column": "G",
    "target_evidence": "本次调用明确要求填写测试对象2；该槽位是F/G列",
    "overwrite_existing": true,
    "headers": [{
      "cell": "F2", "expected": "读取时的原表头完整文本",
      "value": "测试对象2（名称/版本：WorkBuddy-5.5.6；大模型名称：glm-5.2）用时"
    }]
  },
  "tasks": [{
    "task_id": "2.3",
    "question_dir": "/data/deepinsight/维度二/问题3",
    "requirements": ["完整可读中国平安报告", "检索与财务截止日期"],
    "excluded_candidates": [],
    "attempts": [
      {
        "session_id": "s-first",
        "submitted_at": "2026-10-14T09:00:00+08:00",
        "submission_evidence": "消息ID、提交时间（含时区）及记录来源",
        "match_evidence": "实际提交正文和素材匹配的证据",
        "status": "failed", "reason": "显示完成但未提供可取得的交付文件",
        "duration": "3m4s", "duration_evidence": "客户端显示及其累计口径"
      },
      {
        "session_id": "s-second",
        "submitted_at": "2026-10-14T09:10:00+08:00",
        "submission_evidence": "第二次实际提交时间、消息ID、记录来源",
        "match_evidence": "同题同素材匹配证据",
        "status": "success", "reason": "报告可读取，必需内容存在",
        "duration": "17m19s", "duration_evidence": "客户端完成状态时长",
        "completeness": "complete",
        "validation_evidence": "实际打开并检查主题、截止日期和完整性；来源记录引用",
        "artifacts": [{
          "session_id": "s-second",
          "source": "/projects/workbuddy-batch/project-b/中国平安报告.docx",
          "sha256": "源文件实际SHA256",
          "evidence": "工具/消息ID、成功交付结果及实际项目文件"
        }]
      }
    ]
  }]
}
```

未能完整查看批次时 `discovery_complete: false`，无会话题即待确认。设true须给 `discovery_evidence`。实际只测试部分维度也可以完整发现范围：例如确认其余维度没有该批次提交记录。不要因没有文件而直接设置未执行。

运行状态仅 `success/failed/running/unknown`。`submitted_at` 必须是包含时区的实际提交时间，数组严格按提交先后排序，顺序无法判断则待确认。当第一次未知时即使第二次成功也暂不选第二次。只有一次失败或大于两次按规则待确认，明确题不受阻。无需在计划复制完整数据库或内部推理；保存证据ID、定位信息即可。

项目文件在每个artifact上增加 `project_name`（原项目名）、`relative_path`（项目内部相对路径，例如 `css/style.css`）；其source必须对应该原文件名。同一项目的所有实际依赖逐文件列出，不用整个工作区目录复制。避免交付根名与单文件重名。

## 两次失败的正式输出

第二次run增加 `official_output_complete: true` 和 `official_output`，按用户界面顺序列所有正式内容。排除推理、thinking和隐藏工具日志，不靠关键词删正文。正式输出无法完整取得时改待确认，不填true。脚本生成DOCX的样式仅为可读归档，不改正文。

```json
{
  "official_output_complete": true,
  "official_output": [
    {"kind": "text", "visible": true, "text": "原样正式回答及可见报错，不写收集者摘要", "evidence": "正式消息/part ID、原记录位置"},
    {"kind": "table", "visible": true, "rows": [["原表头1", "原表头2"], ["值1", "值2"]], "evidence": "界面正式表格出处"},
    {"kind": "image", "visible": true, "source": "/projects/project-b/正式图表.png", "sha256": "实际SHA256", "evidence": "界面显示图片出处"}
  ]
}
```

仅在真实有表格或图片时使用相应块。DOCX生成后对照原正式输出核查正文、表格和图像；有可用渲染工具时检查实际页面。失败部分附件不默认进入正式归档。原本正式输出为空且没有任何可见内容时，应如实标明证据缺口，不能填造回答。

## 重复执行和限制

每次生成独立审计目录，JSON持续保存已执行动作。文件同名同SHA跳过，同名不同SHA待确认。结果表前置SHA检查、单元格expected检查和原件备份防止无意覆盖；若其它对象同时更新表，应重新读取并生成计划，不重用过期SHA。

Excel助手只针对普通transitional XLSX的文本记录单元格，不覆盖公式、合并区域非锚点，不增加工作表/列或重构格式。非XLSX、受保护/加密、复杂不支持布局，agent使用能保留格式的可用工具；否则记待确认。其它ZIP部件原内容保持不变，不做公式重算。

清单里长于Excel单元格上限的内容显示“完整见JSON”，JSON保留完整证据。人工核查单元格初始为空，agent不得填“人工已验证”。
