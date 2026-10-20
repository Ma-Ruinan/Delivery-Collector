# 执行计划 JSON 约定

计划是agent读取证据后形成的内部文件，不让使用者手工填写。脚本不猜题号、会话身份或内容质量。新计划用schema_version 2，原schema 1及显式列映射仍兼容。

## 发现与执行

1. 定时任务在发现前运行 `python scripts/collector.py cutoff`，将返回值存为cutoff_at，设deadline_mode true；预检、执行、重试复用，不用后来的遍历结束时间。已有cutoff_at而漏设开关时脚本也启用截止判断；普通非截止收集显式设deadline_mode false。
2. 按调用读取根会话、正式输出、组件/路径；匹配题号原目录，attempts按真实提交时间排序。继续/子代理不算重试，排除项写excluded_candidates及依据。状态/路径不明先自主查证，仍不明记pending，不中途提问。
3. success必须核查全部必需文件，列artifacts、哈希、来源引用。`delivery_channel: formal_path`表示正式答复给出路径，无需下载组件。必须检查实际最终文件，不能因没按钮写failed。
4. 结果表默认数据集根，结合题号/名称核实；指定序号写excel.object_slot，按[excel.md](excel.md)定位/扩展。`layout FILE --slot 3 --name MobileWork1009 --version 1009 --model DSV4`可只读生成计划，不让用户提供列字母或技术JSON。
5. `execute PLAN.json`预检，`execute PLAN.json --apply`执行明确项。预检只输出独立审计目录，不改问题目录/原结果表，失败DOCX样本仅在审计目录。异常汇总，明确项继续；以实际copy/excel状态判断完成，不把planned当完成。

## 新版计划示例

路径/哈希是占位，替换为实际值。题号用字符串。object.display_name可选完整显示名，避免重复拼版本。对象、模型、版本、前缀来自本次输入。

```json
{
  "schema_version": 2,
  "dataset_root": "/data/常规能力测试数据集",
  "batch_id": "workspace-20261014",
  "unattended": true,
  "deadline_mode": true,
  "cutoff_at": "2026-10-15T10:00:00+08:00",
  "object": {
    "name": "MobileWork", "version": "1009", "display_name": "MobileWork1009",
    "model": "DSV4", "prefix": "MobileWork1009"
  },
  "read_roots": ["/projects/mobilework-batch"],
  "discovery_complete": true,
  "discovery_evidence": "精准工作空间全部根会话、正文、素材及截止状态的实际查询来源",
  "excel": {
    "path": "/data/常规能力测试数据集/常规能力测试结果记录.xlsx",
    "sha256": "读取时实际SHA256",
    "object_slot": 1
  },
  "tasks": [{
    "task_id": "W1",
    "question_dir": "/data/常规能力测试数据集/维度1/问题1",
    "requirements": ["完整可读的最终报告"],
    "excluded_candidates": [],
    "attempts": [
      {
        "session_id": "first-session",
        "submitted_at": "2026-10-15T08:00:00+08:00",
        "submission_evidence": "实际提交消息ID和时间，包含时区",
        "match_evidence": "正文与上传源素材匹配证据",
        "status": "success",
        "completed_at": "2026-10-15T08:05:00+08:00",
        "duration": "5m0s", "duration_evidence": "客户端完成状态时长",
        "completeness": "complete", "validation_evidence": "实际打开，题目要求已齐全",
        "artifacts": [{
          "session_id": "first-session",
          "source": "/projects/mobilework-batch/project1/报告.docx",
          "sha256": "源文件实际SHA256",
          "delivery_channel": "formal_path",
          "delivered_at": "2026-10-15T08:05:00+08:00",
          "evidence": "正式答复给出该路径；文件真实存在且完整可读，无下载组件"
        }]
      },
      {
        "session_id": "second-session",
        "submitted_at": "2026-10-15T08:10:00+08:00",
        "submission_evidence": "第二次实际提交记录",
        "match_evidence": "同题同素材",
        "status": "running", "status_at_cutoff": "running",
        "cutoff_state_evidence": "根会话在统一截止仍运行的消息/状态记录",
        "delivery_search_complete": true, "completeness": "none",
        "official_output_complete": false,
        "output_snapshot_complete": true, "official_output": []
      }
    ]
  }]
}
```

此例采用第一次，Excel填“一次运行成功”，第二次超时仅审计，不生成失败交付。

## 截止字段与状态

- submitted_at是带时区真实提交时间；started_at可为实际执行开始，不用项目创建时间。数组顺序错误或不明时待确认。
- status_at_cutoff是截止时根会话的可验证原始状态，cutoff_state_evidence给出处，不是稍后看到的最新状态。若结束于截止之后，根据消息时间重建截止快照，不能当作截止前完成。
- 截止仍running，先查正式组件、正式路径及项目最终文件才能设delivery_search_complete true。无实物用completeness none，明确仅部分用incomplete，未查清用unknown；无法读取不等于没有。
- running + none/incomplete由deadline.py归一failed并添加超时原因及截止耗时。客户端值有duration_as_of与cutoff一致且duration_evidence充分才优先，否则算真实开始/提交至截止的经过时长。明确暂停可用pause_seconds及pause_evidence扣除。开始/暂停等计时证据异常时，清除过时耗时，保留failed并添加duration_issue及“用时无法确认”，不因计时错误漏记已确认失败。
- running + complete可成功，但须全部artifacts、validation_evidence和delivery_at（全部正式交付完成时间）不晚于cutoff。任意文件不够。已结束成功仍按客户端执行时长，不用截止时间计算。
- 截止后才正式交付的artifact不纳入本次成功，保留排除说明；同路径被截止后版本覆盖而无法取得原版本则记异常。不根据mtime推断交付时间。
- 状态未知仍unknown；首次success优先，额外运行超时不改变结论。两次超时或失败+超时归both_failed。仅一次失败不虚构第二次。未执行attempts为空，完整发现范围有证据才能判未执行。

已确认截止仍运行、交付核查完整的会话仍写入attempts，不因没有下载文件、回答未写完或没有最终完成事件而将整题设pending或漏列；这些情形按截止规则处理。真正归属/证据不明时才保留待确认。

## 正式输出和超时 DOCX

两次失败只归档第二次界面输出DOCX。已结束失败沿用official_output_complete true表示正式输出已完整查清。超时运行无须等到回答完成：第二次run给output_snapshot_complete true表示**截至cutoff已经显示的正式内容读全**，此时official_output_complete可以为false。旧计划的official_output_complete true仍兼容。按界面顺序列块，排除thinking/reasoning，不靠关键词删正文；不能将截止后新输出加入快照。

```json
{
  "official_output_complete": true,
  "official_output": [
    {"kind": "text", "visible": true, "text": "原样正式答复及可见错误，不写摘要", "evidence": "正式消息/partID"},
    {"kind": "table", "visible": true, "rows": [["原表头1", "原表头2"], ["值1", "值2"]], "evidence": "正式表格出处"},
    {"kind": "image", "visible": true, "source": "/projects/project1/正式图表.png", "sha256": "实际SHA256", "evidence": "正式图像出处"}
  ]
}
```

实际没有的表格/图片不添加。超时且已确认没有正式回答可用空数组和output_snapshot_complete true，脚本在DOCX独立添加“收集核查记录”，忠实记录截止时间、会话仍运行、无正式输出/无完整交付、耗时，不把描述伪装成受测对象回答。超时快照只读取到部分时设output_snapshot_complete false，仍列已取得且有可见性/来源证据的块，保存可取得内容与缺口说明；无法取得任何块也只记核查事实和读取缺口，不能声称没有输出。非超时失败、无可见内容时仍保留证据缺口，不编造正文。

both_failed先安排已确认的Excel结论，再生成/复制DOCX；导出异常只影响文件收集状态，Excel仍登记两次情况并追加快照异常说明。表格自身被锁定/变动、题号映射不明等保护仍生效，不能为了登记而绕过。核对每题excel_status和collection_status，失败结论、文件是否保存、快照是否读全分开核查，不能因没有下载文件将整题跳过。

生成后对照源检查文字/表格/图片，有渲染工具时检查页面。失败部分附件不默认归档。

## 文件与执行限制

artifact含source绝对路径、session_id、sha256、evidence，可选delivery_channel（formal_path/download_component等）、delivered_at。不能跨attempt拼接，成功须全部必需文件；缺组件不是缺交付。项目文件增加project_name、relative_path，内部原名/路径保留；逐一列依赖，不复制整个工作区。

每次独立审计目录保留计划、截止快照、选取理由、解析后的Excel位置、逐文件哈希和操作结果。预检/执行复用cutoff而不复用审计目录。同名同哈希跳过，不同内容冲突；不改源文件。结果表先备份，SHA/expected变化重新读取定位，按excel.md有限重试，不覆盖其它对象。

新计划优先object_slot；旧式sheet/task_id_column/time_column/note_column/headers兼容。非标准布局可由agent核实后给显式映射或使用可保留布局的工具，不绕过保护或公式写入。清单长文本超Excel限制时标完整见JSON，JSON保留全量；人工核查栏留空。
