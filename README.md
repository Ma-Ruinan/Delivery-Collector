# Delivery-Collector

用于测评后的交付收集：一次处理一个受测对象，自动识别题目与运行记录，选择有效运行，整理文件并填写数据集目录中的测试结果 Excel。支持 AlphaData、DeepInsight、常规能力，也可按实际题目扩展。

## 功能与实现

- 根据提交题目、源素材、会话、工作空间及批次识别任务，不依赖窗口改名。
- 只统计客户端当前有效任务，已删除会话和残留目录不计入运行次数。
- 第一次实际成功则采用第一次；第一次失败、第二次成功则采用第二次。
- 正式输出提供最终文件路径、文件实际存在且满足交付要求，即可成功，不要求下载组件。
- 核对全部成功交付组件及正式路径，研究计划等正式附件也收集，多次交付不漏项。
- 两次失败时，第二次正式界面输出（排除思考过程）保存为 DOCX，作为该题交付，Excel 记录两次情况。
- 定时收集按统一截止时间判断超时；首次成功不受第二次超时影响。两次失败时保存第二次截至截止时的正式输出为DOCX，不等任务结束；Excel独立记录两次耗时及失败原因，Word保存异常不导致漏记。
- 按“测试对象1/2/3”自主更新表头和记录；列组不足时按原样式向右扩展，不改其他对象数据。
- 按完整同次运行登记用时，保留压缩前后各段依据；本次填写的表头、用时和说明统一居中。
- 文件按“对象标识-题号-原文件名”保存到各问题文件夹；网站保留内部目录。
- 自动生成发现记录、文件核查清单和待确认项。未执行题目不阻碍其他题目。
- Agent 负责证据匹配、完整性判断；Python 脚本负责只读发现、哈希核验、复制、DOCX 和 Excel 定点更新。规则与操作系统无关，路径和应用记录位置按实际环境识别。

需要 Agent 能访问应用记录、本地文件并执行 Python。普通操作只用 Python 标准库；失败输出 DOCX 需要 `python-docx`。不对交付评分，不修补受测对象的答案。自动执行结束后仍需人工核查。

## 文件作用

| 文件 | 作用 |
|---|---|
| `delivery-collector/SKILL.md` | 技能入口、背景、执行流程 |
| `delivery-collector/references/rules.md` | 运行选择、命名和 Excel 规则 |
| `delivery-collector/references/excel.md` | 无人值守对象定位和缺失列组扩展 |
| `delivery-collector/references/datasets.md` | 三类数据集的交付核查参考 |
| `delivery-collector/references/discovery.md` | WorkBuddy、MobileWork 证据读取与匹配 |
| `delivery-collector/references/plan.md` | Agent 与执行脚本之间的 JSON 约定 |
| `delivery-collector/scripts/collector.py` | 发现、预检、收集、清单、记录更新 |
| `delivery-collector/scripts/xlsx_io.py` | Excel 读取、核查表输出、保留格式的定点修改 |
| `delivery-collector/scripts/excel_layout.py` | 按对象序号定位列组、生成表头及扩展计划 |
| `delivery-collector/scripts/deadline.py` | 统一截止快照、超时分类与耗时口径 |
| `delivery-collector/scripts/evidence.py` | 有效会话筛选、全部组件交付提取、清单覆盖及分段用时核查 |
| `delivery-collector/agents/openai.yaml` | 显示名称和调用提示 |
| `tests/` | 隔离样例与回归测试，不包含真实测试数据 |

## 安装

在支持技能的 AI Agent 中输入：

```text
请从 https://github.com/Ma-Ruinan/Delivery-Collector 安装技能，
技能目录为 delivery-collector。请确认安装后可以调用 Delivery-Collector。
```

由 Agent 根据自身的技能机制安装并检查执行依赖。内部技能标识为 `delivery-collector`，显示名称为 `Delivery-Collector`。

已经安装时，让Agent从同一仓库更新该skill，再启动新的收集任务。

## 使用示例

```text
使用已安装的 skill [Delivery-Collector]，收集本次测试交付并填写结果记录。

数据集路径：D:\mobilework\wb-mw评测数据集\20261012\常规能力测试数据集
受测对象：MobileWork1009
模型：DSV4
文件名前缀：MobileWork1009
项目根目录：D:\mobilework
工作空间名（精准匹配）：20261014
Excel填写位置：测试对象1

本次为定时收集，以收集启动时间作为本批次测评截止时间。
按照skill规则自动完成运行结果选择、交付归档、Excel登记及核查清单输出，无需交互确认。
```

收集其他对象时，在新窗口替换对应参数；填写第三个对象时改为“Excel填写位置：测试对象3”。默认寻找数据集目录内的结果记录表，也可指定Excel路径、维度/题号，或要求“只发现，不复制、不修改Excel”。异常项会汇总，不停等选择，不阻碍其他明确任务。收集后仍须人工核查。

测试：`python -m unittest discover -s tests -v`。
