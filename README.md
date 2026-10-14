# Delivery-Collector

用于测评后的交付收集：一次处理一个受测对象，自动识别题目与运行记录，选择有效运行，整理文件并填写数据集目录中的测试结果 Excel。支持 AlphaData、DeepInsight、常规能力，也可按实际题目扩展。

## 功能与实现

- 根据提交题目、源素材、会话、工作空间及批次识别任务，不依赖窗口改名。
- 第一次实际成功则采用第一次；第一次失败、第二次成功则采用第二次。
- 两次失败时，第二次正式界面输出（排除思考过程）保存为 DOCX，作为该题交付，Excel 记录两次情况。
- 文件按“对象标识-题号-原文件名”保存到各问题文件夹；网站保留内部目录。
- 自动生成发现记录、文件核查清单和待确认项。未执行题目不阻碍其他题目。
- Agent 负责证据匹配、完整性判断；Python 脚本负责只读发现、哈希核验、复制、DOCX 和 Excel 定点更新。规则与操作系统无关，路径和应用记录位置按实际环境识别。

需要 Agent 能访问应用记录、本地文件并执行 Python。普通操作只用 Python 标准库；失败输出 DOCX 需要 `python-docx`。不对交付评分，不修补受测对象的答案。自动执行结束后仍需人工核查。

## 文件作用

| 文件 | 作用 |
|---|---|
| `delivery-collector/SKILL.md` | 技能入口、背景、执行流程 |
| `delivery-collector/references/rules.md` | 运行选择、命名和 Excel 规则 |
| `delivery-collector/references/datasets.md` | 三类数据集的交付核查参考 |
| `delivery-collector/references/discovery.md` | WorkBuddy、MobileWork 证据读取与匹配 |
| `delivery-collector/references/plan.md` | Agent 与执行脚本之间的 JSON 约定 |
| `delivery-collector/scripts/collector.py` | 发现、预检、收集、清单、记录更新 |
| `delivery-collector/scripts/xlsx_io.py` | Excel 读取、核查表输出、保留格式的定点修改 |
| `delivery-collector/agents/openai.yaml` | 显示名称和调用提示 |
| `tests/` | 隔离样例与回归测试，不包含真实测试数据 |

## 安装

在支持技能的 AI Agent 中输入：

```text
请从 https://github.com/Ma-Ruinan/Delivery-Collector 安装技能，
技能目录为 delivery-collector。请确认安装后可以调用 Delivery-Collector。
```

由 Agent 根据自身的技能机制安装并检查执行依赖。内部技能标识为 `delivery-collector`，显示名称为 `Delivery-Collector`。

## 使用示例

```text
使用 Delivery-Collector，自动收集本批次 WorkBuddy 的测试交付。

数据集：F:\综合测试数据集\deepinsight测评数据集
受测对象：WorkBuddy-5.5.6，模型 glm-5.2
文件名前缀：workbuddy5.5.6
项目根目录：D:\APPDATA\workbuddy
批次：指定工作空间，2026-10-14 提交的任务
Excel：数据集文件夹中的 DeepInsight测试结果记录.xlsx
填写对象：测试对象2，表头按上述版本和模型更新。

自行识别实际执行题目和两次运行，按技能规则选择结果。
交付保存到对应问题文件夹，填写该对象的用时及说明。
输出发现记录、文件核查清单和待确认项，其余对象不要修改。
```

收集 MobileWork 时，在新窗口另行调用，替换对象、前缀、项目路径及批次信息即可。也可以指定“仅维度2、4”或“只发现、输出清单，不复制、不修改 Excel”。批次不明确时 Agent 会先询问；单题歧义不阻碍其他明确任务。

测试：`python -m unittest discover -s tests -v`。
