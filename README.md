# dialogue-emotion-predictor

一个用于对话情绪预测的 Python 工具。项目读取 Excel 中的逐轮对话数据，调用通义百炼 / DashScope OpenAI-compatible API，结合人物背景、场景背景、few-shot 示例、历史对话和当前问题，预测回复方在回答当前问题时的主导情绪标签，并将结果写入新的 Excel 文件。

本仓库是公开安全版本：代码、Prompt 模板和脱敏示例可以提交；真实数据、API Key、人物材料、场景材料和模型输出默认不会提交。

## 项目概览

本项目研究的问题是：**在不读取当前轮真实回答的前提下，能否根据谈话上下文预测
被谈话人的主导情绪，并利用该情绪改善下游模拟回答的真实性。**

系统包含三个主要环节：

```text
原始 SQL / Excel 对话
        │
        ▼
对话清洗与逐轮工作表整理
        │
        ▼
qwen3-8b 情绪预测
（人物背景 + 案件背景 + few-shot + 真实历史 + 当前问题）
        │
        ├──────────────┐
        ▼              ▼
无情绪基线回答      情绪增强回答
deepseek-v4-flash   deepseek-v4-flash + 情绪标签
        │              │
        └──────┬───────┘
               ▼
与当前轮真实回答并排输出到 Excel
```

下游对照实验中的两组回答使用相同的背景、历史、当前问题、模型和生成参数。两组之间
唯一的核心变量是：情绪增强组会额外收到小模型预测的情绪标签，基线组不会收到。

## 核心设计

- **逐轮预测**：预测的是“被谈话人在回答当前问题时可能出现的情绪”，不是读取已有
  回答后进行文本分类。
- **防止答案泄漏**：第 N 轮只使用第 1 至 N-1 轮真实问答；第 N 轮真实回答不会进入
  情绪模型或生成模型。
- **受控开放环实验**：模型生成的回答不会传入下一轮，下一轮仍使用真实历史，便于
  逐轮比较并控制误差传播。闭环模拟可作为后续独立实验。
- **工作表隔离**：多工作表代表不同对话时，每张表都会重置历史和对话状态。
- **人物档案绑定**：通过 `输入文件名::工作表名` 为不同场景绑定独立人物背景，情绪
  预测和下游生成共用同一份档案。
- **可复现实验参数**：情绪预测默认 temperature 为 `0.0`；下游回答生成默认
  temperature 为 `0.4`，也可以通过命令行覆盖。
- **事实边界控制**：下游生成不得凭空补充人物、金额、日期、地点等具体事实，缺少
  依据时应表达不确定。
- **表达约束**：回答长度会参考此前真实回答动态调整，并清理“沉默片刻”“低下头”
  等舞台动作，避免过度表演。

## 展示建议

向老师或同学介绍时，可以按照以下顺序演示：

1. 展示原始 Excel 的“提问人员”和“被谈话人”两列；
2. 说明情绪模型只能看到当前问题之前的真实历史；
3. 展示新增的预测标签、原始标签、预测原因和原始模型输出；
4. 展示下游 Excel 中并排的真实回答、无情绪回答和情绪增强回答；
5. 对比同一轮中两种模拟回答在措辞、配合程度和情绪表达上的差异；
6. 使用 Excel 合并工具比较全历史、20 轮历史或不同 Prompt 版本的实验结果。

## Project Structure

```text
dialogue-emotion-predictor/
├── data/
│   └── .gitkeep
├── outputs/
│   └── .gitkeep
├── models/
│   └── .gitkeep
├── local_profiles/
│   └── .gitkeep
├── downstream/
│   ├── prompts/
│   ├── scripts/
│   ├── outputs/
│   └── README.md
├── prompts/
│   ├── emotion_prompt.txt
│   ├── emotion_examples.example.txt
│   ├── person_profile.example.txt
│   └── case_background.example.txt
├── scripts/
│   ├── predict_emotion.py
│   ├── profile_loader.py
│   └── utils.py
├── tool/
│   ├── sql_dialogue_converter/
│   │   ├── input/
│   │   ├── output/
│   │   ├── clean_sql_data.py
│   │   ├── test_clean_sql_data.py
│   │   └── README.md
│   └── excel_comparison_merger/
│       ├── output/
│       ├── app.py
│       ├── merge_excel.py
│       ├── test_merge_excel.py
│       └── README.md
├── .env.example
├── .gitignore
├── profile_config.example.json
├── requirements.txt
└── README.md
```

`downstream/` 提供端到端下游验证：实时调用小模型预测情绪，再使用
`deepseek-v4-flash` 分别生成无情绪回答和情绪增强回答，最后与真实回答并排输出。
具体用法见 `downstream/README.md`。

`data/`、`outputs/`、`models/` 通过 `.gitkeep` 保留目录结构。真实输入数据、预测结果和本地模型文件由 `.gitignore` 忽略。

## Shared Profiles

本体情绪预测和下游回答生成共用 `local_profiles/profiles.json`。配置格式参考：

```text
profile_config.example.json
```

绑定键采用 `输入文件名::工作表名`，例如：

```text
example_dialogue.xlsx::对话整理
multi_dialogue.xlsx::case_sheet_001
```

每个档案包含人物姓名和完整 `.docx` 或 UTF-8 `.txt` 背景。背景会同时传给情绪模型
和下游生成模型，避免两阶段人物设定不一致。`local_profiles/` 包含敏感材料，已被
`.gitignore` 忽略。

## Installation

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Configuration

复制环境变量模板：

```bash
cp .env.example .env
```

编辑 `.env`：

```text
DASHSCOPE_API_KEY=your_api_key_here
DASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
```

不要把 `.env` 提交到 GitHub。API Key 只从 `.env` 或系统环境变量读取。

## Prompt Materials

仓库只保存脱敏模板：

```text
prompts/person_profile.example.txt
prompts/case_background.example.txt
prompts/emotion_examples.example.txt
```

运行前复制模板，并在本地填写真实内容：

```bash
cp prompts/person_profile.example.txt prompts/person_profile.txt
cp prompts/case_background.example.txt prompts/case_background.txt
cp prompts/emotion_examples.example.txt prompts/emotion_examples.txt
```

脚本运行时读取：

```text
prompts/emotion_prompt.txt
prompts/person_profile.txt
prompts/case_background.txt
prompts/emotion_examples.txt
```

其中 `person_profile.txt`、`case_background.txt`、`emotion_examples.txt` 默认被 `.gitignore` 忽略，避免提交真实人物、场景和示例材料。

主 Prompt 文件 `prompts/emotion_prompt.txt` 必须包含：

```text
{person_profile}
{case_background}
{emotion_examples}
{dialogue_history}
{current_question}
```

## Input Excel Format

输入 Excel 建议放在：

```text
data/
```

支持两种格式。

格式一：Excel 已经提供历史列：

```text
对话历史
审调人员当前问题
```

格式二：Excel 是逐轮问答表，没有历史列：

```text
审调人员问题
被谈话人回答
```

脚本也兼容更通用的逐轮对话列名：

```text
提问人员
被谈话人
```

如果没有 `对话历史`，脚本会自动构造历史：

```text
第 N 行 history = 第 1 行到第 N-1 行的完整问答
```

第 N 行的回复方回答不会进入 `dialogue_history`。如果 Excel 自带 `对话历史` 且里面包含当前行回答，脚本会报错拦截，避免数据泄漏。

## Usage

```bash
python3 scripts/predict_emotion.py \
  --input data/dialogue.xlsx \
  --output outputs/emotion_result.xlsx \
  --model qwen-plus
```

如果 API Key 只授权了其他模型，例如 `qwen3-8b`：

```bash
python3 scripts/predict_emotion.py \
  --input data/dialogue.xlsx \
  --output outputs/emotion_result.xlsx \
  --model qwen3-8b
```

测试前 10 条：

```bash
python3 scripts/predict_emotion.py \
  --input data/dialogue.xlsx \
  --output outputs/emotion_result_test.xlsx \
  --model qwen-plus \
  --limit 10
```

处理包含 `index` 和多个独立对话工作表的工作簿：

```bash
python3 scripts/predict_emotion.py \
  --input tool/sql_dialogue_converter/output/dialogue_review.xlsx \
  --output outputs/emotion_result.xlsx \
  --model qwen-plus \
  --all-sheets
```

多工作表模式会跳过 `index`，每张工作表独立构造历史。进入下一张工作表时历史会重置，当前行的被谈话人回答仍不会进入历史。输出文件保留原有工作表，并在每张对话工作表中增加预测列。

建议先对整个工作簿测试 10 条 API 请求：

```bash
python3 scripts/predict_emotion.py \
  --input tool/sql_dialogue_converter/output/dialogue_review.xlsx \
  --output outputs/emotion_result_test.xlsx \
  --model qwen-plus \
  --all-sheets \
  --limit 10
```

在 `--all-sheets` 模式中，`--limit` 是整个工作簿的预测总数，而不是每张工作表分别处理 10 条。

如需处理前 10 个对话工作表，并完整处理这些表中的所有轮次：

```bash
python3 scripts/predict_emotion.py \
  --input tool/sql_dialogue_converter/output/dialogue_review.xlsx \
  --output outputs/emotion_first10sheets.xlsx \
  --model qwen-plus \
  --all-sheets \
  --sheet-limit 10
```

`--sheet-limit` 限制工作表数量，`--limit` 限制 API 预测行数，两者含义不同，也可以同时使用。

默认情况下，输出文件名会自动追加运行时间，例如：

```text
emotion_result_20260713_140506.xlsx
```

如需覆盖或使用完全指定的文件名，可以增加 `--no-timestamp`。

限制每张工作表只使用最近 20 轮历史：

```bash
python3 scripts/predict_emotion.py \
  --input tool/sql_dialogue_converter/output/dialogue_review.xlsx \
  --output outputs/emotion_result_history20.xlsx \
  --model qwen-plus \
  --all-sheets \
  --history-turns 20
```

不传 `--history-turns` 时使用当前工作表的完整历史。历史窗口不会跨越工作表，也不会包含当前行的被谈话人回答。

## Downstream Validation

下游实验会在同一轮中实时执行一次情绪预测和两次回答生成，并将真实回答、无情绪
模拟回答、情绪增强模拟回答并排写入 Excel。测试前 10 轮：

```bash
python downstream/scripts/generate_responses.py \
  --input data/example_dialogue.xlsx \
  --limit 10 \
  --emotion-model qwen3-8b \
  --generation-model deepseek-v4-flash \
  --emotion-temperature 0.0 \
  --generation-temperature 0.4
```

输出会自动保存到 `downstream/outputs/`，文件名带日期时间。完整的实验控制变量、
人物档案绑定、多工作表运行和输出字段说明见
[downstream/README.md](downstream/README.md)。

## Compare Experiments

项目提供桌面工具，用于按同名工作表和相同轮次合并两份实验结果：

```bash
python3 tool/excel_comparison_merger/app.py
```

在界面中选择两个 Excel 文件并填写输出文件名。结果保存到：

```text
tool/excel_comparison_merger/output/
```

合并结果将实验1和实验2的情绪标签、原始标签、原因及原始模型输出成对排列，并用浅黄色标记预测情绪标签不同的轮次。

## Output Columns

输出 Excel 会保留原表字段，并新增：

```text
预测情绪标签
原始情绪标签
预测原因
原始模型输出
```

`预测情绪标签` 是后处理后的标签，`原始情绪标签` 是模型直接输出的标签。

输出文件建议写到：

```text
outputs/
```

脚本每处理 20 条会自动保存一次：

```text
outputs/temp_result.xlsx
```

## Emotion Labels

```text
neutral
confusion
disappointment
nervousness
anger
annoyance
disapproval
disgust
embarrassment
fear
sadness
remorse
```

如果模型输出的标签不在列表中，脚本会把情绪标记为：

```text
PARSE_ERROR
```

并保留 `原始模型输出` 方便排查。

## Public Repository Safety

以下内容默认不提交：

```text
.env
data/*.xlsx
data/*.xls
data/*.csv
data/*.tsv
outputs/*
models/*
prompts/person_profile.txt
prompts/case_background.txt
prompts/emotion_examples.txt
new_classification_standards/
```

提交前建议检查：

```bash
git add -n .
```

确认只包含代码、README、`.env.example`、Prompt 模板和 `.gitkeep` 占位文件。
