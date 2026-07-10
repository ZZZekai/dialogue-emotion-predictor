# dialogue-emotion-predictor

一个用于对话情绪预测的 Python 工具。项目读取 Excel 中的逐轮对话数据，调用通义百炼 / DashScope OpenAI-compatible API，结合人物背景、场景背景、few-shot 示例、历史对话和当前问题，预测回复方在回答当前问题时的主导情绪标签，并将结果写入新的 Excel 文件。

本仓库是公开安全版本：代码、Prompt 模板和脱敏示例可以提交；真实数据、API Key、人物材料、场景材料和模型输出默认不会提交。

## Project Structure

```text
dialogue-emotion-predictor/
├── data/
│   └── .gitkeep
├── outputs/
│   └── .gitkeep
├── models/
│   └── .gitkeep
├── prompts/
│   ├── emotion_prompt.txt
│   ├── emotion_examples.example.txt
│   ├── person_profile.example.txt
│   └── case_background.example.txt
├── scripts/
│   ├── predict_emotion.py
│   └── utils.py
├── .env.example
├── .gitignore
├── requirements.txt
└── README.md
```

`data/`、`outputs/`、`models/` 通过 `.gitkeep` 保留目录结构。真实输入数据、预测结果和本地模型文件由 `.gitignore` 忽略。

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
