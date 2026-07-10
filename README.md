# InterrogationAgent

读取 Excel 中的审讯对话数据，调用通义百炼 / DashScope OpenAI-compatible API，根据人物背景、案件背景、few-shot 示例、对话历史和当前问题，预测被谈话人在回答当前问题时的主导情绪标签。

## 项目结构

```text
InterrogationAgent/
├── data/
├── outputs/
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

## 安装依赖

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 配置 .env

```bash
cp .env.example .env
```

编辑 `.env`：

```text
DASHSCOPE_API_KEY=你的真实APIKey
DASHSCOPE_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
```

API Key 只从 `.env` 或系统环境变量读取，不会写死在代码中。

## Prompt 文件

仓库只保存脱敏模板文件：

```text
prompts/emotion_examples.example.txt
prompts/person_profile.example.txt
prompts/case_background.example.txt
```

实际运行前，复制模板并填写本地真实内容：

```bash
cp prompts/person_profile.example.txt prompts/person_profile.txt
cp prompts/case_background.example.txt prompts/case_background.txt
cp prompts/emotion_examples.example.txt prompts/emotion_examples.txt
```

脚本运行时会读取以下本地文件：

```text
prompts/emotion_prompt.txt
prompts/emotion_examples.txt
prompts/person_profile.txt
prompts/case_background.txt
```

`person_profile.txt`、`case_background.txt`、`emotion_examples.txt` 默认被 `.gitignore` 忽略，避免提交真实人物、案件和示例材料。

`emotion_prompt.txt` 必须包含：

```text
{person_profile}
{case_background}
{emotion_examples}
{dialogue_history}
{current_question}
```

## Excel 输入要求

支持两种输入格式。

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

脚本也兼容当前数据中的列名：

```text
提问人员
被谈话人
```

如果没有 `对话历史`，脚本会自动构造历史：

```text
第 N 行 history = 第 1 行到第 N-1 行的完整问答
```

当前第 N 行的被谈话人回答不会进入 `dialogue_history`。如果 Excel 自带 `对话历史` 且里面包含当前行回答，脚本会报错拦截，避免数据泄漏。

## 运行

```bash
python3 scripts/predict_emotion.py \
  --input data/dialogue.xlsx \
  --output outputs/emotion_result.xlsx \
  --model qwen-plus
```

如果你的 API Key 只授权了其他模型，例如 `qwen3-8b`：

```bash
python3 scripts/predict_emotion.py \
  --input data/dialogue.xlsx \
  --output outputs/emotion_result.xlsx \
  --model qwen3-8b
```

## 测试前 10 条

```bash
python3 scripts/predict_emotion.py \
  --input data/dialogue.xlsx \
  --output outputs/emotion_result_test.xlsx \
  --model qwen-plus \
  --limit 10
```

## 输出列

输出 Excel 会保留原表字段，并新增：

```text
预测情绪标签
原始情绪标签
预测原因
原始模型输出
```

其中 `预测情绪标签` 是后处理后的标签，`原始情绪标签` 是模型直接输出的标签。

## 临时保存

脚本每处理 20 条会自动保存一次：

```text
outputs/temp_result.xlsx
```

用于避免中断后完全丢失结果。

## 可选情绪标签

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

如果模型输出的标签不在以上列表中，脚本会把情绪标记为：

```text
PARSE_ERROR
```

并保留 `原始模型输出` 方便排查。
