# 下游对话生成验证

该模块直接读取原始逐轮对话，使用同一个 DashScope OpenAI-compatible API 完成
端到端下游验证。每一轮依次执行：

1. 使用 `qwen3-8b` 实时预测新的情绪标签；
2. 使用 `deepseek-v4-flash` 生成不含情绪信息的基线回答；
3. 使用 `deepseek-v4-flash` 结合本轮新标签生成情绪增强回答。

两种模拟回答分别为：

1. 无情绪模拟回答：不向模型提供情绪预测；
2. 情绪增强模拟回答：只向模型额外提供小模型输出的情绪标签。

`预测原因`会保留在输出 Excel 中用于分析，但不会发送给回答生成模型，避免原因文本
额外透露回答策略，使实验能够更准确地比较“是否加入情绪标签”这一单一变量。

输出 Excel 将真实回答、无情绪模拟回答和情绪增强模拟回答并排排列，便于人工比较。

## 实验目的与控制变量

该实验用于验证：**小模型预测的情绪标签，是否能让大模型生成的回答更接近真实谈话
中的表达方式。** 它不是让模型复述当前轮真实答案，也不是直接复用已经生成好的情绪
结果。

| 项目 | 无情绪基线组 | 情绪增强组 |
| --- | --- | --- |
| 情绪预测模型 | `qwen3-8b` 实时预测 | `qwen3-8b` 实时预测 |
| 回答生成模型 | `deepseek-v4-flash` | `deepseek-v4-flash` |
| 人物和案件背景 | 相同 | 相同 |
| 当前轮之前的真实历史 | 相同 | 相同 |
| 当前问题 | 相同 | 相同 |
| 情绪标签 | 不发送 | 发送预测标签 |
| 预测原因 | 不发送 | 不发送 |
| 生成 temperature | 默认 `0.4` | 默认 `0.4` |

因此，两组回答的主要实验变量只有“是否向生成模型提供预测情绪标签”。情绪预测自身
默认 temperature 为 `0.0`，用于降低标签随机性；回答生成默认 temperature 为 `0.4`，
在稳定性和自然表达之间取相对保守的平衡。

## 单轮数据流

以第 N 轮为例：

```text
第 1 至 N-1 轮真实问答 + 第 N 轮问题 + 对应人物背景
                         │
                         ▼
                 qwen3-8b 预测情绪
                         │
                ┌────────┴────────┐
                ▼                 ▼
        不带情绪生成回答      带情绪标签生成回答
                │                 │
                └────────┬────────┘
                         ▼
          与第 N 轮真实回答并排写入 Excel
```

第 N 轮真实回答只用于最终人工对比，不会作为本轮模型输入。两种模型生成回答也不会
进入第 N+1 轮历史，从而避免生成误差逐轮累积。

## 输入要求

输入应是原始逐轮对话 Excel，至少包含：

```text
提问人员（或审调人员问题、审调人员当前问题）
被谈话人（或被谈话人回答）
```

不需要提前运行情绪预测主程序。即使输入中存在旧预测列，脚本也会清空并重新生成，
不会复用旧标签。

三次模型调用都使用当前轮之前的真实问答构造历史。当前轮真实回答不会发送给
`qwen3-8b` 或 `deepseek-v4-flash`，模型生成回答也不会进入下一轮历史。

生成脚本不再按照工作表总轮数机械划分初期、中期和后期。每轮会根据此前真实回答和
当前问题动态生成对话状态，包括全局配合度、当前事项是否首次出现、证据强度、具体
金额时间、矛盾追问和建议披露策略。面对新敏感事项时，即使此前已经配合，也可以重新
谨慎试探；只有出现明确证据、持续追问或既有承认时才逐步增加披露。该计算不会增加
额外 API 请求。

为减少模拟回答过长、过度表演和背景信息泄漏，生成阶段还会执行以下约束：

- 根据此前最多 10 轮真实回答的长度和当前问题类型，动态给出回答篇幅建议；
- 程序性确认和简短事实问题通常只回答 1 句，细节追问才允许适当展开；
- 人名、金额、日期、地点、疾病、药物、资金用途和事件细节必须能从人物背景或此前
  真实对话中找到依据，依据不足时回答“不清楚”“记不太清”或“具体看记录”；
- 禁止输出“沉默片刻”“低下头”“声音发颤”等舞台动作，脚本还会对遗漏的动作描写
  做一次清理；
- 情绪默认以低到中等强度影响措辞，不把 nervousness、fear、remorse 等标签直接
  扩写成失控、立即认罪或长篇忏悔。

生成 Prompt 还会根据问题类型选择主要应答方式，覆盖程序性事实、残缺问句、指代
省略、重复提问、笼统证据、具体书证、他人证言、前后矛盾、法律定性和亲情劝导等
情况。决策优先级为“事实与既有口径 > 当前披露状态 > 应答策略 > 情绪 > 表达方式”。
情绪不能改变事实或越过当前披露状态。当前版本尚未接入正式环境中的长期记忆检索，
仍使用完整背景与真实历史。

## Prompt

当前模板位于：

```text
downstream/prompts/response_baseline_prompt.txt
downstream/prompts/response_emotion_prompt.txt
```

后续可以直接替换模板正文，但必须保留已有 `{...}` 占位符。

人物姓名、完整背景和工作表绑定统一配置在：

```text
local_profiles/profiles.json
```

配置采用“人物档案注册 + 输入文件/工作表绑定”结构：

```text
profiles：定义 profile_id、人物姓名和 .docx/.txt 背景文件
bindings：将 输入文件名::工作表名 绑定到 profile_id
```

格式参考项目根目录的 `profile_config.example.json`。同一档案会同时提供给 `qwen3-8b`
和 `deepseek-v4-flash`。`local_profiles/` 默认被 Git 忽略；未绑定工作表会直接报错，
避免静默套用错误人物背景。单次实验也可用 `--profile-id` 显式指定人物。

## 小规模测试

直接使用示例对话文件测试前 10 轮：

```bash
python downstream/scripts/generate_responses.py \
  --input data/example_dialogue.xlsx \
  --limit 10
```

默认情绪模型为 `qwen3-8b`，默认生成模型为 `deepseek-v4-flash`。输出自动保存到
`downstream/outputs/`，文件名包含日期时间，不会覆盖之前的结果。
使用 `--limit 10` 时，输出文件也只保留实际处理的前 10 行。
回答生成的默认 temperature 为 `0.4`，可通过 `--generation-temperature` 调整。

用于讲解或对照实验时，建议显式写出温度参数，方便复现实验：

```bash
python downstream/scripts/generate_responses.py \
  --input data/example_dialogue.xlsx \
  --limit 10 \
  --emotion-temperature 0.0 \
  --generation-temperature 0.4
```

温度参数的含义：

- `--emotion-temperature 0.0`：情绪标签更稳定，适合作为实验自变量；
- `--generation-temperature 0.4`：回答较克制，同时保留一定语言变化；
- 提高生成温度会增加表达多样性，也会增加事实偏移和不同批次结果差异。

指定模型或输出文件：

```bash
python downstream/scripts/generate_responses.py \
  --input data/example_dialogue.xlsx \
  --output downstream/outputs/example_dialogue_comparison.xlsx \
  --emotion-model qwen3-8b \
  --generation-model deepseek-v4-flash \
  --limit 10
```

限制真实历史为最近 20 轮：

```bash
python downstream/scripts/generate_responses.py \
  --input data/example_dialogue.xlsx \
  --history-turns 20 \
  --limit 10
```

处理多工作表文件时，可以跳过 `index` 并只测试前两个对话工作表。建议首次每表测试
5 轮，共处理 10 轮：10 次情绪预测加 20 次回答生成，总计 30 次 API 请求。

```bash
python downstream/scripts/generate_responses.py \
  --input tool/sql_dialogue_converter/output/dialogue_review.xlsx \
  --all-sheets \
  --sheet-limit 2 \
  --rows-per-sheet 5
```

每个工作表独立构造历史，切换工作表时历史和谈话阶段都会重置。输出工作簿保留两个
同名结果工作表。确认小规模结果后，去掉 `--rows-per-sheet` 才会完整处理两张表。

## 输出列

主要对比列依次为：

```text
提问人员
被谈话人
无情绪模拟回答
情绪增强模拟回答
预测情绪标签
原始情绪标签
预测原因
情绪模型原始输出
情绪预测状态
```

另外保留基线和情绪增强两列生成状态。API 调用成功时状态为 `OK`，连续三次失败时
记录 `API_ERROR`。如果小模型预测失败，基线回答仍会生成，情绪增强回答会跳过。

## 当前边界

- 当前属于受控的开放环逐轮验证，不是让生成回答持续影响后续对话的闭环模拟；
- 人物背景用于提供事实与利害关系，不应替代当前问题和邻近历史成为情绪依据；
- 当前尚未接入独立谈话策略库，披露程度主要由真实历史、问题类型、证据强度和情绪
  标签共同约束；
- 评价仍需结合真实回答进行人工分析，不能仅以语言更丰富或情绪更强作为效果更好。
