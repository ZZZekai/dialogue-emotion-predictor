# SQL Dialogue Converter

将 SQL 文件中的审讯对话记录解析成便于人工查看、筛选和后续情绪预测的 Excel 工作簿。

这个工具只读取 SQL 文本，不执行 SQL、不连接数据库，也不会修改原始 SQL 文件。

## 主要功能

- 从 `INSERT INTO ... VALUES (...)` 语句中读取 `chat_id`、`message` 和 `create_time`。
- 正确处理 SQL 字符串中的逗号、分号、换行和转义单引号。
- 根据“审调人员”和“被谈话人”角色标记拆分问答，不简单按分号切割。
- 每个 `chat_id` 生成一个独立 Excel 工作表。
- 同一个 `chat_id` 的多条 SQL 记录按时间合并，轮次连续编号。
- 第一张工作表生成 `index`，集中展示对话长度、污染标记和复核优先级。
- 保留问题缺回答、回答缺问题等不完整轮次，不自动删除原始内容。
- 单条 SQL 记录解析失败时记录错误并继续处理后续记录。
- 额外生成 JSON 汇总，便于统计和程序读取。

## 运行环境

- Python 3.10 或更高版本
- `openpyxl`
- Windows、macOS 和 Linux 均可运行
- 不需要安装 Excel、LibreOffice 或数据库客户端

## 目录结构

```text
sql_dialogue_converter/
├── input/
│   └── talk_message.sql          # 待转换的 SQL 文件
├── output/
│   ├── dialogue_review.xlsx      # 转换后的主要 Excel 文件
│   └── dialogue_screening_summary.json
├── clean_sql_data.py             # 主程序
├── test_clean_sql_data.py        # 自动化测试
├── requirements.txt              # 独立依赖清单
└── README.md
```

`input/` 和 `output/` 可以为空。程序运行时会自动创建不存在的输出目录。

## 安装方法

### 方式一：直接安装依赖

进入工具目录：

```bash
cd sql_dialogue_converter
```

安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

Windows 如果没有 `python3` 命令，可以使用：

```powershell
python -m pip install -r requirements.txt
```

### 方式二：使用虚拟环境

macOS 或 Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell：

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 快速使用

1. 将 SQL 文件放入 `input/`，例如：

```text
input/talk_message.sql
```

2. 在 `sql_dialogue_converter/` 目录运行：

```bash
python3 clean_sql_data.py input/talk_message.sql --output-dir output
```

Windows：

```powershell
python clean_sql_data.py input\talk_message.sql --output-dir output
```

3. 处理完成后查看：

```text
output/dialogue_review.xlsx
output/dialogue_screening_summary.json
```

终端会输出解析记录数、`chat_id` 数量、总轮数、失败数量和最终文件路径。程序不会把完整对话打印到终端。

## 命令行参数

查看帮助：

```bash
python3 clean_sql_data.py --help
```

标准格式：

```text
python3 clean_sql_data.py <输入SQL> [--output-dir 输出目录] [--output-file Excel文件名]
```

| 参数 | 是否必填 | 默认值 | 说明 |
|---|---|---|---|
| `input` | 是 | 无 | 输入 SQL 文件路径 |
| `--output-dir` | 否 | `output` | Excel 和 JSON 的输出目录 |
| `--output-file` | 否 | `dialogue_review.xlsx` | Excel 输出文件名，必须以 `.xlsx` 结尾 |

自定义 Excel 文件名：

```bash
python3 clean_sql_data.py input/talk_message.sql \
  --output-dir output \
  --output-file dialogue_review_2026.xlsx
```

路径中包含空格时应使用引号：

```bash
python3 clean_sql_data.py "input/my dialogue.sql" --output-dir "my output"
```

## 支持的 SQL 格式

程序主要面向包含明确列名的 `INSERT` 语句，例如：

```sql
INSERT INTO `talk_message` (`chat_id`, `message`, `create_time`) VALUES (
  '1000000000000000001',
  '审调人员：你叫什么名字？；被谈话人：我叫陈某。',
  '2026-04-22 08:20:45'
);
```

必要字段：

| 字段 | 说明 |
|---|---|
| `chat_id` | 一组对话的唯一标识，也是工作表名称的主要来源 |
| `message` | 包含角色标记的完整对话文本 |
| `create_time` | SQL 记录时间，用于合并同一 `chat_id` 的多条记录 |

字段顺序可以不同，也可以存在其他字段，但必须在 `INSERT` 列名列表中包含以上三个字段。

程序支持：

- 中文逗号、分号、句号和换行。
- 英文逗号、分号和冒号。
- SQL 标准转义单引号，例如 `O''Brien`。
- 常见反斜杠转义，例如 `\n`、`\r` 和 `\t`。
- 一条 `INSERT` 中包含一个或多个 `VALUES` 元组。
- SQL 字符串内部包含普通逗号或分号。

文件必须使用 UTF-8 编码。如果 UTF-8 解码失败，程序会给出明确错误，不会猜测其他编码。

## 对话拆分规则

程序识别以下四种角色标记：

```text
审调人员：
审调人员:
被谈话人：
被谈话人:
```

例如：

```text
审调人员：问题1；被谈话人：回答1。审调人员:问题2;被谈话人:回答2。
```

会被拆分为：

| 轮次 | 提问人员 | 被谈话人 |
|---:|---|---|
| 1 | 问题1 | 回答1。 |
| 2 | 问题2 | 回答2。 |

角色内容内部的分号、逗号、引号和换行会保留。程序优先使用角色标记的位置确定边界。

异常轮次处理：

- 只有问题没有回答：保留该轮，回答列为空。
- 开头只有被谈话人回答：单独形成一轮，问题列为空。
- 连续出现多个问题：前一个无回答问题仍会保留。
- 没有任何可识别角色标记：该条记录不会生成对话轮次，但 SQL 记录本身仍算解析成功。

## Excel 输出结构

### `index` 工作表

`index` 始终是第一张工作表，每行对应一个 `chat_id`。

| 字段 | 含义 |
|---|---|
| `sheet_name` | 实际生成的 Excel 工作表名称 |
| `chat_id` | SQL 中的原始对话 ID |
| `first_create_time` | 最早一条 SQL 记录时间 |
| `last_create_time` | 最晚一条 SQL 记录时间 |
| `question_count` | 识别到的问题数量 |
| `answer_count` | 识别到的回答数量 |
| `complete_turn_count` | 同时有问题和回答的完整轮数 |
| `turn_count_mismatch` | 问题数量与回答数量是否不同 |
| `dialogue_group` | 根据完整轮数生成的长度分组 |
| `ai_identity_contamination` | 是否出现 AI 身份相关内容 |
| `device_instruction_contamination` | 是否出现设备或操作说明 |
| `high_repetition` | 是否存在相同规范化问题至少重复3次 |
| `possible_transcription_issue` | 是否存在严重问答不匹配或明显连续重复字 |
| `review_priority` | 建议人工复核优先级 |
| `contamination_flags` | 将所有命中的标记合并为中文文本 |

`index` 按完整问答轮数从高到低排列；轮数相同时按最早记录时间从早到晚排列。后续对话工作表也按同样顺序创建。

### 对话工作表

每个 `chat_id` 对应一张独立工作表，并且只包含三列：

```text
轮次 | 提问人员 | 被谈话人
```

同一 `chat_id` 出现多条 SQL 记录时：

- 先按 `create_time` 从早到晚排序。
- 将不同记录中的对话依次追加。
- 轮次连续编号，不重新从1开始。
- 完全重复的问答暂不删除。

Excel 工作表名称超过31个字符、包含非法字符或发生重复时，程序会自动生成合法且唯一的名称，并在 `index.sheet_name` 中保存实际名称。

格式设置包括：

- 表头加粗并使用浅蓝色背景。
- 所有单元格自动换行、顶部对齐。
- 冻结首行并开启自动筛选。
- 对话列使用较宽列宽。
- 根据文本长度估算行高，最大行高限制为120。

## 对话分组规则

`dialogue_group` 使用完整问答轮数分组：

| 分组 | 完整问答轮数 |
|---|---:|
| `long` | 大于或等于15 |
| `medium` | 8～14 |
| `short` | 4～7 |
| `very_short` | 小于或等于3 |

## 质量与污染标记

这些标记只用于辅助筛选，不会自动删除或修改原始对话。

### AI 身份污染

检测“语言模型”“通义千问”“作为AI”“作为人工智能”等关键词。

### 设备或操作说明污染

检测“按V键”“点击按钮”“发送给模型”“麦克风”“设备调试”“操作界面”等关键词。

### 高重复内容

对问题去除空格和常见标点并统一大小写。如果相同规范化问题出现至少3次，则标记为高重复。

### 疑似转写异常

采用较保守的简单规则，主要检测：

- 问答数量严重不一致。
- 中文字符出现明显连续重复。

该字段只是人工复核提示，不代表内容一定存在错误。

## 人工复核优先级

`review_priority` 当前规则：

- `high`：完整问答不少于8轮，并且未命中质量或污染标记。
- `medium`：完整问答不少于4轮，但存在轻微重复、问答不匹配或疑似转写问题；其他不属于高或低的情况也归入此组。
- `low`：完整问答不超过3轮，或存在明显 AI 身份污染，或存在设备操作说明污染。

优先级用于安排人工复核顺序，不会影响对话内容。

## JSON 汇总文件

`dialogue_screening_summary.json` 至少包含：

```json
{
  "total_records": 0,
  "total_chat_ids": 0,
  "total_turns": 0,
  "parse_failed_records": 0,
  "failed_records": [],
  "group_counts": {},
  "priority_counts": {},
  "contamination_counts": {},
  "turn_statistics": {
    "min": 0,
    "max": 0,
    "mean": 0,
    "median": 0
  }
}
```

失败记录只保存记录序号、估算行号和错误原因，不复制完整超长对话。

## 常见问题

### 提示 `No module named 'openpyxl'`

说明依赖尚未安装：

```bash
python3 -m pip install -r requirements.txt
```

### 提示“找不到输入 SQL 文件”

确认当前终端位于 `sql_dialogue_converter/`，并检查文件名和扩展名。Windows 可以在文件资源管理器中确认文件没有被保存为 `talk_message.sql.txt`。

### 提示“SQL 文件不是有效的 UTF-8 编码”

请先使用文本编辑器将 SQL 文件另存为 UTF-8，再重新运行。工具不会自动使用 GBK 等编码读取，以避免中文内容被错误转换。

### 输出 Excel 无法保存或提示权限错误

关闭已经打开的同名 Excel 文件，确认输出目录可写，然后重新运行。Windows 上 Excel 打开文件时通常会阻止程序覆盖该文件。

### 生成了 Excel，但没有对话轮次

检查 `message` 是否使用了工具支持的角色标记。如果原始文本使用“询问人”“回答人”等其他名称，需要修改程序中的 `ROLE_PATTERN`。

### 大文件运行较慢

程序会将解析结果和工作簿保存在内存中。SQL 记录很多、`chat_id` 数量很多时，生成包含大量工作表的 Excel 可能需要数分钟，并占用较多内存。运行期间不要重复启动同一任务。

## 当前格式假设与限制

程序不是通用 SQL 解析器，当前主要假设如下：

- 数据通过带列名的 `INSERT INTO ... VALUES ...` 写入。
- 必要列名固定为 `chat_id`、`message` 和 `create_time`，大小写不敏感。
- `create_time` 可以被 Python ISO 日期时间格式解析，例如 `2026-04-22 08:20:45`。
- 对话角色固定为“审调人员”和“被谈话人”。
- SQL 文件使用 UTF-8 编码。

以下情况可能无法正确解析：

- 使用 `UPDATE`、`COPY`、数据库二进制备份或压缩备份保存数据。
- `message` 使用十六进制、Base64 或数据库专用函数编码。
- SQL 引号、括号已经损坏或没有闭合。
- 使用数据库厂商特有的复杂字符串语法。
- 列名被改成其他名称。
- 角色名称被替换成工具未识别的名称。
- 语音转写异常但没有触发当前的简单规则。

无法解析的 `INSERT` 会记录在 JSON 的 `failed_records` 中，其他记录仍会继续处理。

## 运行测试

在工具目录运行：

```bash
python3 -m unittest -v test_clean_sql_data.py
```

测试覆盖普通多轮对话、多 `chat_id`、同一对话按时间合并、不完整问答、中文和英文冒号、SQL 转义、污染标记、重复问题、非法工作表名称和异常记录不中断。

## 打包给其他人

建议压缩包保留以下内容：

```text
sql_dialogue_converter/
├── input/
├── output/
├── clean_sql_data.py
├── test_clean_sql_data.py
├── requirements.txt
└── README.md
```

如果 SQL 和输出结果包含真实对话或敏感材料，打包前应清空 `input/` 和 `output/`。不要把真实 SQL、转换后的 Excel 或其他业务材料放进公开仓库或发送给无权限人员。

同学收到压缩包后的最短操作流程：

```bash
python3 -m pip install -r requirements.txt
python3 clean_sql_data.py input/talk_message.sql --output-dir output
```
