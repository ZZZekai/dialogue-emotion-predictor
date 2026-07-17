import argparse
import os
import re
import statistics
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from openai import OpenAI
from tqdm import tqdm


PROJECT_ROOT = Path(__file__).resolve().parents[2]
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

from utils import (  # noqa: E402
    ALLOWED_EMOTIONS,
    build_prompt as build_emotion_prompt,
    call_model_with_retry as call_emotion_model_with_retry,
    load_prompt_parts as load_emotion_prompt_parts,
    parse_emotion_response,
    postprocess_emotion,
)
from profile_loader import (  # noqa: E402
    Profile,
    apply_profile_to_emotion_prompt_parts,
    load_profile_registry,
)


DEFAULT_PROMPTS_DIR = PROJECT_ROOT / "downstream" / "prompts"
DEFAULT_EMOTION_PROMPTS_DIR = PROJECT_ROOT / "prompts"
DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "downstream" / "outputs"
DEFAULT_PROFILE_CONFIG = PROJECT_ROOT / "local_profiles" / "profiles.json"
DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_EMOTION_MODEL = "qwen3-8b"
DEFAULT_GENERATION_MODEL = "deepseek-v4-flash"

HISTORY_COL = "对话历史"
QUESTION_COL_CANDIDATES = ["审调人员当前问题", "审调人员问题", "提问人员"]
ANSWER_COL_CANDIDATES = ["被谈话人回答", "被谈话人"]
EMOTION_COL = "预测情绪标签"
RAW_EMOTION_COL = "原始情绪标签"
EMOTION_REASON_COL = "预测原因"
EMOTION_RAW_OUTPUT_COL = "情绪模型原始输出"
EMOTION_PREDICTION_STATUS_COL = "情绪预测状态"

BASELINE_OUTPUT_COL = "无情绪模拟回答"
EMOTION_OUTPUT_COL = "情绪增强模拟回答"
BASELINE_STATUS_COL = "无情绪生成状态"
EMOTION_GENERATION_STATUS_COL = "情绪增强生成状态"

AUTOSAVE_EVERY = 10
MAX_RETRIES = 3
RETRY_WAIT_SECONDS = 2

SENSITIVE_TOPIC_TERMS = [
    "礼品", "礼物", "现金", "红包", "购物卡", "烟卡", "借款", "利息",
    "购房", "房款", "优惠", "工程", "项目", "招标", "拆迁", "资金",
    "转账", "收受", "请托", "关照", "宴请", "吃饭", "旅游", "违纪",
    "责任", "好处", "回扣", "利益",
]
DIRECT_EVIDENCE_TERMS = [
    "转账记录", "银行流水", "聊天记录", "通话记录", "证人证言", "书证",
    "借条", "合同", "发票", "我们已经掌握", "我们掌握", "已经查明",
    "经查明", "经查", "记录显示", "证据显示", "已经核实", "交代了",
]
CONTRADICTION_TERMS = [
    "前后矛盾", "刚才还说", "之前说", "为什么不一致", "怎么解释",
    "不是说", "还不承认", "如实说明", "如实交代",
]
ADMISSION_TERMS = [
    "我承认", "确实收", "确实拿", "确实借", "确实有", "有这回事",
    "是我安排", "是我帮", "我同意了", "我接受了", "这件事是有的",
    "这个情况属实", "确实不合适",
]
REMORSE_TERMS = ["我后悔", "很后悔", "我愧疚", "我自责", "我认错", "反思责任"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run emotion prediction and downstream response generation end to end."
    )
    parser.add_argument("--input", required=True, help="Raw dialogue Excel path.")
    parser.add_argument(
        "--output",
        default=None,
        help="Output Excel path. A timestamped path is generated when omitted.",
    )
    parser.add_argument("--sheet", default="0", help="Sheet name or index. Default: 0.")
    parser.add_argument(
        "--all-sheets",
        action="store_true",
        help="Process dialogue worksheets independently and skip the index sheet.",
    )
    parser.add_argument(
        "--sheet-limit",
        type=int,
        default=None,
        help="With --all-sheets, process only the first N dialogue worksheets.",
    )
    parser.add_argument("--limit", type=int, default=None, help="Process only the first N rows.")
    parser.add_argument(
        "--rows-per-sheet",
        type=int,
        default=None,
        help="With --all-sheets, process at most N eligible rows in each worksheet.",
    )
    parser.add_argument(
        "--history-turns",
        type=int,
        default=None,
        help="Use only the most recent N real dialogue turns. Default: full history.",
    )
    parser.add_argument(
        "--emotion-model", default=DEFAULT_EMOTION_MODEL, help="Emotion prediction model."
    )
    parser.add_argument(
        "--generation-model",
        "--model",
        dest="generation_model",
        default=DEFAULT_GENERATION_MODEL,
        help="Response generation model. --model is retained as an alias.",
    )
    parser.add_argument("--emotion-temperature", type=float, default=0.0)
    parser.add_argument("--emotion-max-tokens", type=int, default=512)
    parser.add_argument("--generation-temperature", type=float, default=0.4)
    parser.add_argument("--generation-max-tokens", type=int, default=512)
    parser.add_argument("--env", default=str(PROJECT_ROOT / ".env"))
    parser.add_argument("--prompts-dir", default=str(DEFAULT_PROMPTS_DIR))
    parser.add_argument(
        "--emotion-prompts-dir",
        default=str(DEFAULT_EMOTION_PROMPTS_DIR),
        help="Directory containing the small-model emotion prompts.",
    )
    parser.add_argument(
        "--profile-config",
        default=str(DEFAULT_PROFILE_CONFIG),
        help="Shared local profile registry JSON.",
    )
    parser.add_argument(
        "--profile-id",
        default=None,
        help="Explicit profile override. Otherwise resolve by input filename and sheet.",
    )
    return parser.parse_args()


def read_text(path: str | Path) -> str:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Required text file not found: {file_path}")
    return file_path.read_text(encoding="utf-8").strip()


def validate_template(template: str, required: set[str], filename: str) -> None:
    missing = sorted(item for item in required if item not in template)
    if missing:
        raise ValueError(f"{filename} missing placeholders: {', '.join(missing)}")


def load_generation_templates(prompts_dir: str | Path) -> dict[str, str]:
    prompts_dir = Path(prompts_dir)
    baseline = read_text(prompts_dir / "response_baseline_prompt.txt")
    emotion = read_text(prompts_dir / "response_emotion_prompt.txt")
    common = {
        "{respondent_name}",
        "{story_background}",
        "{dialogue_history}",
        "{current_question}",
        "{dialogue_state}",
        "{response_length_guidance}",
    }
    validate_template(baseline, common, "response_baseline_prompt.txt")
    validate_template(
        emotion,
        common | {"{emotion}"},
        "response_emotion_prompt.txt",
    )
    return {
        "baseline_template": baseline,
        "emotion_template": emotion,
    }


def build_generation_materials(
    templates: dict[str, str], profile: Profile
) -> dict[str, str]:
    return {
        **templates,
        "respondent_name": profile.respondent_name,
        "story_background": profile.story_background,
    }


def create_client(env_path: str | Path) -> OpenAI:
    load_dotenv(env_path)
    api_key = os.getenv("DASHSCOPE_API_KEY")
    if not api_key:
        raise RuntimeError("Missing DASHSCOPE_API_KEY in .env or environment variables.")
    base_url = os.getenv("DASHSCOPE_BASE_URL", DEFAULT_BASE_URL)
    return OpenAI(api_key=api_key, base_url=base_url)


def find_column(df: pd.DataFrame, candidates: list[str], role: str) -> str:
    for candidate in candidates:
        if candidate in df.columns:
            return candidate
    raise ValueError(f"Missing {role} column. Supported names: {', '.join(candidates)}")


def validate_columns(df: pd.DataFrame) -> tuple[str, str]:
    question_col = find_column(df, QUESTION_COL_CANDIDATES, "question")
    answer_col = find_column(df, ANSWER_COL_CANDIDATES, "real answer")
    return question_col, answer_col


def to_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def build_real_history(
    df: pd.DataFrame,
    row_position: int,
    question_col: str,
    answer_col: str,
    history_turns: int | None,
) -> str:
    if HISTORY_COL in df.columns:
        history = to_text(df.iloc[row_position].get(HISTORY_COL, ""))
        current_answer = to_text(df.iloc[row_position].get(answer_col, ""))
        if current_answer and current_answer in history:
            raise ValueError(
                f"Row {row_position + 2}: current real answer appears in '{HISTORY_COL}'. "
                "Remove it to prevent leakage."
            )
        if history_turns is None:
            return history
        blocks = [block for block in history.split("\n\n") if block.strip()]
        return "\n\n".join(blocks[-history_turns:])

    start = 0 if history_turns is None else max(0, row_position - history_turns)
    lines: list[str] = []
    # The range ends before row_position, so the current real answer is never included.
    for previous_position in range(start, row_position):
        previous = df.iloc[previous_position]
        question = to_text(previous.get(question_col, ""))
        answer = to_text(previous.get(answer_col, ""))
        if question:
            lines.append(f"审调人员：{question}")
        if answer:
            lines.append(f"被谈话人：{answer}")
    return "\n".join(lines)


def build_prompts(
    materials: dict[str, str],
    dialogue_history: str,
    current_question: str,
    emotion: str,
    dialogue_state: str = "全局配合状态：尚未形成。\n当前事项状态：首次询问。",
    response_length_guidance: str = "优先使用1至2句话，简短回答当前问题。",
) -> tuple[str, str]:
    common = {
        "respondent_name": materials["respondent_name"],
        "story_background": materials["story_background"],
        "dialogue_history": dialogue_history or "（无此前对话）",
        "current_question": current_question,
        "dialogue_state": dialogue_state,
        "response_length_guidance": response_length_guidance,
    }
    baseline = materials["baseline_template"].format(**common)
    conditioned = materials["emotion_template"].format(
        **common,
        emotion=emotion,
    )
    return baseline, conditioned


def build_response_length_guidance(
    df: pd.DataFrame,
    row_position: int,
    answer_col: str,
    current_question: str,
    window: int = 10,
) -> str:
    start = max(0, row_position - window)
    previous_lengths = [
        len(answer)
        for answer in (
            to_text(df.iloc[position].get(answer_col, ""))
            for position in range(start, row_position)
        )
        if answer
    ]
    recent_median = (
        int(round(statistics.median(previous_lengths))) if previous_lengths else None
    )

    detailed_terms = [
        "详细", "说清楚", "具体经过", "怎么回事", "为什么", "原因",
        "资金来源", "如何", "过程", "全部情况",
    ]
    procedural_terms = [
        "听明白", "听清楚", "申请回避", "能否接受", "是否能接受",
        "有没有问题", "身体状况", "是不是", "是否",
    ]
    asks_for_detail = any(term in current_question for term in detailed_terms)
    is_short_confirmation = any(term in current_question for term in procedural_terms)

    if is_short_confirmation and not asks_for_detail:
        return "这是程序性确认或简短事实问题：使用1句话，通常控制在5至30字。"
    if asks_for_detail:
        upper = 90 if recent_median is None else min(120, max(60, recent_median * 3))
        return (
            "当前问题要求解释原因或经过：使用2至4句话，只说明被问到的事项，"
            f"通常控制在30至{upper}字。"
        )
    if recent_median is not None:
        lower = max(8, int(recent_median * 0.7))
        upper = min(80, max(lower + 15, int(recent_median * 1.6)))
        return (
            f"此前最近回答的中位长度约为{recent_median}字；延续其简洁程度，"
            f"优先使用1至2句话，通常控制在{lower}至{upper}字。"
        )
    return "这是一般事实问题：优先使用1至2句话，通常控制在15至50字。"


def extract_respondent_history(dialogue_history: str) -> str:
    """Extract prior respondent utterances so questions do not count as admissions."""
    answers: list[str] = []
    collecting_answer = False
    for line in dialogue_history.splitlines():
        stripped = line.strip()
        if stripped.startswith(("被谈话人：", "被谈话人:")):
            answers.append(stripped.split("：", 1)[-1] if "：" in stripped else stripped.split(":", 1)[-1])
            collecting_answer = True
        elif stripped.startswith(("审调人员：", "审调人员:")):
            collecting_answer = False
        elif collecting_answer and stripped:
            answers.append(stripped)
    return "\n".join(answers)


def count_term_hits(text: str, terms: list[str]) -> int:
    return sum(text.count(term) for term in terms)


def determine_dialogue_state(dialogue_history: str, current_question: str) -> str:
    """Build a turn-independent disclosure state from observed dialogue signals."""
    respondent_history = extract_respondent_history(dialogue_history)
    admission_hits = count_term_hits(respondent_history, ADMISSION_TERMS)
    remorse_hits = count_term_hits(respondent_history, REMORSE_TERMS)

    if remorse_hits:
        global_state = "已出现明确认错或责任反思，但不代表对所有事项都已充分交代"
    elif admission_hits >= 3:
        global_state = "已多次承认可核实事实，整体配合度有所提高"
    elif admission_hits:
        global_state = "已有限承认部分事实，仍保留解释和自我保护倾向"
    else:
        global_state = "尚未出现明确承认，整体保持试探、防御和侥幸心理"

    matched_topics = [term for term in SENSITIVE_TOPIC_TERMS if term in current_question]
    topic_repeated = any(term in dialogue_history for term in matched_topics)
    evidence_hits = [term for term in DIRECT_EVIDENCE_TERMS if term in current_question]
    contradiction_hits = [term for term in CONTRADICTION_TERMS if term in current_question]
    has_specific_detail = bool(
        re.search(r"\d+(?:\.\d+)?\s*(?:元|万|万元|年|月|日|次|套|笔|人)", current_question)
    )

    pressure_signals: list[str] = []
    if evidence_hits:
        pressure_signals.append("当前问题出现明确证据或核实性表述")
    if contradiction_hits:
        pressure_signals.append("当前问题指出矛盾或要求如实交代")
    if has_specific_detail:
        pressure_signals.append("当前问题包含具体金额、时间或数量")
    if topic_repeated:
        pressure_signals.append("相关事项此前已经出现，属于持续追问")
    if not pressure_signals:
        pressure_signals.append("当前问题未呈现明显直接证据或连续施压")

    if evidence_hits and (topic_repeated or admission_hits):
        issue_state = "相关事项受到明确证据约束，且已有追问或既有承认"
        strategy = "对无法回避的当前事实逐步松动，只承认可核实部分，同时保持事实边界"
    elif evidence_hits:
        issue_state = "当前事项首次或较少出现，但问题已经给出明确证据"
        strategy = "避免无依据硬性否认，先确认最小事实，再谨慎解释性质和动机"
    elif topic_repeated and (contradiction_hits or has_specific_detail):
        issue_state = "同一事项被具体化并持续追问，压力正在增加"
        strategy = "可以有限承认表层事实并解释，但不主动扩展到其他人物和事项"
    elif matched_topics and not topic_repeated:
        issue_state = "新的敏感事项首次被明确问及，审调人员掌握程度尚不清楚"
        strategy = "保持试探防御，淡化性质或要求具体说明，只披露不可回避的最小事实"
    elif topic_repeated:
        issue_state = "相关话题此前已经出现，但当前问题没有新增直接证据"
        strategy = "延续此前口径，谨慎补充必要信息，不主动扩大披露范围"
    else:
        issue_state = "普通或背景性询问，尚未形成明确敏感事项压力"
        strategy = "对基本事实自然回答；若问题转向责任或利益性质，再保持有限披露"

    return "\n".join(
        [
            f"- 全局配合状态：{global_state}",
            f"- 当前事项状态：{issue_state}",
            f"- 压力信号：{'；'.join(pressure_signals)}",
            f"- 建议应答策略：{strategy}",
        ]
    )


def predict_fresh_emotion(
    client: OpenAI,
    model: str,
    prompt_parts: dict[str, str],
    dialogue_history: str,
    current_question: str,
    temperature: float,
    max_tokens: int,
) -> tuple[str, str, str, str, str]:
    prompt = build_emotion_prompt(
        prompt_template=prompt_parts["prompt_template"],
        person_profile=prompt_parts["person_profile"],
        case_background=prompt_parts["case_background"],
        emotion_examples=prompt_parts["emotion_examples"],
        dialogue_history=dialogue_history,
        current_question=current_question,
    )
    try:
        raw_output = call_emotion_model_with_retry(
            client=client,
            model=model,
            prompt=prompt,
            max_retries=MAX_RETRIES,
            retry_wait_seconds=RETRY_WAIT_SECONDS,
            temperature=temperature,
            max_tokens=max_tokens,
        )
        raw_emotion, reason = parse_emotion_response(raw_output, ALLOWED_EMOTIONS)
        if raw_emotion == "PARSE_ERROR":
            return "PARSE_ERROR", raw_emotion, reason, raw_output, "PARSE_ERROR"
        final_emotion = postprocess_emotion(raw_emotion, reason)
        return final_emotion, raw_emotion, reason, raw_output, "OK"
    except Exception as exc:
        error = str(exc)
        return "API_ERROR", "API_ERROR", error, error, f"API_ERROR: {error}"


def call_generation_model(
    client: OpenAI,
    model: str,
    prompt: str,
    temperature: float,
    max_tokens: int,
) -> str:
    response = client.chat.completions.create(
        model=model,
        messages=[
            {
                "role": "system",
                "content": "你负责模拟被谈话人的自然回答，只输出回答正文。",
            },
            {"role": "user", "content": prompt},
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )
    return clean_generation_output(response.choices[0].message.content or "")


def clean_generation_output(text: str) -> str:
    cleaned = text.strip()
    action_terms = "沉默|叹气|低头|抬头|声音|语气|迟疑|停顿|思考|阅读|点头|摇头|发颤"
    cleaned = re.sub(
        rf"[（(][^）)]*(?:{action_terms})[^）)]*[）)]",
        "",
        cleaned,
    ).strip()
    cleaned = re.sub(
        rf"^(?:{action_terms})[^，。！？…]{{0,20}}[，。！？…]+",
        "",
        cleaned,
    ).strip()
    return cleaned


def call_with_retry(
    client: OpenAI,
    model: str,
    prompt: str,
    temperature: float,
    max_tokens: int,
) -> tuple[str, str]:
    last_error: Exception | None = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            output = call_generation_model(client, model, prompt, temperature, max_tokens)
            if not output:
                raise ValueError("Model returned an empty response.")
            return output, "OK"
        except Exception as exc:
            last_error = exc
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_WAIT_SECONDS)
    return "", f"API_ERROR: {last_error}"


def make_output_path(input_path: str | Path, output: str | None) -> Path:
    if output:
        path = Path(output)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = DEFAULT_OUTPUT_DIR / f"{Path(input_path).stem}_response_comparison_{timestamp}.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def reorder_comparison_columns(
    df: pd.DataFrame, question_col: str, answer_col: str
) -> pd.DataFrame:
    comparison = [
        question_col,
        answer_col,
        BASELINE_OUTPUT_COL,
        EMOTION_OUTPUT_COL,
        EMOTION_COL,
        RAW_EMOTION_COL,
        EMOTION_REASON_COL,
        EMOTION_RAW_OUTPUT_COL,
        EMOTION_PREDICTION_STATUS_COL,
        BASELINE_STATUS_COL,
        EMOTION_GENERATION_STATUS_COL,
    ]
    prefix = ["轮次"] if "轮次" in df.columns else []
    ordered = prefix + comparison
    remaining = [column for column in df.columns if column not in ordered]
    return df.loc[:, ordered + remaining]


def save_workbook(results: dict[str, pd.DataFrame], output_path: Path) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in results.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])


def process_sheet(
    source_df: pd.DataFrame,
    sheet_name: str,
    args: argparse.Namespace,
    client: OpenAI,
    materials: dict[str, str],
    emotion_prompt_parts: dict[str, str],
    row_limit: int | None,
    completed_results: dict[str, pd.DataFrame],
    output_path: Path,
) -> pd.DataFrame:
    question_col, answer_col = validate_columns(source_df)
    result_df = source_df.copy()
    for column in (
        EMOTION_COL,
        RAW_EMOTION_COL,
        EMOTION_REASON_COL,
        EMOTION_RAW_OUTPUT_COL,
        EMOTION_PREDICTION_STATUS_COL,
        BASELINE_OUTPUT_COL,
        EMOTION_OUTPUT_COL,
        BASELINE_STATUS_COL,
        EMOTION_GENERATION_STATUS_COL,
    ):
        # Always reset result columns so existing predictions can never be reused.
        result_df[column] = ""

    row_positions = [
        position
        for position in range(len(result_df))
        if to_text(source_df.iloc[position].get(question_col, ""))
        and to_text(source_df.iloc[position].get(answer_col, ""))
    ]
    if row_limit is not None:
        row_positions = row_positions[:row_limit]

    processed_positions: list[int] = []
    for row_position in tqdm(row_positions, desc=f"{sheet_name}"):
        row = source_df.iloc[row_position]
        current_question = to_text(row.get(question_col, ""))
        history = build_real_history(
            source_df, row_position, question_col, answer_col, args.history_turns
        )

        final_emotion, raw_emotion, reason, raw_emotion_output, emotion_prediction_status = (
            predict_fresh_emotion(
                client=client,
                model=args.emotion_model,
                prompt_parts=emotion_prompt_parts,
                dialogue_history=history,
                current_question=current_question,
                temperature=args.emotion_temperature,
                max_tokens=args.emotion_max_tokens,
            )
        )

        baseline_prompt, emotion_prompt = build_prompts(
            materials=materials,
            dialogue_history=history,
            current_question=current_question,
            emotion=final_emotion,
            dialogue_state=determine_dialogue_state(history, current_question),
            response_length_guidance=build_response_length_guidance(
                source_df, row_position, answer_col, current_question
            ),
        )
        baseline_answer, baseline_status = call_with_retry(
            client,
            args.generation_model,
            baseline_prompt,
            args.generation_temperature,
            args.generation_max_tokens,
        )
        if emotion_prediction_status == "OK":
            emotion_answer, emotion_generation_status = call_with_retry(
                client,
                args.generation_model,
                emotion_prompt,
                args.generation_temperature,
                args.generation_max_tokens,
            )
        else:
            emotion_answer = ""
            emotion_generation_status = "SKIPPED_EMOTION_PREDICTION_ERROR"

        index = result_df.index[row_position]
        result_df.at[index, EMOTION_COL] = final_emotion
        result_df.at[index, RAW_EMOTION_COL] = raw_emotion
        result_df.at[index, EMOTION_REASON_COL] = reason
        result_df.at[index, EMOTION_RAW_OUTPUT_COL] = raw_emotion_output
        result_df.at[index, EMOTION_PREDICTION_STATUS_COL] = emotion_prediction_status
        result_df.at[index, BASELINE_OUTPUT_COL] = baseline_answer
        result_df.at[index, EMOTION_OUTPUT_COL] = emotion_answer
        result_df.at[index, BASELINE_STATUS_COL] = baseline_status
        result_df.at[index, EMOTION_GENERATION_STATUS_COL] = emotion_generation_status
        processed_positions.append(row_position)

        if len(processed_positions) % AUTOSAVE_EVERY == 0:
            partial = reorder_comparison_columns(
                result_df.iloc[processed_positions].copy(), question_col, answer_col
            )
            save_workbook({**completed_results, sheet_name: partial}, output_path)

    return reorder_comparison_columns(
        result_df.iloc[processed_positions].copy(), question_col, answer_col
    )


def main() -> None:
    args = parse_args()
    if args.limit is not None and args.limit <= 0:
        raise ValueError("--limit must be a positive integer.")
    if args.history_turns is not None and args.history_turns <= 0:
        raise ValueError("--history-turns must be a positive integer.")
    if args.sheet_limit is not None and args.sheet_limit <= 0:
        raise ValueError("--sheet-limit must be a positive integer.")
    if args.rows_per_sheet is not None and args.rows_per_sheet <= 0:
        raise ValueError("--rows-per-sheet must be a positive integer.")
    if args.sheet_limit is not None and not args.all_sheets:
        raise ValueError("--sheet-limit requires --all-sheets.")
    if args.rows_per_sheet is not None and not args.all_sheets:
        raise ValueError("--rows-per-sheet requires --all-sheets.")
    if args.all_sheets and args.limit is not None:
        raise ValueError("Use --rows-per-sheet instead of --limit with --all-sheets.")

    input_path = Path(args.input)
    generation_templates = load_generation_templates(args.prompts_dir)
    profile_registry = load_profile_registry(args.profile_config)
    emotion_prompt_parts = load_emotion_prompt_parts(args.emotion_prompts_dir)
    client = create_client(args.env)
    output_path = make_output_path(input_path, args.output)
    excel_file = pd.ExcelFile(input_path)
    if args.all_sheets:
        sheet_names = [name for name in excel_file.sheet_names if name.lower() != "index"]
        if args.sheet_limit is not None:
            sheet_names = sheet_names[: args.sheet_limit]
        row_limit = args.rows_per_sheet
    else:
        sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
        sheet_name = excel_file.sheet_names[sheet] if isinstance(sheet, int) else sheet
        sheet_names = [sheet_name]
        row_limit = args.limit

    results: dict[str, pd.DataFrame] = {}
    for sheet_name in sheet_names:
        source_df = pd.read_excel(input_path, sheet_name=sheet_name)
        profile = profile_registry.resolve(
            input_path, sheet_name, profile_id=args.profile_id
        )
        sheet_materials = build_generation_materials(
            generation_templates, profile
        )
        sheet_emotion_prompt_parts = apply_profile_to_emotion_prompt_parts(
            emotion_prompt_parts, profile
        )
        print(
            f"Sheet {sheet_name}: profile={profile.profile_id}, "
            f"respondent={profile.respondent_name}, "
            f"background_sha256={profile.background_sha256[:12]}"
        )
        results[sheet_name] = process_sheet(
            source_df=source_df,
            sheet_name=sheet_name,
            args=args,
            client=client,
            materials=sheet_materials,
            emotion_prompt_parts=sheet_emotion_prompt_parts,
            row_limit=row_limit,
            completed_results=results,
            output_path=output_path,
        )
        save_workbook(results, output_path)

    print(f"Saved: {output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1)
