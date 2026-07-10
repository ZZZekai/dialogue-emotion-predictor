import argparse
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from utils import (
    ALLOWED_EMOTIONS,
    build_prompt,
    call_model_with_retry,
    create_openai_client,
    ensure_parent_dir,
    get_model_name,
    is_blank,
    load_prompt_parts,
    parse_emotion_response,
    postprocess_emotion,
    to_text,
)


# Column names are centralized here so they can be changed for new Excel formats.
HISTORY_COL = "对话历史"
QUESTION_COL = "审调人员当前问题"
FALLBACK_QUESTION_COL = "审调人员问题"
FALLBACK_ANSWER_COL = "被谈话人回答"

# Extra aliases keep compatibility with common transcript exports.
QUESTION_COL_CANDIDATES = [QUESTION_COL, FALLBACK_QUESTION_COL, "提问人员"]
ANSWER_COL_CANDIDATES = [FALLBACK_ANSWER_COL, "被谈话人"]

OUTPUT_EMOTION_COL = "预测情绪标签"
OUTPUT_RAW_EMOTION_COL = "原始情绪标签"
OUTPUT_REASON_COL = "预测原因"
OUTPUT_RAW_MODEL_COL = "原始模型输出"

TEMP_OUTPUT_PATH = Path("outputs/temp_result.xlsx")
AUTOSAVE_EVERY = 20
MAX_RETRIES = 3
RETRY_WAIT_SECONDS = 2


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Predict respondent emotions from interrogation dialogue Excel data."
    )
    parser.add_argument("--input", required=True, help="Input Excel path.")
    parser.add_argument("--output", required=True, help="Output Excel path.")
    parser.add_argument("--model", default=None, help="DashScope model name, e.g. qwen-plus.")
    parser.add_argument("--limit", type=int, default=None, help="Only process the first N rows.")
    parser.add_argument("--sheet", default=0, help="Excel sheet name or index. Default: first sheet.")
    parser.add_argument("--env", default=".env", help="Environment file path. Default: .env.")
    parser.add_argument("--prompts-dir", default="prompts", help="Prompt directory. Default: prompts.")
    parser.add_argument("--temperature", type=float, default=0.0, help="Model temperature.")
    parser.add_argument("--max-tokens", type=int, default=512, help="Max output tokens.")
    return parser.parse_args()


def parse_sheet_arg(sheet: str) -> str | int:
    return int(sheet) if str(sheet).isdigit() else sheet


def detect_column(df: pd.DataFrame, candidates: list[str], column_role: str) -> str:
    for column in candidates:
        if column in df.columns:
            return column
    raise ValueError(
        f"Missing {column_role} column. Supported column names: {', '.join(candidates)}"
    )


def detect_optional_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    return next((column for column in candidates if column in df.columns), None)


def validate_input_columns(df: pd.DataFrame) -> tuple[str, str | None]:
    question_col = detect_column(df, QUESTION_COL_CANDIDATES, "question")
    answer_col = detect_optional_column(df, ANSWER_COL_CANDIDATES)

    if HISTORY_COL not in df.columns and not answer_col:
        raise ValueError(
            f"Input Excel must contain '{HISTORY_COL}', or contain a respondent answer column "
            f"to construct history automatically. Supported answer columns: "
            f"{', '.join(ANSWER_COL_CANDIDATES)}"
        )

    return question_col, answer_col


def build_history_from_previous_rows(
    df: pd.DataFrame,
    row_position: int,
    question_col: str,
    answer_col: str,
) -> str:
    # Important: range(..., row_position) excludes the current row, so the current
    # respondent answer can never enter dialogue_history in auto-history mode.
    lines: list[str] = []
    for previous_position in range(0, row_position):
        previous_row = df.iloc[previous_position]
        previous_question = to_text(previous_row.get(question_col, ""))
        previous_answer = to_text(previous_row.get(answer_col, ""))

        if previous_question:
            lines.append(f"审调人员：{previous_question}")
        if previous_answer:
            lines.append(f"被谈话人：{previous_answer}")

    return "\n".join(lines)


def get_prompt_inputs(
    df: pd.DataFrame,
    row_position: int,
    question_col: str,
    answer_col: str | None,
) -> tuple[str, str]:
    row = df.iloc[row_position]
    current_question = to_text(row.get(question_col, ""))

    if HISTORY_COL in df.columns:
        dialogue_history = to_text(row.get(HISTORY_COL, ""))
        ensure_current_answer_not_in_history(
            df=df,
            row_position=row_position,
            answer_col=answer_col,
            dialogue_history=dialogue_history,
        )
        return dialogue_history, current_question

    if not answer_col:
        raise ValueError(f"Cannot build dialogue history without an answer column.")

    dialogue_history = build_history_from_previous_rows(
        df=df,
        row_position=row_position,
        question_col=question_col,
        answer_col=answer_col,
    )
    return dialogue_history, current_question


def ensure_current_answer_not_in_history(
    df: pd.DataFrame,
    row_position: int,
    answer_col: str | None,
    dialogue_history: str,
) -> None:
    if not answer_col:
        return

    current_answer = to_text(df.iloc[row_position].get(answer_col, ""))
    if current_answer and current_answer in dialogue_history:
        row_label = df.index[row_position]
        raise ValueError(
            f"Data leakage detected at row {row_label}: current respondent answer appears "
            f"in '{HISTORY_COL}'. Remove the current answer from history before prediction."
        )


def should_predict_row(df: pd.DataFrame, row_position: int, answer_col: str | None) -> bool:
    # If an answer column exists, blank respondent-answer rows are usually interviewer-only
    # fragments and should not be predicted as respondent emotions.
    if not answer_col:
        return True
    return not is_blank(df.iloc[row_position].get(answer_col, ""))


def predict_row(
    client,
    model: str,
    prompt_parts: dict[str, str],
    dialogue_history: str,
    current_question: str,
    temperature: float,
    max_tokens: int,
) -> tuple[str, str, str, str]:
    prompt = build_prompt(
        prompt_template=prompt_parts["prompt_template"],
        person_profile=prompt_parts["person_profile"],
        case_background=prompt_parts["case_background"],
        emotion_examples=prompt_parts["emotion_examples"],
        dialogue_history=dialogue_history,
        current_question=current_question,
    )

    raw_output = call_model_with_retry(
        client=client,
        model=model,
        prompt=prompt,
        max_retries=MAX_RETRIES,
        retry_wait_seconds=RETRY_WAIT_SECONDS,
        temperature=temperature,
        max_tokens=max_tokens,
    )
    raw_emotion, reason = parse_emotion_response(raw_output, ALLOWED_EMOTIONS)
    final_emotion = (
        "PARSE_ERROR"
        if raw_emotion == "PARSE_ERROR"
        else postprocess_emotion(raw_emotion, reason)
    )
    return final_emotion, raw_emotion, reason, raw_output


def save_excel(df: pd.DataFrame, output_path: Path) -> None:
    ensure_parent_dir(output_path)
    df.to_excel(output_path, index=False)


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(f"Input Excel file not found: {input_path}")

    prompt_parts = load_prompt_parts(args.prompts_dir)
    client = create_openai_client(args.env)
    model = get_model_name(args.model)

    df = pd.read_excel(input_path, sheet_name=parse_sheet_arg(args.sheet))
    work_df = df.head(args.limit).copy() if args.limit else df.copy()
    question_col, answer_col = validate_input_columns(work_df)

    for output_col in [
        OUTPUT_EMOTION_COL,
        OUTPUT_RAW_EMOTION_COL,
        OUTPUT_REASON_COL,
        OUTPUT_RAW_MODEL_COL,
    ]:
        if output_col not in work_df.columns:
            work_df[output_col] = ""

    row_positions = [
        row_position
        for row_position in range(len(work_df))
        if should_predict_row(work_df, row_position, answer_col)
    ]

    processed_count = 0
    for row_position in tqdm(row_positions, total=len(row_positions), desc="Predicting emotions"):
        try:
            dialogue_history, current_question = get_prompt_inputs(
                df=work_df,
                row_position=row_position,
                question_col=question_col,
                answer_col=answer_col,
            )
            final_emotion, raw_emotion, reason, raw_output = predict_row(
                client=client,
                model=model,
                prompt_parts=prompt_parts,
                dialogue_history=dialogue_history,
                current_question=current_question,
                temperature=args.temperature,
                max_tokens=args.max_tokens,
            )
        except Exception as exc:
            final_emotion = "API_ERROR"
            raw_emotion = "API_ERROR"
            reason = str(exc)
            raw_output = str(exc)

        row_index = work_df.index[row_position]
        work_df.at[row_index, OUTPUT_EMOTION_COL] = final_emotion
        work_df.at[row_index, OUTPUT_RAW_EMOTION_COL] = raw_emotion
        work_df.at[row_index, OUTPUT_REASON_COL] = reason
        work_df.at[row_index, OUTPUT_RAW_MODEL_COL] = raw_output

        processed_count += 1
        if processed_count % AUTOSAVE_EVERY == 0:
            save_excel(work_df, TEMP_OUTPUT_PATH)

    save_excel(work_df, output_path)
    print(f"Saved result to: {output_path}")


if __name__ == "__main__":
    main()
