import os
import re
import time
from pathlib import Path
from typing import Iterable

from dotenv import load_dotenv
from openai import OpenAI


DEFAULT_BASE_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1"
DEFAULT_MODEL = "qwen-plus"

ALLOWED_EMOTIONS = {
    "neutral",
    "confusion",
    "disappointment",
    "nervousness",
    "anger",
    "annoyance",
    "disapproval",
    "disgust",
    "embarrassment",
    "fear",
    "sadness",
    "remorse",
}


def read_text_file(path: str | Path) -> str:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Required file not found: {file_path}")
    return file_path.read_text(encoding="utf-8").strip()


def load_prompt_parts(prompts_dir: str | Path = "prompts") -> dict[str, str]:
    base_dir = Path(prompts_dir)
    prompt_template = read_text_file(base_dir / "emotion_prompt.txt")
    required_placeholders = {
        "{person_profile}",
        "{case_background}",
        "{emotion_examples}",
        "{dialogue_history}",
        "{current_question}",
    }
    missing = [item for item in required_placeholders if item not in prompt_template]
    if missing:
        raise ValueError(
            "emotion_prompt.txt missing placeholders: " + ", ".join(sorted(missing))
        )

    return {
        "prompt_template": prompt_template,
        "person_profile": read_text_file(base_dir / "person_profile.txt"),
        "case_background": read_text_file(base_dir / "case_background.txt"),
        "emotion_examples": read_text_file(base_dir / "emotion_examples.txt"),
    }


def build_prompt(
    prompt_template: str,
    person_profile: str,
    case_background: str,
    emotion_examples: str,
    dialogue_history: str,
    current_question: str,
) -> str:
    return prompt_template.format(
        person_profile=person_profile,
        case_background=case_background,
        emotion_examples=emotion_examples,
        dialogue_history=dialogue_history,
        current_question=current_question,
    )


def create_openai_client(env_path: str | Path = ".env") -> OpenAI:
    load_dotenv(env_path)
    api_key = os.getenv("DASHSCOPE_API_KEY")
    if not api_key:
        raise RuntimeError("Missing DASHSCOPE_API_KEY. Please configure it in .env.")

    base_url = os.getenv("DASHSCOPE_BASE_URL", DEFAULT_BASE_URL)
    return OpenAI(api_key=api_key, base_url=base_url)


def get_model_name(model: str | None = None) -> str:
    return model or os.getenv("DASHSCOPE_MODEL") or DEFAULT_MODEL


def call_model_with_retry(
    client: OpenAI,
    model: str,
    prompt: str,
    max_retries: int = 3,
    retry_wait_seconds: float = 2.0,
    temperature: float = 0.0,
    max_tokens: int = 512,
) -> str:
    last_error: Exception | None = None

    for attempt in range(1, max_retries + 1):
        try:
            return call_model(
                client=client,
                model=model,
                prompt=prompt,
                temperature=temperature,
                max_tokens=max_tokens,
            )
        except Exception as exc:
            last_error = exc
            if attempt < max_retries:
                time.sleep(retry_wait_seconds)

    raise RuntimeError(f"API call failed after {max_retries} attempts: {last_error}")


def call_model(
    client: OpenAI,
    model: str,
    prompt: str,
    temperature: float = 0.0,
    max_tokens: int = 512,
) -> str:
    request_kwargs = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": "你是审讯对话情绪预测助手，必须严格按指定格式输出。",
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
    }

    # Qwen3 non-streaming calls on DashScope require disabling thinking explicitly.
    if model.lower().startswith("qwen3"):
        request_kwargs["extra_body"] = {"enable_thinking": False}

    response = client.chat.completions.create(**request_kwargs)
    return response.choices[0].message.content or ""


def parse_emotion_response(
    raw_output: str,
    allowed_emotions: Iterable[str] = ALLOWED_EMOTIONS,
) -> tuple[str, str]:
    text = raw_output.strip()
    allowed = set(allowed_emotions)

    emotion_match = re.search(
        r"Emotion\s*[:：]\s*(?P<emotion>[A-Za-z_ -]+)",
        text,
        flags=re.IGNORECASE,
    )
    reason_match = re.search(
        r"Reason\s*[:：]\s*(?P<reason>.*)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )

    emotion = ""
    if emotion_match:
        emotion = emotion_match.group("emotion").splitlines()[0].strip().lower()
        emotion = emotion.replace(" ", "_").replace("-", "_")

    reason = reason_match.group("reason").strip() if reason_match else ""

    if emotion not in allowed:
        return "PARSE_ERROR", reason

    return emotion, reason


def postprocess_emotion(emotion: str, reason: str) -> str:
    reason_text = reason or ""

    if emotion == "anger" and not contains_any(
        reason_text, ["激烈", "爆发", "强烈", "对抗", "攻击"]
    ):
        return "annoyance"

    if emotion == "remorse":
        factual_terms = ["承认事实", "金额", "借款", "时间", "见面", "经办过程"]
        remorse_terms = ["后悔", "愧疚", "自责", "认错", "反思责任"]
        if contains_any(reason_text, factual_terms) and not contains_any(reason_text, remorse_terms):
            return "nervousness"

    if emotion == "confusion" and contains_any(reason_text, ["回避", "含糊", "拖延", "解释"]):
        return "nervousness"

    if emotion == "fear":
        weak_pressure_terms = ["一般追问", "一般压力"]
        fear_terms = ["关键证据", "严重后果", "家庭牵连", "组织处理", "暴露"]
        if contains_any(reason_text, weak_pressure_terms) and not contains_any(reason_text, fear_terms):
            return "nervousness"

    return emotion


def contains_any(text: str, keywords: Iterable[str]) -> bool:
    return any(keyword in text for keyword in keywords)


def ensure_parent_dir(path: str | Path) -> None:
    Path(path).parent.mkdir(parents=True, exist_ok=True)


def is_blank(value: object) -> bool:
    if value is None:
        return True
    try:
        import pandas as pd

        if pd.isna(value):
            return True
    except (ImportError, TypeError, ValueError):
        pass
    return str(value).strip() == ""


def to_text(value: object) -> str:
    return "" if is_blank(value) else str(value).strip()
