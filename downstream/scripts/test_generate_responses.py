import importlib.util
import unittest
from unittest.mock import patch
from pathlib import Path

import pandas as pd


SCRIPT_PATH = Path(__file__).with_name("generate_responses.py")
SPEC = importlib.util.spec_from_file_location("generate_responses", SCRIPT_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class GenerateResponsesTests(unittest.TestCase):
    def test_history_excludes_current_answer(self) -> None:
        df = pd.DataFrame(
            {
                "提问人员": ["问题一", "问题二", "问题三"],
                "被谈话人": ["回答一", "回答二", "当前真实回答"],
            }
        )
        history = MODULE.build_real_history(df, 2, "提问人员", "被谈话人", None)
        self.assertIn("回答一", history)
        self.assertIn("回答二", history)
        self.assertNotIn("当前真实回答", history)
        self.assertNotIn("问题三", history)

    def test_history_limit_uses_latest_real_turns(self) -> None:
        df = pd.DataFrame(
            {
                "提问人员": ["问题一", "问题二", "问题三"],
                "被谈话人": ["回答一", "回答二", "回答三"],
            }
        )
        history = MODULE.build_real_history(df, 2, "提问人员", "被谈话人", 1)
        self.assertNotIn("问题一", history)
        self.assertNotIn("回答一", history)
        self.assertIn("问题二", history)
        self.assertIn("回答二", history)

    def test_three_answers_are_adjacent(self) -> None:
        df = pd.DataFrame(
            {
                "轮次": [1],
                "提问人员": ["问题"],
                "被谈话人": ["真实回答"],
                MODULE.EMOTION_COL: ["neutral"],
                MODULE.RAW_EMOTION_COL: ["neutral"],
                MODULE.EMOTION_REASON_COL: ["原因"],
                MODULE.EMOTION_RAW_OUTPUT_COL: ["原始输出"],
                MODULE.BASELINE_OUTPUT_COL: ["基线回答"],
                MODULE.EMOTION_OUTPUT_COL: ["情绪回答"],
                MODULE.BASELINE_STATUS_COL: ["OK"],
                MODULE.EMOTION_PREDICTION_STATUS_COL: ["OK"],
                MODULE.EMOTION_GENERATION_STATUS_COL: ["OK"],
            }
        )
        result = MODULE.reorder_comparison_columns(df, "提问人员", "被谈话人")
        self.assertEqual(
            result.columns[1:5].tolist(),
            [
                "提问人员",
                "被谈话人",
                MODULE.BASELINE_OUTPUT_COL,
                MODULE.EMOTION_OUTPUT_COL,
            ],
        )

    def test_only_emotion_prompt_receives_label(self) -> None:
        materials = {
            "respondent_name": "测试人物",
            "story_background": "背景事实",
            "baseline_template": "{story_background}|{dialogue_history}|{current_question}",
            "emotion_template": "{story_background}|{dialogue_history}|{current_question}|{emotion}",
        }
        baseline, conditioned = MODULE.build_prompts(
            materials, "此前历史", "当前问题", "nervousness"
        )
        self.assertNotIn("nervousness", baseline)
        self.assertIn("nervousness", conditioned)
        self.assertNotIn("预测原因文本", baseline)
        self.assertNotIn("预测原因文本", conditioned)

    def test_fresh_emotion_is_parsed_and_postprocessed(self) -> None:
        prompt_parts = {
            "prompt_template": (
                "{person_profile}|{case_background}|{emotion_examples}|"
                "{dialogue_history}|{current_question}"
            ),
            "person_profile": "人物",
            "case_background": "案件",
            "emotion_examples": "示例",
        }
        raw_output = "Emotion：\nanger\n\nReason：\n只是一般不耐烦。"
        with patch.object(
            MODULE, "call_emotion_model_with_retry", return_value=raw_output
        ):
            result = MODULE.predict_fresh_emotion(
                client=object(),
                model="qwen3-8b",
                prompt_parts=prompt_parts,
                dialogue_history="此前历史",
                current_question="当前问题",
                temperature=0.0,
                max_tokens=512,
            )
        self.assertEqual(result[0], "annoyance")
        self.assertEqual(result[1], "anger")
        self.assertEqual(result[4], "OK")

    def test_new_sensitive_topic_starts_guarded(self) -> None:
        state = MODULE.determine_dialogue_state("", "你是否收过某企业负责人送的礼品？")
        self.assertIn("尚未出现明确承认", state)
        self.assertIn("新的敏感事项", state)
        self.assertIn("试探防御", state)

    def test_direct_evidence_increases_disclosure_pressure(self) -> None:
        history = "审调人员：你收过礼品吗？\n被谈话人：确实收过一些普通礼品。"
        question = "银行流水和转账记录显示有10万元，这怎么解释？"
        state = MODULE.determine_dialogue_state(history, question)
        self.assertIn("明确证据", state)
        self.assertIn("具体金额", state)
        self.assertIn("逐步松动", state)

    def test_interviewer_claim_is_not_counted_as_admission(self) -> None:
        history = "审调人员：你已经承认收了礼品。\n被谈话人：我没有这样说。"
        state = MODULE.determine_dialogue_state(history, "你再考虑一下。")
        self.assertIn("尚未出现明确承认", state)

    def test_length_guidance_uses_previous_real_answers_only(self) -> None:
        df = pd.DataFrame(
            {
                "被谈话人": ["简短回答", "这也是一条比较简短的回答", "当前真实回答非常非常长但不能使用"],
            }
        )
        guidance = MODULE.build_response_length_guidance(
            df, 2, "被谈话人", "你和他是什么关系？"
        )
        self.assertIn("中位长度", guidance)
        self.assertNotIn("当前真实回答", guidance)

    def test_procedural_question_gets_short_guidance(self) -> None:
        df = pd.DataFrame({"被谈话人": ["之前回答"]})
        guidance = MODULE.build_response_length_guidance(
            df, 0, "被谈话人", "是否申请回避？"
        )
        self.assertIn("5至30字", guidance)

    def test_stage_directions_are_removed(self) -> None:
        raw = "（叹气，低头沉默片刻）这件事我记不太清了。"
        self.assertEqual(MODULE.clean_generation_output(raw), "这件事我记不太清了。")


if __name__ == "__main__":
    unittest.main()
