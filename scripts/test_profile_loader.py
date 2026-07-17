import json
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from profile_loader import (
    apply_profile_to_emotion_prompt_parts,
    load_profile_registry,
    make_binding_key,
)


class ProfileLoaderTests(unittest.TestCase):
    def test_resolve_profile_by_workbook_and_sheet(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "story.txt").write_text("完整人物背景", encoding="utf-8")
            config = {
                "profiles": {
                    "person_a": {
                        "respondent_name": "人物甲",
                        "story_background": "story.txt",
                    }
                },
                "bindings": {"dialogue.xlsx::sheet1": "person_a"},
            }
            config_path = root / "profiles.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            registry = load_profile_registry(config_path)
            profile = registry.resolve("some/path/dialogue.xlsx", "sheet1")
            self.assertEqual(profile.respondent_name, "人物甲")
            self.assertEqual(profile.story_background, "完整人物背景")
            self.assertEqual(len(profile.background_sha256), 64)

    def test_binding_key_ignores_parent_directory(self) -> None:
        self.assertEqual(
            make_binding_key("/tmp/a/example_dialogue.xlsx", "对话整理"),
            "example_dialogue.xlsx::对话整理",
        )

    def test_profile_overrides_emotion_background(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "story.txt").write_text("新背景", encoding="utf-8")
            config_path = root / "profiles.json"
            config_path.write_text(
                json.dumps(
                    {
                        "profiles": {
                            "p": {
                                "respondent_name": "新人物",
                                "story_background": "story.txt",
                            }
                        },
                        "bindings": {},
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )
            profile = load_profile_registry(config_path).resolve(
                "anything.xlsx", "anything", profile_id="p"
            )
            result = apply_profile_to_emotion_prompt_parts(
                {"person_profile": "旧人物", "case_background": "旧背景"}, profile
            )
            self.assertIn("新人物", result["person_profile"])
            self.assertEqual(result["case_background"], "新背景")


if __name__ == "__main__":
    unittest.main()
