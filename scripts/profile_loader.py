import hashlib
import json
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Profile:
    profile_id: str
    respondent_name: str
    story_background: str
    background_path: Path
    background_sha256: str


class ProfileRegistry:
    def __init__(
        self,
        profiles: dict[str, Profile],
        bindings: dict[str, str],
        config_path: Path,
    ) -> None:
        self.profiles = profiles
        self.bindings = bindings
        self.config_path = config_path

    def resolve(
        self,
        input_path: str | Path,
        sheet_name: str,
        profile_id: str | None = None,
    ) -> Profile:
        selected_id = profile_id or self.bindings.get(
            make_binding_key(input_path, sheet_name)
        )
        if not selected_id:
            raise KeyError(
                "No profile binding for "
                f"'{make_binding_key(input_path, sheet_name)}'. "
                "Add it to the profile registry or pass --profile-id."
            )
        try:
            return self.profiles[selected_id]
        except KeyError as exc:
            raise KeyError(f"Unknown profile_id in binding: {selected_id}") from exc


def make_binding_key(input_path: str | Path, sheet_name: str) -> str:
    return f"{Path(input_path).name}::{sheet_name}"


def read_background(path: str | Path) -> str:
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(f"Profile background not found: {file_path}")
    if file_path.suffix.lower() == ".docx":
        from docx import Document

        document = Document(file_path)
        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs]
        return "\n".join(text for text in paragraphs if text)
    if file_path.suffix.lower() != ".txt":
        raise ValueError(
            f"Unsupported profile background format: {file_path.suffix}. "
            "Use .docx or UTF-8 .txt."
        )
    return file_path.read_text(encoding="utf-8").strip()


def file_sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_profile_registry(path: str | Path) -> ProfileRegistry:
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Profile registry not found: {config_path}")
    data = json.loads(config_path.read_text(encoding="utf-8"))
    raw_profiles = data.get("profiles")
    raw_bindings = data.get("bindings")
    if not isinstance(raw_profiles, dict) or not isinstance(raw_bindings, dict):
        raise ValueError("Profile registry requires object fields: profiles and bindings.")

    profiles: dict[str, Profile] = {}
    for profile_id, config in raw_profiles.items():
        if not isinstance(config, dict):
            raise ValueError(f"Profile '{profile_id}' must be a JSON object.")
        respondent_name = str(config.get("respondent_name", "")).strip()
        background_value = str(config.get("story_background", "")).strip()
        if not respondent_name or not background_value:
            raise ValueError(
                f"Profile '{profile_id}' requires respondent_name and story_background."
            )
        background_path = Path(background_value)
        if not background_path.is_absolute():
            background_path = config_path.parent / background_path
        profiles[profile_id] = Profile(
            profile_id=profile_id,
            respondent_name=respondent_name,
            story_background=read_background(background_path),
            background_path=background_path.resolve(),
            background_sha256=file_sha256(background_path),
        )

    bindings = {str(key): str(value) for key, value in raw_bindings.items()}
    unknown = sorted(set(bindings.values()) - set(profiles))
    if unknown:
        raise ValueError("Bindings reference unknown profiles: " + ", ".join(unknown))
    return ProfileRegistry(profiles, bindings, config_path.resolve())


def apply_profile_to_emotion_prompt_parts(
    prompt_parts: dict[str, str], profile: Profile
) -> dict[str, str]:
    return {
        **prompt_parts,
        "person_profile": f"被谈话人姓名：{profile.respondent_name}",
        "case_background": profile.story_background,
    }

