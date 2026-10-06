from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, replace
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from .curriculum import chapter_names
from .parser import QuestionDraft
from .taxonomy import DIFFICULTIES, MAJOR_CATEGORIES, QUESTION_TYPES


@dataclass(frozen=True)
class AIConfig:
    enabled: bool = False
    base_url: str = ""
    api_key: str = ""
    model: str = ""
    timeout: float = 20.0


DEFAULT_AI_BASE_URL = "https://api.deepseek.com/v1"
DEFAULT_AI_MODEL = "deepseek-chat"


def load_ai_config(project_root=None, env=None) -> AIConfig:
    values = dict(os.environ if env is None else env)
    if project_root:
        path = Path(project_root) / ".env"
        if path.exists():
            for raw in path.read_text(encoding="utf-8-sig").splitlines():
                if raw.strip() and not raw.lstrip().startswith("#") and "=" in raw:
                    key, value = raw.split("=", 1)
                    values[key.strip()] = value.strip().strip('"').strip("'")
    api_key = values.get("AI_API_KEY", "")
    enabled_value = values.get("AI_ENABLED", "").lower()
    enabled = enabled_value in {"1", "true", "yes"} or (not enabled_value and bool(api_key))
    return AIConfig(
        enabled,
        (values.get("AI_BASE_URL", "").rstrip("/") or (DEFAULT_AI_BASE_URL if api_key else "")),
        api_key,
        (values.get("AI_MODEL", "") or (DEFAULT_AI_MODEL if api_key else "")),
        float(values.get("AI_TIMEOUT", "20")),
    )


def save_ai_config(project_root, config: dict) -> Path:
    path = Path(project_root) / ".env"
    api_key = str(config.get("api_key", "")).strip()
    enabled_value = config.get("enabled")
    enabled = bool(api_key) and (
        enabled_value is None or str(enabled_value).lower() in {"1", "true", "yes", "on"}
    )
    desired = {
        "AI_ENABLED": "true" if enabled else "false",
        "AI_BASE_URL": str(config.get("base_url", "")).strip().rstrip("/") or (DEFAULT_AI_BASE_URL if api_key else ""),
        "AI_API_KEY": api_key,
        "AI_MODEL": str(config.get("model", "")).strip() or (DEFAULT_AI_MODEL if api_key else ""),
        "AI_TIMEOUT": str(config.get("timeout", "20")).strip() or "20",
    }

    lines: list[str] = []
    seen: set[str] = set()
    if path.exists():
        for raw in path.read_text(encoding="utf-8-sig").splitlines():
            stripped = raw.strip()
            if stripped and not stripped.startswith("#") and "=" in raw:
                key, _ = raw.split("=", 1)
                key = key.strip()
                if key in desired:
                    lines.append(f"{key}={desired[key]}")
                    seen.add(key)
                    continue
            lines.append(raw)

    for key, value in desired.items():
        if key not in seen:
            lines.append(f"{key}={value}")

    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return path


def classify_low_confidence(questions: list[QuestionDraft], config: AIConfig) -> tuple[list[QuestionDraft], list[str]]:
    targets = [q for q in questions if (q.classification_confidence or 0) < 0.6]
    return _classify_with_ai(questions, targets, config, apply_all=False)


def classify_import_questions(questions: list[QuestionDraft], config: AIConfig) -> tuple[list[QuestionDraft], list[str]]:
    return _classify_with_ai(questions, questions, config, apply_all=True)


def _classify_with_ai(
    questions: list[QuestionDraft],
    targets: list[QuestionDraft],
    config: AIConfig,
    *,
    apply_all: bool,
) -> tuple[list[QuestionDraft], list[str]]:
    if not config.enabled:
        return questions, []
    if not all((config.base_url, config.api_key, config.model)):
        return questions, ["AI 已启用但配置不完整，已使用本地分类。"]
    if not targets:
        return questions, []

    allowed_categories = "、".join(MAJOR_CATEGORIES)
    payload = {
        "model": config.model,
        "temperature": 0,
        "messages": [
            {
                "role": "system",
                "content": (
                    "你是高中数学题目分类器。请根据题干、选项、答案和解析重新识别题目。"
                    "重点输出大类 category_major。"
                    "优先级：导数 > 解析几何 > 概率统计 > 空间几何 > 解三角形 > 三角函数 > 集合 > 不等式 > 数列 > 向量 > 函数。"
                    "当题目同时出现多个大类线索时，按上述优先级判断。"
                    f"category_major 只能从这些大类中选择：{allowed_categories}。"
                    "每题必须返回 number、category_major、category_minor、question_type、difficulty、confidence。"
                    "confidence 用 0 到 1 的小数，例如 0.92。"
                    "只返回 JSON 对象，格式为 {\"results\": [...]}。"
                ),
            },
            {
                "role": "user",
                "content": json.dumps([_serialize_question(q) for q in targets], ensure_ascii=False),
            },
        ],
    }

    try:
        req = Request(
            config.base_url + "/chat/completions",
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": "Bearer " + config.api_key,
                "Content-Type": "application/json",
            },
        )
        with urlopen(req, timeout=config.timeout) as response:
            raw = json.loads(response.read().decode("utf-8", errors="ignore"))
        content = raw["choices"][0]["message"]["content"]
        parsed = _parse_json_object(content)
        results = parsed.get("results", [])
    except (HTTPError, URLError, TimeoutError, ValueError, KeyError, json.JSONDecodeError) as exc:
        return questions, [f"AI 分类不可用，已使用本地分类：{type(exc).__name__}"]

    by_number = {str(item.get("number")): item for item in results if isinstance(item, dict)}
    output: list[QuestionDraft] = []
    valid_chapters = set(chapter_names())
    missing_count = 0
    invalid_count = 0

    for question in questions:
        item = by_number.get(question.number)
        if not item:
            if question in targets:
                missing_count += 1
            output.append(question)
            continue

        updated = _apply_ai_result(question, item, valid_chapters)
        if updated is None:
            invalid_count += 1
            output.append(question)
            continue
        if not apply_all and (question.classification_confidence or 0) >= 0.6:
            output.append(question)
            continue
        output.append(updated)

    warnings = []
    if missing_count or invalid_count:
        warnings.append(
            f"AI 已调用，但有 {missing_count + invalid_count} 题没有返回合法大分类，已保留本地规则。"
        )
    return output, warnings


def _apply_ai_result(
    question: QuestionDraft,
    item: dict,
    valid_chapters: set[str],
) -> QuestionDraft | None:
    category_major = str(item.get("category_major", "")).strip()
    category_minor = str(item.get("category_minor", "")).strip()
    qtype = str(item.get("question_type", question.question_type)).strip()
    difficulty = str(item.get("difficulty", question.difficulty)).strip()
    chapter = str(item.get("chapter", "")).strip()
    section = str(item.get("section", "")).strip()
    review_reason = str(item.get("review_reason", "")).strip()

    confidence_value = _normalize_confidence(item.get("confidence"))

    if not category_major or category_major not in MAJOR_CATEGORIES:
        return None
    if qtype and qtype not in QUESTION_TYPES:
        return None
    if difficulty and difficulty not in DIFFICULTIES:
        return None
    if chapter and chapter not in valid_chapters:
        return None

    if not category_minor:
        category_minor = question.category_minor

    return replace(
        question,
        question_type=qtype or question.question_type,
        category_major=category_major or question.category_major,
        category_minor=category_minor or question.category_minor,
        difficulty=difficulty or question.difficulty,
        curriculum_chapter=chapter or question.curriculum_chapter,
        curriculum_section=section or question.curriculum_section,
        classification_source="ai",
        classification_confidence=confidence_value,
        review_reason=review_reason if review_reason else ("" if confidence_value is None or confidence_value >= 0.75 else "AI 分类待人工确认"),
        rule_category_major=question.category_major,
        rule_category_minor=question.category_minor,
        ai_category_major=category_major or question.category_major,
        ai_category_minor=category_minor or question.category_minor,
        ai_review_note=_ai_review_note(question.category_major, category_major or question.category_major),
    )


def _ai_review_note(rule_major: str, ai_major: str) -> str:
    rule_major = (rule_major or "未分类").strip() or "未分类"
    ai_major = (ai_major or "未分类").strip() or "未分类"
    if rule_major == ai_major:
        return f"AI确认本地分类：{ai_major}"
    return f"AI改分类：规则{rule_major} -> AI{ai_major}"


def _normalize_confidence(value) -> float | None:
    if value in (None, ""):
        return None
    try:
        confidence = float(value)
    except (TypeError, ValueError):
        return None
    if 0 <= confidence <= 1:
        return confidence
    if 1 < confidence <= 100:
        return confidence / 100
    return None


def _serialize_question(question: QuestionDraft) -> dict:
    return {
        "number": question.number,
        "stem": question.stem[:1500],
        "options": "\n".join(question.options)[:1000],
        "answer": question.answer[:500],
        "analysis": question.analysis[:1500],
        "question_type": question.question_type,
        "category_major": question.category_major,
        "category_minor": question.category_minor,
        "difficulty": question.difficulty,
        "classification_source": question.classification_source,
        "classification_confidence": question.classification_confidence,
    }


def _parse_json_object(content: str) -> dict:
    text = content.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z0-9]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        text = text[start : end + 1]
    return json.loads(text)


def masked_ai_status(config: AIConfig) -> dict:
    return {
        "enabled": config.enabled,
        "configured": bool(config.base_url and config.api_key and config.model),
        "base_url": config.base_url,
        "model": config.model,
        "api_key": "***" + config.api_key[-4:] if config.api_key else "",
    }
