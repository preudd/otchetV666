from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List


@dataclass(frozen=True)
class Rule:
    match_mode: str  # "any" | "all"
    keywords: List[str]


DEFAULT_RULES: Dict[str, Rule] = {
    # qty+sum categories (show "шт ... руб" in template)
    "ticket_unlimited": Rule(match_mode="any", keywords=["билет(безлимит)", "вход безлимит"]),
    "ticket_1hour": Rule(match_mode="any", keywords=["билет(1час)", "билет 1час"]),
    # Важно: не добавляем слишком общие слова типа "час" — иначе будет матчиться "3 часа" и т.п.
    "action_happy_hours": Rule(match_mode="any", keywords=["счастливые часы"]),
    "action_last_hour": Rule(match_mode="any", keywords=["последний час"]),
    # sum-only categories
    "aquagrim": Rule(match_mode="any", keywords=["аквагрим"]),
    "viar": Rule(match_mode="any", keywords=["viar"]),
    "balls": Rule(match_mode="any", keywords=["шар", "щар"]),
    "accompany": Rule(match_mode="any", keywords=["сопровождающ"]),
    # advance/rent fields in template
    "advance_dr": Rule(match_mode="any", keywords=["аванс др", "аванс день рождения"]),
    "advance_graduation": Rule(match_mode="any", keywords=["аванс выпускной"]),
    "rent_room": Rule(match_mode="any", keywords=["аренда комнаты"]),
    "advance_animator": Rule(match_mode="all", keywords=["аниматор", "аванс"]),
}


UI_CATEGORIES: List[Dict[str, str]] = [
    {"rule_key": "ticket_unlimited", "label": "Вход безлимит"},
    {"rule_key": "ticket_1hour", "label": "Билет 1час"},
    {"rule_key": "action_happy_hours", "label": "Акция счастливые часы"},
    {"rule_key": "action_last_hour", "label": "Акция последний час"},
    {"rule_key": "aquagrim", "label": "Аквагрим"},
    {"rule_key": "viar", "label": "Виар"},
    {"rule_key": "balls", "label": "Шары"},
    {"rule_key": "accompany", "label": "Сопровождающий"},
    {"rule_key": "advance_dr", "label": "Аванс ДР"},
    {"rule_key": "advance_animator", "label": "Аниматор выезд (аванс)"},
    {"rule_key": "rent_room", "label": "Аренда комнаты"},
    {"rule_key": "advance_graduation", "label": "Аванс выпускной"},
]


def _rules_path() -> Path:
    data_dir = os.getenv("DATA_DIR")
    base = Path(data_dir) if data_dir else Path(__file__).resolve().parent
    if not base.exists():
        base.mkdir(parents=True, exist_ok=True)
    return base / "category_rules.json"


def _normalize_keyword(s: str) -> str:
    s = (s or "").strip()
    s = " ".join(s.split())
    return s.lower()


def load_rules() -> Dict[str, Rule]:
    path = _rules_path()
    if not path.exists():
        # On first run: create file with defaults for easier troubleshooting.
        save_rules(DEFAULT_RULES)
        return dict(DEFAULT_RULES)

    raw: Dict[str, Any] = json.loads(path.read_text(encoding="utf-8"))
    out: Dict[str, Rule] = {}
    changed = False
    for rule_key, default_rule in DEFAULT_RULES.items():
        node = raw.get(rule_key) or {}
        match_mode = str(node.get("match_mode", default_rule.match_mode))
        keywords = node.get("keywords") or default_rule.keywords
        kw_list = [str(x).strip().lower() for x in keywords if str(x).strip()]
        # Миграция от ранних "слишком широких" дефолтов
        if rule_key == "action_happy_hours":
            before = list(kw_list)
            # убираем слишком общие ключевые слова прошлых версий
            kw_list = [k for k in kw_list if k not in ("час", "счастливые")]
            if not kw_list:
                kw_list = list(DEFAULT_RULES[rule_key].keywords)
            if kw_list != before:
                changed = True

        if rule_key == "action_last_hour":
            before = list(kw_list)
            kw_list = [k for k in kw_list if k != "последний"]
            if not kw_list:
                kw_list = list(DEFAULT_RULES[rule_key].keywords)
            if kw_list != before:
                changed = True

        out[rule_key] = Rule(match_mode=match_mode, keywords=kw_list)

    # Если применили миграцию — сохраняем обратно, чтобы эффект был сразу на хостинге.
    if changed:
        save_rules(out)
    return out


def save_rules(rules: Dict[str, Rule]) -> None:
    path = _rules_path()
    data = {k: {"match_mode": v.match_mode, "keywords": v.keywords} for k, v in rules.items()}
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def match_rule(nl: str, rule: Rule) -> bool:
    keywords = rule.keywords
    if not keywords:
        return False
    if rule.match_mode == "all":
        return all(kw in nl for kw in keywords)
    # default: "any"
    return any(kw in nl for kw in keywords)


def add_keyword(rule_key: str, keyword: str) -> Rule:
    if rule_key not in DEFAULT_RULES:
        raise ValueError(f"Unknown rule_key: {rule_key}")
    kw = _normalize_keyword(keyword)
    if not kw:
        raise ValueError("Keyword is empty")

    rules = load_rules()
    rule = rules[rule_key]
    if kw not in rule.keywords:
        rule = Rule(match_mode=rule.match_mode, keywords=[*rule.keywords, kw])
    rules[rule_key] = rule
    save_rules(rules)
    return rule


def get_rule(rule_key: str) -> Rule:
    if rule_key not in DEFAULT_RULES:
        raise ValueError(f"Unknown rule_key: {rule_key}")
    return load_rules()[rule_key]


def remove_keyword(rule_key: str, keyword: str) -> Rule:
    if rule_key not in DEFAULT_RULES:
        raise ValueError(f"Unknown rule_key: {rule_key}")
    kw = _normalize_keyword(keyword)
    if not kw:
        raise ValueError("Keyword is empty")

    rules = load_rules()
    rule = rules[rule_key]
    new_keywords = [x for x in rule.keywords if x != kw]
    if not new_keywords:
        raise ValueError("Нельзя удалить последнее ключевое слово категории.")
    rule = Rule(match_mode=rule.match_mode, keywords=new_keywords)
    rules[rule_key] = rule
    save_rules(rules)
    return rule

