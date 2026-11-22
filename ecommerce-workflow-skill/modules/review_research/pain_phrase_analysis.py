from __future__ import annotations

import math
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Any, Iterable


WORD_PATTERN = re.compile(r"[A-Za-zÀ-ÖØ-öø-ÿ]+(?:['’][A-Za-zÀ-ÖØ-öø-ÿ]+)?|\d+")
NEGATIVE_CUES = {
    "bad", "break", "broken", "cheap", "crack", "damage", "defect", "difficult", "disappoint",
    "expensive", "fail", "flimsy", "hard", "issue", "leak", "loose", "missing", "never", "no",
    "noise", "noisy", "not", "odor", "poor", "problem", "refund", "return", "rust", "scam",
    "small", "smell", "sticky", "stop", "tight", "useless", "waste", "weak", "wrong",
    "kaputt", "mangelhaft", "nicht", "rost", "schlecht", "schwach", "schwierig",
    "difícil", "malo", "olor", "roto",
}
BOUNDARY_STOPWORDS = {
    "a", "an", "and", "are", "as", "at", "be", "been", "but", "by", "for", "from", "had", "has",
    "have", "he", "her", "hers", "him", "his", "i", "if", "in", "is", "it", "its", "me", "my",
    "of", "on", "or", "our", "ours", "she", "so", "that", "the", "their", "theirs", "them", "they",
    "this", "those", "to", "too", "us", "was", "we", "were", "with", "you", "your", "yours",
    "actually", "after", "already", "also", "anything", "before", "course", "day", "days", "did", "do", "does",
    "doing", "done", "during", "even", "however", "honestly", "item", "just", "know", "machine", "make",
    "made", "merely", "month", "months", "only", "product", "really", "second", "seconds", "sometime",
    "still", "tell", "than", "then", "there", "time", "times", "unfortunately", "use", "used", "using",
    "vegetable", "vegetables", "very", "what", "when", "where", "which",
}
NEGATION_WORDS = {"never", "no", "not", "nicht"}
TOKEN_NORMALIZATION = {
    "broke": "break", "broken": "break", "breaking": "break", "breaks": "break",
    "cracked": "crack", "cracking": "crack", "cracks": "crack",
    "damaged": "damage", "damaging": "damage",
    "defective": "defect", "defects": "defect",
    "disappointed": "disappoint", "disappointing": "disappoint",
    "failed": "fail", "failing": "fail", "fails": "fail",
    "leaked": "leak", "leaking": "leak", "leaks": "leak",
    "noises": "noise", "problems": "problem", "issues": "issue",
    "returned": "return", "returning": "return", "returns": "return",
    "rusted": "rust", "rusting": "rust", "rusts": "rust",
    "smelled": "smell", "smelling": "smell", "smells": "smell",
    "stopped": "stop", "stopping": "stop", "stops": "stop",
    "worked": "work", "working": "work", "works": "work",
}


@dataclass(frozen=True)
class Candidate:
    key: str
    surface: str
    score: float


@dataclass(frozen=True)
class FrequencyRow:
    phrase: str
    review_count: int
    review_ratio: float
    occurrence_count: int


@dataclass(frozen=True)
class SemanticGroup:
    english: str
    chinese: str


@dataclass(frozen=True)
class GroupedFrequencyRow:
    phrase: str
    translation: str
    review_count: int
    review_ratio: float
    variants: tuple[str, ...]


def tokenize(value: Any) -> list[str]:
    text = str(value or "").replace("’", "'").lower()
    return WORD_PATTERN.findall(text)


def canonical_token(token: str) -> str:
    normalized = TOKEN_NORMALIZATION.get(token, token)
    if normalized.endswith("'t"):
        return "not"
    return normalized


def canonical_key(tokens: Iterable[str]) -> str:
    return " ".join(canonical_token(token) for token in tokens)


def normalize_semantic_groups(
    groups: dict[str, SemanticGroup] | None,
) -> dict[str, SemanticGroup]:
    return {
        canonical_key(tokenize(key)): value
        for key, value in (groups or {}).items()
        if canonical_key(tokenize(key))
    }


def custom_group_matches_context(group: SemanticGroup, review_text: Any = "") -> bool:
    """防止跨品类缓存把同一个通用短语错误归到另一种产品。"""
    context_tokens = set(canonical_key(tokenize(review_text)).split())
    english = group.english.lower()
    english_tokens = set(canonical_key(tokenize(english)).split())
    domain_rules = (
        (
            {"flight", "propeller", "parachute", "launch"},
            {"flight", "fly", "flies", "flying", "propeller", "rocket", "parachute", "launch", "drift"},
        ),
        (
            {"steam", "water leakage", "floor glide", "cleaning"},
            {"steam", "water", "mop", "floor", "clean", "vacuum", "reservoir", "tank"},
        ),
        (
            {"ear", "camera", "lens", "phone connection"},
            {"ear", "canal", "otoscope", "camera", "lens", "iphone", "android", "phone", "app"},
        ),
    )
    for group_markers, context_markers in domain_rules:
        has_domain_marker = any(
            marker in english if " " in marker else marker in english_tokens
            for marker in group_markers
        )
        if has_domain_marker and not context_tokens.intersection(context_markers):
            return False
    return True


def semantic_group_for_phrase(
    phrase: str,
    custom_groups: dict[str, SemanticGroup] | None = None,
    *,
    review_text: Any = "",
) -> SemanticGroup | None:
    key = canonical_key(tokenize(phrase))
    token_set = set(key.split())
    # Agent 已审核的精确缓存应覆盖通用启发式，避免词义被旧规则误判。
    if custom_groups and key in custom_groups:
        group = custom_groups[key]
        if custom_group_matches_context(group, review_text):
            return group
    if "sticky" in token_set:
        return SemanticGroup("surface stickiness", "表面黏腻")
    if "rust" in token_set or "corrosion" in token_set:
        return SemanticGroup("rust and corrosion", "生锈与腐蚀")
    if key in {"expensive return", "return shipping"}:
        return SemanticGroup("high return cost", "退货成本高")
    if key == "return window" or key == "return":
        return SemanticGroup("returns and refunds", "退货与退款")
    if (
        "break" in token_set
        or "fail" in token_set
        or "not stay" in key
        or "never work" in key
        or "stop work" in key
    ):
        return SemanticGroup("product failure", "产品失效")
    if "durable" in token_set or token_set & {"flimsy", "weak"}:
        return SemanticGroup("quality and durability", "质量与耐用性差")
    if "material" in token_set or "cheap" in token_set or "poor quality" in key:
        return SemanticGroup("quality and durability", "质量与耐用性差")
    if key in {"hand wash", "not taste", "does it work"}:
        return SemanticGroup("unclear effectiveness", "效果不明显")
    if (
        "waste money" in key
        or "scam" in token_set
        or {"not", "sure", "worth"} <= token_set
    ):
        return SemanticGroup("poor value for money", "性价比低")
    if "leak" in token_set:
        return SemanticGroup("water leakage", "漏水")
    if token_set & {"small", "large", "tight", "loose"}:
        return SemanticGroup("size and fit", "尺寸或适配问题")
    if token_set & {"odor", "smell"}:
        return SemanticGroup("odor issue", "异味问题")
    if token_set & {"noise", "noisy", "loud"}:
        return SemanticGroup("noise issue", "噪音过大")
    if token_set & {"missing", "damage", "crack"}:
        return SemanticGroup("damaged or missing parts", "破损或缺件")
    return None


def aggregate_semantic_groups(
    reviews: list[dict[str, Any]],
    terms_by_review: list[list[str]],
    *,
    custom_groups: dict[str, SemanticGroup] | None = None,
    phrase_translations: dict[str, str] | None = None,
) -> tuple[list[GroupedFrequencyRow], list[str]]:
    custom_groups = normalize_semantic_groups(custom_groups)
    phrase_translations = phrase_translations or {}
    coverage: Counter[str] = Counter()
    labels: dict[str, SemanticGroup] = {}
    variants: dict[str, set[str]] = defaultdict(set)
    missing_groups: list[str] = []

    for review, phrases in zip(reviews, terms_by_review):
        groups_in_review: dict[str, SemanticGroup] = {}
        for phrase in phrases:
            review_text = " ".join(
                str(review.get(field) or "")
                for field in ("title", "body", "review_text")
            )
            group = semantic_group_for_phrase(
                phrase,
                custom_groups,
                review_text=review_text,
            )
            if group is None:
                group = SemanticGroup(phrase, phrase_translations.get(phrase, ""))
                missing_groups.append(phrase)
            group_key = group.english.lower()
            groups_in_review[group_key] = group
            labels[group_key] = group
        variant = str(review.get("variant") or "").strip()
        for group_key, group in groups_in_review.items():
            coverage[group_key] += 1
            if variant:
                variants[group_key].add(variant)
            labels[group_key] = group

    total_reviews = len(reviews)
    rows = [
        GroupedFrequencyRow(
            phrase=labels[key].english,
            translation=labels[key].chinese,
            review_count=count,
            review_ratio=count / total_reviews if total_reviews else 0,
            variants=tuple(sorted(variants[key])),
        )
        for key, count in coverage.items()
    ]
    rows.sort(key=lambda item: (-item.review_count, item.phrase))
    return rows, list(dict.fromkeys(missing_groups))


def add_candidate(store: dict[str, Candidate], tokens: list[str], score: float) -> None:
    if not tokens:
        return
    key = canonical_key(tokens)
    surface = " ".join(tokens)
    if not key or all(token in BOUNDARY_STOPWORDS for token in tokens):
        return
    existing = store.get(key)
    candidate = Candidate(key=key, surface=surface, score=score)
    if existing is None or candidate.score > existing.score:
        store[key] = candidate


def review_candidates(review: dict[str, Any]) -> dict[str, Candidate]:
    candidates: dict[str, Candidate] = {}
    title = str(review.get("title") or "")
    body_sections = re.split(r"[.!?;\n]+", str(review.get("body") or ""))
    sections = [(title, 5.0), *((section, 0.0) for section in body_sections)]
    for text, section_bonus in sections:
        tokens = tokenize(text)
        if not tokens:
            continue
        canonical = [canonical_token(token) for token in tokens]
        if section_bonus and 1 <= len(tokens) <= 5:
            title_has_cue = any(token in NEGATIVE_CUES for token in canonical)
            if title_has_cue or len(tokens) <= 4:
                add_candidate(candidates, tokens, 8.0 + section_bonus)

        for size in (1, 2, 3):
            for start in range(0, len(tokens) - size + 1):
                phrase_tokens = tokens[start : start + size]
                phrase_canonical = canonical[start : start + size]
                cue_count = sum(token in NEGATIVE_CUES for token in phrase_canonical)
                if cue_count == 0:
                    continue
                if size > 1:
                    if phrase_tokens[0] in BOUNDARY_STOPWORDS and phrase_canonical[0] not in NEGATION_WORDS:
                        continue
                    if phrase_tokens[-1] in BOUNDARY_STOPWORDS:
                        continue
                    if phrase_canonical[-1] in NEGATION_WORDS:
                        continue
                if size == 1 and phrase_canonical[0] in NEGATION_WORDS:
                    continue
                length_bonus = {1: 0.5, 2: 3.0, 3: 2.0}[size]
                add_candidate(candidates, phrase_tokens, section_bonus + cue_count * 4.0 + length_bonus)
    return candidates


def phrase_occurrences(review: dict[str, Any], key: str) -> int:
    tokens = [canonical_token(token) for token in tokenize(f"{review.get('title', '')} {review.get('body', '')}")]
    phrase = key.split()
    if not phrase or len(tokens) < len(phrase):
        return 0
    return sum(tokens[index : index + len(phrase)] == phrase for index in range(len(tokens) - len(phrase) + 1))


def analyze_pain_phrases(
    reviews: list[dict[str, Any]],
    *,
    max_terms_per_review: int = 4,
) -> tuple[list[list[str]], list[FrequencyRow]]:
    candidate_sets = [review_candidates(review) for review in reviews]
    document_frequency = Counter(key for candidates in candidate_sets for key in candidates)
    selected: list[list[Candidate]] = []

    for candidates in candidate_sets:
        ranked = sorted(
            candidates.values(),
            key=lambda candidate: (
                candidate.score + math.log1p(document_frequency[candidate.key]) * 3,
                document_frequency[candidate.key],
                len(candidate.key.split()),
                candidate.surface,
            ),
            reverse=True,
        )
        chosen: list[Candidate] = []
        for candidate in ranked:
            candidate_tokens = set(candidate.key.split())
            if any(
                candidate.key in existing.key
                or existing.key in candidate.key
                or (len(candidate_tokens) == 1 and candidate_tokens <= set(existing.key.split()))
                for existing in chosen
            ):
                continue
            chosen.append(candidate)
            if len(chosen) >= max_terms_per_review:
                break
        selected.append(chosen)

    coverage = Counter()
    occurrences = Counter()
    surfaces: dict[str, Counter[str]] = defaultdict(Counter)
    for review, phrases in zip(reviews, selected):
        for phrase in phrases:
            coverage[phrase.key] += 1
            occurrences[phrase.key] += max(1, phrase_occurrences(review, phrase.key))
            surfaces[phrase.key][phrase.surface] += 1

    total_reviews = len(reviews)
    frequency_rows = [
        FrequencyRow(
            phrase=surfaces[key].most_common(1)[0][0],
            review_count=review_count,
            review_ratio=review_count / total_reviews if total_reviews else 0,
            occurrence_count=occurrences[key],
        )
        for key, review_count in coverage.items()
    ]
    frequency_rows.sort(
        key=lambda item: (-item.review_count, -item.occurrence_count, item.phrase)
    )
    return [[candidate.surface for candidate in phrases] for phrases in selected], frequency_rows
