"""
Domain Models — مذكرتي Pro v28
Pure data classes with validation. No I/O, no side effects.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional
import re as _re

# حد أقصى للحماية من DoS
_MAX_STR   = 2000
_MAX_LIST  = 10
_MAX_TITLE = 400

def _clean(s, max_len=_MAX_STR) -> str:
    """تنظيف وتقليص النص مع إزالة null bytes"""
    if not s: return ""
    return str(s).replace('\x00', '').strip()[:max_len]

def _clean_list(lst, max_items=_MAX_LIST, max_each=_MAX_STR) -> list[str]:
    if not lst: return []
    return [_clean(x, max_each) for x in lst[:max_items] if _clean(x)]


@dataclass
class SlideConfig:
    """Which slides to include"""
    cover: bool = True
    intro: bool = True
    plan: bool = True
    problem: bool = True
    objectives: bool = True
    importance: bool = True
    methodology: bool = True
    kpi: bool = True
    results: bool = True
    conclusion: bool = True
    recommendations: bool = True
    future: bool = True
    references: bool = True
    thankyou: bool = True

    @classmethod
    def from_dict(cls, d: dict) -> "SlideConfig":
        if not d:
            return cls()
        return cls(**{k: bool(v) for k, v in d.items() if k in cls.__dataclass_fields__})


@dataclass
class StatCard:
    label: str
    value: str
    unit: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> Optional["StatCard"]:
        if not d or not d.get("label") or not d.get("value"):
            return None
        return cls(
            label=_clean(d["label"], 120),
            value=_clean(d["value"], 60),
            unit=_clean(d.get("unit", ""), 40),
        )


@dataclass
class Chapter:
    title: str
    pages: str = ""

    @classmethod
    def from_dict(cls, d: dict) -> Optional["Chapter"]:
        if not d or not d.get("title"):
            return None
        return cls(
            title=_clean(d["title"], 200),
            pages=_clean(d.get("pages", ""), 40),
        )


@dataclass
class PresentationRequest:
    """Validated, normalized input from frontend"""
    # Required
    student_name: str
    title_ar: str

    # Optional metadata
    title_en: str = ""
    supervisor: str = ""
    co_supervisor: str = ""
    institution: str = ""
    year: str = ""
    specialization: str = ""
    lang: str = "ar"
    engine: str = "canva"
    theme: str = "navy_gold"

    # Content
    intro_overview: str = ""
    intro_approach: str = ""
    main_problem: str = ""
    main_question: str = ""
    sub_questions: list[str] = field(default_factory=list)
    objectives: list[str] = field(default_factory=list)
    hypotheses: list[str] = field(default_factory=list)
    importance: list[str] = field(default_factory=list)
    reasons: str = ""
    methodology: str = ""
    sample_type: str = ""
    sample_size: str = ""
    tool: str = ""
    stats: list[StatCard] = field(default_factory=list)
    main_results: list[str] = field(default_factory=list)
    general_conclusion: str = ""
    recommendations: list[str] = field(default_factory=list)
    future_work: list[str] = field(default_factory=list)
    references: list[str] = field(default_factory=list)
    chapters: list[Chapter] = field(default_factory=list)

    # Slide toggles
    slides: SlideConfig = field(default_factory=SlideConfig)

    # يُحسب من pipeline قبل البناء
    _total_slides: int = 13

    VALID_ENGINES = {"canva", "premium", "classic"}
    VALID_THEMES = {
        'navy_gold', 'dark_teal', 'burgundy', 'forest',
        'midnight_purple', 'charcoal_orange', 'ice_blue',
        'sand_gold', 'slate_crimson', 'noir', 'atlas', 'sakura'
    }

    @classmethod
    def from_dict(cls, raw: dict) -> "PresentationRequest":
        if not raw or not isinstance(raw, dict):
            return cls(student_name="", title_ar="")

        theme = str(raw.get("theme", "navy_gold"))
        if theme not in cls.VALID_THEMES:
            theme = "navy_gold"

        engine = str(raw.get("engine", "canva"))
        if engine not in cls.VALID_ENGINES:
            engine = "canva"

        stats    = [s for s in (StatCard.from_dict(x) for x in (raw.get("stats") or [])[:6]) if s]
        chapters = [c for c in (Chapter.from_dict(x) for x in (raw.get("chapters") or [])[:8]) if c]
        slides   = SlideConfig.from_dict(raw.get("slides") or {})

        return cls(
            student_name   = _clean(raw.get("studentName", ""), 120),
            title_ar       = _clean(raw.get("titleAr", ""),      _MAX_TITLE),
            title_en       = _clean(raw.get("titleEn", ""),      _MAX_TITLE),
            supervisor     = _clean(raw.get("supervisor", ""),   120),
            co_supervisor  = _clean(raw.get("coSupervisor", ""), 120),
            institution    = _clean(raw.get("institution", ""),  200),
            year           = _clean(raw.get("year", ""),         20),
            specialization = _clean(raw.get("specialization",""), 200),
            lang           = str(raw.get("lang", "ar"))[:5],
            engine         = engine,
            theme          = theme,
            intro_overview = _clean(raw.get("introOverview", ""), _MAX_STR),
            intro_approach = _clean(raw.get("introApproach", ""), _MAX_STR),
            main_problem   = _clean(raw.get("mainProblem",   ""), _MAX_STR),
            main_question  = _clean(raw.get("mainQuestion",  ""), _MAX_STR),
            sub_questions  = _clean_list(raw.get("subQuestions"), 5),
            objectives     = _clean_list(raw.get("objectives"),   6),
            hypotheses     = _clean_list(raw.get("hypotheses"),   5),
            importance     = _clean_list(raw.get("importance"),   6),
            reasons        = _clean(raw.get("reasons", ""),       _MAX_STR),
            methodology    = _clean(raw.get("methodology", ""),   400),
            sample_type    = _clean(raw.get("sampleType", ""),    200),
            sample_size    = _clean(raw.get("sampleSize", ""),    100),
            tool           = _clean(raw.get("tool", ""),          400),
            stats          = stats,
            main_results   = _clean_list(raw.get("mainResults"),  8, 400),
            general_conclusion = _clean(raw.get("generalConclusion", ""), _MAX_STR),
            recommendations    = _clean_list(raw.get("recommendations"), 8),
            future_work        = _clean_list(raw.get("futureWork"),      6),
            references         = _clean_list(raw.get("references"),      12, 400),
            chapters           = chapters,
            slides             = slides,
        )

    def validate(self) -> list[str]:
        errors = []
        if not self.student_name:
            errors.append("اسم الطالب مطلوب")
        if not self.title_ar:
            errors.append("عنوان المذكرة مطلوب")
        if len(self.student_name) > 120:
            errors.append("اسم الطالب طويل جداً")
        if len(self.title_ar) > _MAX_TITLE:
            errors.append("عنوان المذكرة طويل جداً")
        return errors
