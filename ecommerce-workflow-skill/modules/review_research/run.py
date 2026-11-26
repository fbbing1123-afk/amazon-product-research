from __future__ import annotations

from pathlib import Path

from shared.browser_runtime import get_debugger_address
from shared.subprocess_runner import (
    DEFAULT_COMMAND_TIMEOUT_SECONDS,
    DEFAULT_INACTIVITY_TIMEOUT_SECONDS,
    run_command,
)
from shared.runtime_monitor import monitor_event

from .config import MODULE_DIR, SCRAPER_PATH
from .scrape_amazon_reviews import extract_review_asins_from_workbook


REVIEW_JOB_TIMEOUT_SECONDS = 10 * 60
REVIEW_STAGE_OVERHEAD_SECONDS = 10 * 60


def calculate_review_command_timeout(asin_count: int, star_filter: str) -> int:
    """按评论任务数计算总超时，同时保留原有 30 分钟下限。"""
    filter_count = 2 if star_filter == "both" else 1
    job_count = max(0, asin_count) * filter_count
    calculated_timeout = (
        job_count * REVIEW_JOB_TIMEOUT_SECONDS + REVIEW_STAGE_OVERHEAD_SECONDS
    )
    return max(DEFAULT_COMMAND_TIMEOUT_SECONDS, calculated_timeout)


def run_review_research(
    *,
    source: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    site: str = "US",
    amazon_postal_code: str | None = None,
    debugger_address: str = get_debugger_address(),
    asins: list[str] | None = None,
    asin_limit: int | None = None,
    max_pages: int = 10,
    max_reviews: int | None = None,
    max_positive_reviews: int = 50,
    max_critical_reviews: int = 100,
    sort_by: str = "helpful",
    reviewer_type: str = "default",
    trusted_only: bool = True,
    star_filter: str = "both",
) -> dict[str, Path]:
    print("[review_research] start")
    review_asins = list(dict.fromkeys(asins or extract_review_asins_from_workbook(source)))
    effective_asin_count = len(review_asins)
    if asin_limit is not None:
        effective_asin_count = min(effective_asin_count, max(0, asin_limit))
    filter_count = 2 if star_filter == "both" else 1
    review_job_count = effective_asin_count * filter_count
    command_timeout = calculate_review_command_timeout(
        effective_asin_count,
        star_filter,
    )
    monitor_event(
        "review_stage_started",
        source=str(source),
        site=site,
        asin_count=effective_asin_count,
        asin_limit=asin_limit,
        job_count=review_job_count,
        command_timeout_seconds=command_timeout,
        positive_limit=max_positive_reviews,
        critical_limit=max_critical_reviews,
    )
    output_path = output_dir / f"{name}_with_product_reviews.xlsx"
    translation_report = output_dir / f"{name}_missing_review_translations.json"
    pain_group_report = output_dir / f"{name}_missing_pain_phrase_groups.json"
    command: list[str | Path] = [
        python_exe,
        SCRAPER_PATH,
        "--source",
        source,
        "--output",
        output_path,
        "--site",
        site,
        "--debugger-address",
        debugger_address,
        "--max-pages",
        str(max_pages),
        "--max-positive-reviews",
        str(max_positive_reviews),
        "--max-critical-reviews",
        str(max_critical_reviews),
        "--sort-by",
        sort_by,
        "--reviewer-type",
        reviewer_type,
        "--trusted-only" if trusted_only else "--no-trusted-only",
        "--star-filter",
        star_filter,
        "--translation-report",
        translation_report,
        "--pain-group-report",
        pain_group_report,
    ]
    if max_reviews is not None:
        command.extend(["--max-reviews", str(max_reviews)])
    if amazon_postal_code:
        command.extend(["--amazon-postal-code", amazon_postal_code])
    if asin_limit is not None:
        command.extend(["--asin-limit", str(asin_limit)])
    for asin in asins or []:
        command.extend(["--asin", asin])

    run_command(
        "review_research.scrape",
        command,
        cwd=MODULE_DIR,
        inactivity_timeout=DEFAULT_INACTIVITY_TIMEOUT_SECONDS,
        command_timeout=command_timeout,
    )
    monitor_event("review_stage_finished", output=str(output_path))
    return {
        "review_research_report": output_path,
        "review_research_missing_translations": translation_report,
        "review_research_missing_pain_phrase_groups": pain_group_report,
        "current_workbook": output_path,
    }
