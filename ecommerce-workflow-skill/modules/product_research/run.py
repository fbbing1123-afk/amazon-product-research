from __future__ import annotations

from pathlib import Path

from shared.browser_runtime import get_debugger_address
from shared.excel_inspect import inspect_image_risk
from shared.image_conversion import convert_wps_dispimg_to_floating
from shared.subprocess_runner import run_command

from .config import CHART_RUNTIME, SCRIPTS_DIR


def same_path(left: Path, right: Path) -> bool:
    try:
        return left.resolve() == right.resolve()
    except OSError:
        return False


def protect_workbook_images(
    *,
    path: Path,
    role: str,
    name: str,
    output_dir: Path,
    force: bool,
    result: dict[str, Path],
) -> Path:
    risk = inspect_image_risk(path)
    if not risk.has_dispimg:
        return path
    protected = output_dir / f"{name}_{role}_wps_images_floating.xlsx"
    print(f"[product_research] converting WPS DISPIMG images to floating images: {protected}")
    try:
        convert_wps_dispimg_to_floating(path, protected)
    except Exception as exc:
        message = (
            f"[product_research] WARNING: sheet '{risk.sheet_name}' has DISPIMG formulas "
            f"but they could not be converted to floating images: {exc}"
        )
        if not force:
            raise RuntimeError(message + " Re-run with --force to continue without guaranteed image preservation.") from exc
        print(message)
        return path
    result[f"product_research_{role}_wps_images_floating"] = protected
    return protected


def extract_source_records(
    *,
    source: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
) -> Path:
    records_json = output_dir / f"{name}_source_records.json"
    run_command(
        "product_research.extract",
        [python_exe, SCRIPTS_DIR / "extract_source_records.py", "--source", source, "--output", records_json],
        cwd=SCRIPTS_DIR.parent,
    )
    return records_json


def run_category_research(
    *,
    source: Path,
    template: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    node_exe: str,
    force: bool = False,
    capture_category_overviews: bool = True,
    category_overview_debugger_address: str = get_debugger_address(),
    category_overview_limit: int | None = 5,
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> dict[str, Path]:
    print("[category_research] start")
    protected_sources: dict[str, Path] = {}
    source_for_processing = protect_workbook_images(
        path=source,
        role="source",
        name=name,
        output_dir=output_dir,
        force=force,
        result=protected_sources,
    )
    if same_path(template, source):
        template_for_processing = source_for_processing
    else:
        template_for_processing = protect_workbook_images(
            path=template,
            role="template",
            name=name,
            output_dir=output_dir,
            force=force,
            result=protected_sources,
        )

    records_json = extract_source_records(
        source=source_for_processing,
        name=name,
        output_dir=output_dir,
        python_exe=python_exe,
    )
    filled_report = output_dir / f"{name}_report_filled.xlsx"
    category_overview_dir = output_dir / f"{name}_category_overviews"
    report_with_category_overviews = output_dir / f"{name}_report_with_category_overviews.xlsx"
    translation_report = output_dir / f"{name}_missing_category_translations.json"

    run_command(
        "category_research.fill_template",
        [
            python_exe,
            SCRIPTS_DIR / "fill_custom_template.py",
            "--template",
            template_for_processing,
            "--source",
            source_for_processing,
            "--output",
            filled_report,
            "--translation-report",
            translation_report,
        ],
        cwd=SCRIPTS_DIR.parent,
    )
    final_report = filled_report
    if capture_category_overviews:
        command = [
            python_exe,
            SCRIPTS_DIR / "capture_category_overviews.py",
            "--input",
            filled_report,
            "--output",
            report_with_category_overviews,
            "--image-dir",
            category_overview_dir,
            "--debugger-address",
            category_overview_debugger_address,
            "--node",
            node_exe,
            "--site",
            site,
        ]
        if amazon_postal_code:
            command.extend(["--amazon-postal-code", amazon_postal_code])
        if category_overview_limit is not None:
                command.extend(["--limit", str(category_overview_limit)])
        run_command(
            "category_research.category_overviews",
            command,
            cwd=SCRIPTS_DIR.parent,
            # 每个类目截图写入目录后都算真实进展，避免五个类目连续处理时被误判为卡死。
            watch_paths=(category_overview_dir,),
        )
        final_report = report_with_category_overviews
    run_command(
        "category_research.verify",
        [python_exe, SCRIPTS_DIR / "verify_custom_template.py", final_report],
        cwd=SCRIPTS_DIR.parent,
    )

    result = {
        "product_research_records": records_json,
        "product_research_filled": filled_report,
        "product_research_report": final_report,
        "product_research_missing_category_translations": translation_report,
    }
    result.update(protected_sources)
    if capture_category_overviews:
        result["product_research_category_overview_dir"] = category_overview_dir
    return result


def run_product_research_charts(
    *,
    source: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    node_exe: str,
) -> dict[str, Path]:
    print("[market_analysis_charts] start")
    records_json = extract_source_records(
        source=source,
        name=name,
        output_dir=output_dir,
        python_exe=python_exe,
    )
    chart_dir = output_dir / f"{name}_chart_images"
    report_with_charts = output_dir / f"{name}_report_with_charts.xlsx"
    run_command(
        "market_analysis.render_charts",
        [node_exe, CHART_RUNTIME, records_json, chart_dir],
        cwd=SCRIPTS_DIR.parent,
    )
    run_command(
        "market_analysis.insert_charts",
        [
            python_exe,
            SCRIPTS_DIR / "insert_chart_images.py",
            "--input",
            source,
            "--image-dir",
            chart_dir,
            "--output",
            report_with_charts,
        ],
        cwd=SCRIPTS_DIR.parent,
    )
    run_command(
        "market_analysis.verify_charts",
        [python_exe, SCRIPTS_DIR / "verify_custom_template.py", report_with_charts],
        cwd=SCRIPTS_DIR.parent,
    )
    return {
        "product_research_records": records_json,
        "market_analysis_chart_dir": chart_dir,
        "market_analysis_charts_report": report_with_charts,
        "current_workbook": report_with_charts,
    }


def run_product_research(
    *,
    source: Path,
    template: Path,
    name: str,
    output_dir: Path,
    python_exe: str,
    node_exe: str,
    force: bool = False,
    capture_category_overviews: bool = True,
    category_overview_debugger_address: str = get_debugger_address(),
    category_overview_limit: int | None = 5,
    site: str = "US",
    amazon_postal_code: str | None = None,
) -> dict[str, Path]:
    # Backward-compatible alias: product_research now means category research.
    return run_category_research(
        source=source,
        template=template,
        name=name,
        output_dir=output_dir,
        python_exe=python_exe,
        node_exe=node_exe,
        force=force,
        capture_category_overviews=capture_category_overviews,
        category_overview_debugger_address=category_overview_debugger_address,
        category_overview_limit=category_overview_limit,
        site=site,
        amazon_postal_code=amazon_postal_code,
    )
