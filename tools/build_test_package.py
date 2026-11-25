"""从已提交源码生成仅含当前流程的同事测试包，不清理开发机环境。"""

import argparse
import hashlib
import io
import json
import re
import subprocess
from datetime import datetime, timezone, timedelta
from pathlib import Path, PurePosixPath
from zipfile import ZipFile, ZIP_DEFLATED


REPO = Path(__file__).resolve().parents[1]
ROOT_FILES = {"README.md", "AGENTS.md", "codex.md", "开始使用.txt", ".gitignore", ".gitattributes"}
SKILL = "ecommerce-workflow-skill/"
TEMPLATE = SKILL + "assets/产品调研模板.xlsx"
OMIT = {
    SKILL + "docs/使用指南.docx": "旧版使用指引，当前只提供最新版说明",
    SKILL + "modules/competitor_analysis/agent_prompt.md": "旧版竞品角色，不用于当前评论分析",
    SKILL + "modules/market_analysis/判断依据说明.md": "独立市场分析页的历史说明",
    SKILL + "resources/review_translations.json": "旧评价翻译缓存，不随测试包发送",
    SKILL + "resources/pain_phrase_groups.json": "旧评价语义缓存，不随测试包发送",
}
REQUIRED = {
    "README.md", "开始使用.txt", "codex.md", TEMPLATE,
    SKILL + "SKILL.md", SKILL + "docs/使用指南.md", SKILL + "docs/模板填写说明.md",
    SKILL + "docs/delivery_checks.md", SKILL + "scripts/runtime_environment.ps1",
    SKILL + "scripts/setup_first_run.ps1", SKILL + "scripts/check_environment.ps1",
    SKILL + "scripts/build_amazon_collection_link.py", SKILL + "scripts/run_review_prefetch.py",
    SKILL + "modules/product_research/scripts/chart-runtime/vendor/echarts.min.js",
}
FORBIDDEN_PARTS = {".git", ".venv", "venv", "node_modules", "__pycache__", ".workflow-state", "outputs", "downloads", "cache", "work"}
FORBIDDEN_SUFFIXES = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv", ".png", ".log", ".jsonl", ".pyc", ".pem", ".key"}


def git(*args):
    return subprocess.check_output(["git", *args], cwd=REPO)


def include(name):
    if name in OMIT:
        return False
    return name in ROOT_FILES or name.startswith(SKILL)


def check_file(name, data):
    """检查打包边界；模板是唯一经用户确认的Excel例外。"""
    path = PurePosixPath(name)
    if path.is_absolute() or ".." in path.parts:
        raise ValueError(f"不安全的包内路径：{name}")
    if any(part in FORBIDDEN_PARTS or "profile" in part.lower() for part in path.parts):
        raise ValueError(f"不能打包运行目录：{name}")
    if path.suffix.lower() in FORBIDDEN_SUFFIXES and name != TEMPLATE:
        raise ValueError(f"不能打包业务或敏感文件：{name}")
    if path.name.startswith(".env") or re.search(r"cookie|storage.state|seller_sprite_config", path.name, re.I):
        raise ValueError(f"不能打包本机账号配置：{name}")
    if path.suffix in {".py", ".ps1", ".mjs", ".json", ".yaml", ".md", ".txt"}:
        # 仅报告文件名，不将疑似凭据打印到终端。
        text = data.decode("utf-8-sig")
        if re.search(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----|gh[pousr]_[A-Za-z0-9]{30,}|sk-[A-Za-z0-9_-]{40,}", text):
            raise ValueError(f"疑似凭据，停止打包：{name}")


def build(output_dir, version):
    if not re.fullmatch(r"[A-Za-z0-9._-]+", version):
        raise ValueError("版本号只能使用字母、数字、点、下划线和短横线")
    if git("status", "--porcelain").strip():
        raise ValueError("先核对并提交当前改动，再打包，确保包内源码可以定位")
    commit = git("rev-parse", "HEAD").decode().strip()
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"amazon-product-research-{version}.zip"
    if target.exists():
        raise FileExistsError(f"输出已存在，不覆盖：{target}")
    prefix = f"amazon-product-research-{version}/"
    files = {}
    with ZipFile(io.BytesIO(git("archive", "--format=zip", "HEAD"))) as archive:
        for entry in archive.infolist():
            if entry.is_dir() or not include(entry.filename):
                continue
            if ((entry.external_attr >> 16) & 0o170000) == 0o120000:
                raise ValueError(f"不能分发本机符号链接：{entry.filename}")
            data = archive.read(entry)
            check_file(entry.filename, data)
            files[entry.filename] = data
    missing = REQUIRED - files.keys()
    if missing:
        raise ValueError(f"缺少当前流程文件：{sorted(missing)}")
    manifest = {
        "version": version, "source_commit": commit,
        "built_at": datetime.now(timezone(timedelta(hours=8))).isoformat(),
        "workflow": "当前最新版，按SKILL.md执行",
        "template": TEMPLATE, "file_count": len(files),
        "excluded": OMIT,
        "files": {name: {"sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data)} for name, data in sorted(files.items())},
    }
    temporary = target.with_suffix(".part")
    try:
        with ZipFile(temporary, "x", compression=ZIP_DEFLATED) as archive:
            for name, data in sorted(files.items()):
                archive.writestr(prefix + name, data)
            archive.writestr(prefix + "PACKAGE-MANIFEST.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        with ZipFile(temporary) as archive:
            if archive.testzip():
                raise ValueError("压缩包完整性校验失败")
            for name, item in manifest["files"].items():
                if hashlib.sha256(archive.read(prefix + name)).hexdigest() != item["sha256"]:
                    raise ValueError(f"包内文件指纹不一致：{name}")
        temporary.replace(target)
    finally:
        if temporary.exists():
            temporary.unlink()
    digest = hashlib.sha256(target.read_bytes()).hexdigest()
    target.with_suffix(".zip.sha256").write_text(f"{digest}  {target.name}\n", encoding="utf-8")
    return {"package": str(target.resolve()), "source_commit": commit, "files": len(files), "bytes": target.stat().st_size, "sha256": digest}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--version", required=True)
    args = parser.parse_args()
    print(json.dumps(build(args.output_dir.resolve(), args.version), ensure_ascii=False, indent=2))
