from __future__ import annotations

import argparse
import json
from pathlib import Path


def load_events(path: Path) -> list[dict]:
    """读取 JSONL 监控事件，跳过损坏或不完整的末行。"""
    events: list[dict] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            events.append(value)
    return events


def main() -> int:
    parser = argparse.ArgumentParser(description="查看亚马逊选品自动化运行监控摘要。")
    parser.add_argument("monitor", type=Path, help="*_workflow_monitor.jsonl 文件")
    args = parser.parse_args()
    path = args.monitor.expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"监控文件不存在：{path}")

    events = load_events(path)
    if not events:
        print("监控文件为空，流程可能还没有写入第一条事件。")
        return 0

    print(f"监控文件：{path}")
    print(f"事件数量：{len(events)}")
    warnings = [item for item in events if item.get("event") == "inactivity_warning"]
    failures = [
        item
        for item in events
        if item.get("event") in {"command_failed", "review_job_failed"}
    ]
    jobs = [item for item in events if item.get("event") == "review_job_finished"]
    print(f"无输出告警：{len(warnings)}")
    print(f"失败事件：{len(failures)}")
    print(f"已完成评论任务：{len(jobs)}")

    for item in warnings[-10:]:
        print(
            "[告警] "
            f"+{item.get('elapsed_seconds', '?')}s "
            f"{item.get('command_step') or item.get('event')} "
            f"无输出 {item.get('idle_seconds', '?')}s"
        )
    for item in failures[-10:]:
        print(
            "[失败] "
            f"+{item.get('elapsed_seconds', '?')}s "
            f"asin={item.get('asin', '-')} "
            f"筛选={item.get('star_filter', '-')} "
            f"阶段={item.get('phase') or item.get('command_step', '-')} "
            f"原因={item.get('error') or item.get('return_code', '-')}"
        )

    last = events[-1]
    print(
        "最后事件："
        f" +{last.get('elapsed_seconds', '?')}s "
        f"{last.get('event')} "
        f"asin={last.get('asin', '-')} "
        f"筛选={last.get('star_filter', '-')}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

