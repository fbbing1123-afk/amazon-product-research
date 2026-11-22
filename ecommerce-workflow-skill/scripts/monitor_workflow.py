from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.error import URLError
from urllib.request import urlopen


ACTIVITY_EVENTS = {"command_output", "file_activity"}
END_EVENTS = {"command_finished", "command_failed", "command_timeout"}
WORKFLOW_END_EVENTS = {"workflow_finished", "scripted_stages_finished"}


def configure_console_encoding() -> None:
    """让 Windows 控制台安全输出中文诊断信息。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


def load_events(path: Path) -> list[dict]:
    """读取监控事件，忽略运行中尚未写完的最后一行。"""
    if not path.exists():
        return []
    events: list[dict] = []
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            events.append(value)
    return events


def event_time(event: dict) -> datetime | None:
    """读取事件的本地时间戳。"""
    raw = str(event.get("timestamp") or "")
    try:
        return datetime.fromisoformat(raw)
    except ValueError:
        return None


def epoch(event: dict) -> float | None:
    value = event_time(event)
    return value.timestamp() if value else None


def format_seconds(value: float | int | None) -> str:
    """把秒数转换为便于人工阅读的时长。"""
    if value is None:
        return "-"
    seconds = max(0, int(round(float(value))))
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    if hours:
        return f"{hours}时{minutes:02d}分{seconds:02d}秒"
    if minutes:
        return f"{minutes}分{seconds:02d}秒"
    return f"{seconds}秒"


def process_alive(pid: int | None) -> bool | None:
    """检查监控记录中的子进程是否仍存在，不触碰其它进程。"""
    if not pid:
        return None
    if os.name == "nt":
        result = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/NH"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
        )
        return str(pid) in result.stdout
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def probe_cdp(address: str) -> str:
    """读取 CDP 页面数量，辅助区分脚本等待和浏览器页面失联。"""
    if not address:
        return "未配置 CDP"
    base = address if address.startswith("http") else f"http://{address}"
    try:
        with urlopen(f"{base.rstrip('/')}/json/list", timeout=2) as response:
            targets = json.loads(response.read().decode("utf-8", errors="replace"))
        pages = [item for item in targets if item.get("type") == "page"]
        labels = []
        for item in pages[:3]:
            title = str(item.get("title") or item.get("url") or "")
            labels.append(title.replace("\n", " ")[:50])
        suffix = "；".join(labels)
        return f"CDP页面={len(pages)}" + (f"（{suffix}）" if suffix else "")
    except (OSError, URLError, ValueError, json.JSONDecodeError) as exc:
        return f"CDP不可用（{type(exc).__name__}）"


def command_records(events: list[dict]) -> list[dict]:
    """把离散监控事件合并成每条命令的一条诊断记录。"""
    records: list[dict] = []
    active: dict[str, dict] = {}
    for event in events:
        event_name = event.get("event")
        step = str(event.get("command_step") or "")
        if event_name == "command_started" and step:
            record = {
                "step": step,
                "start": epoch(event),
                "start_elapsed": event.get("elapsed_seconds"),
                "pid": event.get("pid"),
                "duration": None,
                "status": "运行中",
                "outputs": 0,
                "files": 0,
                "heartbeats": 0,
                "warnings": [],
                "max_idle": 0.0,
                "last_activity": epoch(event),
            }
            records.append(record)
            active[step] = record
            continue

        record = active.get(step)
        if record is None:
            continue
        if event_name in ACTIVITY_EVENTS:
            record["last_activity"] = epoch(event) or record["last_activity"]
            if event_name == "command_output":
                record["outputs"] += 1
            else:
                record["files"] += 1
        elif event_name == "command_heartbeat":
            record["heartbeats"] += 1
            idle = float(event.get("idle_seconds") or 0)
            record["max_idle"] = max(record["max_idle"], idle)
        elif event_name == "inactivity_warning":
            idle = float(event.get("idle_seconds") or 0)
            record["warnings"].append(idle)
            record["max_idle"] = max(record["max_idle"], idle)
        elif event_name in END_EVENTS:
            duration = event.get("duration_seconds")
            if duration is not None:
                record["duration"] = float(duration)
            if event_name == "command_finished":
                record["status"] = "完成"
            elif event_name == "command_timeout":
                record["status"] = "超时"
            else:
                record["status"] = "失败"
                record["error"] = event.get("reason") or event.get("return_code")
            active.pop(step, None)
    for record in records:
        if record["duration"] is None and record["start"] is not None:
            last = events[-1] if events else {}
            last_time = epoch(last) or record["start"]
            record["duration"] = max(0.0, last_time - record["start"])
    return records


def print_summary(path: Path, events: list[dict], cdp: str = "") -> None:
    """输出历史运行的耗时、告警和失败原因。"""
    records = command_records(events)
    first = next((item for item in events if item.get("event") == "workflow_started"), None)
    last = events[-1] if events else None
    total = last.get("elapsed_seconds") if last else None
    print(f"监控文件：{path}")
    print(f"事件数量：{len(events)}")
    print(f"监控范围总耗时：{format_seconds(total)}")
    if first:
        print(f"站点：{first.get('site', '-')}；邮编：{first.get('postal_code', '-')}；步骤：{first.get('steps', '-')}")
    print("\n命令耗时明细：")
    print("阶段".ljust(38) + "耗时".ljust(12) + "状态".ljust(8) + "输出/文件/心跳".ljust(18) + "最长无进展")
    for record in records:
        counts = f"{record['outputs']}/{record['files']}/{record['heartbeats']}"
        print(
            str(record["step"]).ljust(38)
            + format_seconds(record["duration"]).ljust(12)
            + str(record["status"]).ljust(8)
            + counts.ljust(18)
            + format_seconds(record["max_idle"])
        )

    warnings = [item for item in events if item.get("event") == "inactivity_warning"]
    failures = [item for item in events if item.get("event") in {"command_failed", "command_timeout"}]
    print(f"\n无进展告警：{len(warnings)}；失败/超时：{len(failures)}")
    for item in warnings:
        print(
            f"[告警] +{item.get('elapsed_seconds', '?')}秒 "
            f"{item.get('command_step', '-')} 无进展 {item.get('idle_seconds', '?')}秒"
        )
    for item in failures:
        print(
            f"[失败] +{item.get('elapsed_seconds', '?')}秒 "
            f"{item.get('command_step', '-')} 原因={item.get('reason') or item.get('return_code', '-') }"
        )

    if records:
        slowest = sorted(records, key=lambda item: item.get("duration") or 0, reverse=True)[:3]
        print("\n耗时最长的步骤：")
        for record in slowest:
            print(f"- {record['step']}：{format_seconds(record['duration'])}")
    if cdp:
        print(f"\n当前 CDP：{probe_cdp(cdp)}")
    if last:
        print(f"最后事件：+{last.get('elapsed_seconds', '?')}秒 {last.get('event', '-')}")


def print_live_status(path: Path, events: list[dict], cdp: str = "") -> None:
    """输出当前仍在运行的命令和最近一次真实活动。"""
    records = command_records(events)
    active = [record for record in records if record["status"] == "运行中"]
    if not active:
        last = events[-1] if events else {}
        print(f"[实时] 流程没有运行中命令；最后事件={last.get('event', '暂无事件')}")
        return
    record = active[-1]
    now = time.time()
    last_activity = record.get("last_activity") or record.get("start") or now
    idle = max(0.0, now - last_activity)
    alive = process_alive(record.get("pid"))
    cdp_text = f"；{probe_cdp(cdp)}" if cdp else ""
    print(
        f"[实时] {record['step']} 已运行 {format_seconds(now - (record['start'] or now))}；"
        f"PID={record.get('pid', '-')} 存活={alive}；"
        f"最近真实活动前 {format_seconds(idle)}；"
        f"输出={record['outputs']} 文件变化={record['files']} 心跳={record['heartbeats']}"
        f"{cdp_text}"
    )


def follow(path: Path, interval: float, cdp: str) -> int:
    """持续跟踪监控文件，流程结束后自动输出完整摘要。"""
    last_count = -1
    while True:
        events = load_events(path)
        if len(events) != last_count:
            print_live_status(path, events, cdp)
            last_count = len(events)
        if events and events[-1].get("event") in WORKFLOW_END_EVENTS:
            print_summary(path, events, cdp)
            return 0
        time.sleep(max(1.0, interval))


def main() -> int:
    configure_console_encoding()
    parser = argparse.ArgumentParser(description="实时诊断亚马逊选品自动化工作流的耗时和卡点。")
    parser.add_argument("monitor", type=Path, help="*_workflow_monitor.jsonl 文件")
    parser.add_argument("--follow", action="store_true", help="持续跟踪运行中的监控文件，结束后自动汇总")
    parser.add_argument("--interval", type=float, default=5.0, help="实时跟踪的轮询间隔秒数")
    parser.add_argument("--cdp", default="", help="可选 CDP 地址，例如 127.0.0.1:9223")
    args = parser.parse_args()
    path = args.monitor.expanduser().resolve()
    if args.follow:
        return follow(path, args.interval, args.cdp)
    events = load_events(path)
    if not events:
        print(f"监控文件不存在或为空：{path}")
        return 1
    print_summary(path, events, args.cdp)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
