from __future__ import annotations

import errno
import json
import os
import threading
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, BinaryIO, Iterator


MONITOR_PATH_ENV = "ECOMMERCE_WORKFLOW_MONITOR_PATH"
MONITOR_LOG_ENV = "ECOMMERCE_WORKFLOW_MONITOR_LOG"
MONITOR_RUN_STARTED_ENV = "ECOMMERCE_WORKFLOW_MONITOR_RUN_STARTED"
MONITOR_STATE_SUFFIX = ".state.json"
MONITOR_WRITE_LOCK_SUFFIX = ".write.lock"
_LOCK_RETRY_SECONDS = 0.01
_WRITE_LOCK = threading.Lock()


def _acquire_process_lock(handle: BinaryIO) -> None:
    """获取当前监控文件组的跨进程独占锁。"""
    handle.seek(0)
    if os.name == "nt":
        import msvcrt

        retryable_errors = {errno.EACCES, errno.EAGAIN, errno.EDEADLK}
        while True:
            try:
                msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                return
            except OSError as exc:
                if exc.errno not in retryable_errors:
                    raise
                time.sleep(_LOCK_RETRY_SECONDS)

    import fcntl

    fcntl.flock(handle.fileno(), fcntl.LOCK_EX)


def _release_process_lock(handle: BinaryIO) -> None:
    """释放当前监控文件组的跨进程独占锁。"""
    handle.seek(0)
    if os.name == "nt":
        import msvcrt

        msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        return

    import fcntl

    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


@contextmanager
def _process_write_lock(
    jsonl_value: str | None,
    text_value: str | None,
) -> Iterator[None]:
    """用同一把跨进程锁串行化 JSONL 与文本日志的成对追加。"""
    base_path = jsonl_value or text_value
    if not base_path:
        yield
        return

    lock_path = Path(str(base_path) + MONITOR_WRITE_LOCK_SUFFIX)
    lock_path.parent.mkdir(parents=True, exist_ok=True)
    with lock_path.open("a+b", buffering=0) as handle:
        # Windows 的字节区间锁需要一个稳定的可锁字节；锁文件不承载业务数据。
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
        _acquire_process_lock(handle)
        try:
            yield
        finally:
            _release_process_lock(handle)


def _append_utf8_line(path_value: str, value: str) -> None:
    """以二进制追加方式完整写入一行 UTF-8 日志。"""
    flags = os.O_APPEND | os.O_CREAT | os.O_WRONLY | getattr(os, "O_BINARY", 0)
    descriptor = os.open(path_value, flags, 0o666)
    try:
        data = (value + "\n").encode("utf-8")
        written = 0
        while written < len(data):
            count = os.write(descriptor, data[written:])
            if count <= 0:
                raise OSError("监控日志追加未写入任何字节")
            written += count
    finally:
        os.close(descriptor)


def initialize_monitor(jsonl_path: Path, text_log_path: Path) -> None:
    """初始化一次运行的监控文件，并让子进程复用同一组路径。"""
    jsonl_path.parent.mkdir(parents=True, exist_ok=True)
    text_log_path.parent.mkdir(parents=True, exist_ok=True)
    run_started = time.time()
    jsonl_path.write_text("", encoding="utf-8")
    text_log_path.write_text("", encoding="utf-8")
    # 交接到另一个进程后环境变量不会自动保留，用旁路状态文件保存原始起始时间。
    state_path = Path(str(jsonl_path) + MONITOR_STATE_SUFFIX)
    state_path.write_text(
        json.dumps({"run_started_epoch": run_started}, ensure_ascii=False),
        encoding="utf-8",
    )
    os.environ[MONITOR_PATH_ENV] = str(jsonl_path)
    os.environ[MONITOR_LOG_ENV] = str(text_log_path)
    os.environ[MONITOR_RUN_STARTED_ENV] = str(run_started)


def _read_persisted_run_started(jsonl_path: str | None) -> float | None:
    """从旁车状态文件或首个事件恢复跨进程的工作流起始时间。"""
    if not jsonl_path:
        return None
    state_path = Path(str(jsonl_path) + MONITOR_STATE_SUFFIX)
    try:
        state = json.loads(state_path.read_text(encoding="utf-8"))
        return float(state["run_started_epoch"])
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
        pass

    try:
        with Path(jsonl_path).open("r", encoding="utf-8") as handle:
            for line in handle:
                if not line.strip():
                    continue
                payload = json.loads(line)
                timestamp = datetime.fromisoformat(payload["timestamp"])
                return timestamp.timestamp()
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError):
        return None
    return None


def monitor_event(event: str, **fields: Any) -> None:
    """把结构化事件和一行人类可读日志同时写入监控文件。"""
    jsonl_value = os.environ.get(MONITOR_PATH_ENV)
    text_value = os.environ.get(MONITOR_LOG_ENV)
    if not jsonl_value and not text_value:
        return

    now = time.time()
    run_started_value = os.environ.get(MONITOR_RUN_STARTED_ENV)
    if run_started_value:
        try:
            run_started = float(run_started_value)
        except ValueError:
            run_started = now
    else:
        run_started = _read_persisted_run_started(jsonl_value) or now
    payload = {
        "timestamp": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "event": event,
        **fields,
    }
    # 累计时间是监控保留字段，任务自身耗时统一使用 duration_seconds。
    payload["elapsed_seconds"] = round(max(0.0, now - run_started), 1)
    line = json.dumps(payload, ensure_ascii=False, default=str)
    readable_fields = " ".join(
        f"{key}={str(value).replace(chr(10), ' ')[:500]}"
        for key, value in fields.items()
        if value is not None
    )
    readable = (
        f"{payload['timestamp']} +{payload['elapsed_seconds']:.1f}s "
        f"{event} {readable_fields}".rstrip()
    )
    with _WRITE_LOCK, _process_write_lock(jsonl_value, text_value):
        if jsonl_value:
            _append_utf8_line(jsonl_value, line)
        if text_value:
            _append_utf8_line(text_value, readable)
