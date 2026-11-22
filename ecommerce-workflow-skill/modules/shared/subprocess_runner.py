from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path
from typing import Sequence

from shared.runtime_monitor import monitor_event


DEFAULT_INACTIVITY_TIMEOUT_SECONDS = 180
DEFAULT_COMMAND_TIMEOUT_SECONDS = 1800
HEARTBEAT_INTERVAL_SECONDS = 15


def _configure_console_encoding() -> None:
    """让 Windows 控制台能够安全输出网页返回的 Unicode 文本。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


_configure_console_encoding()


def _path_signatures(paths: Sequence[str | Path]) -> dict[str, int]:
    """读取监控路径的修改时间，目录变化可捕获新生成的截图文件。"""
    signatures: dict[str, int] = {}
    for raw_path in paths:
        path = Path(raw_path)
        try:
            signatures[str(path)] = path.stat().st_mtime_ns
        except OSError:
            signatures[str(path)] = -1
    return signatures


def _terminate_process_tree(process: subprocess.Popen) -> None:
    """只终止当前命令及其子进程，不影响其它浏览器或 Python 进程。"""
    if process.poll() is not None:
        return
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(process.pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    else:
        process.kill()


def run_command(
    step: str,
    command: Sequence[str | Path],
    cwd: str | Path | None = None,
    *,
    watch_paths: Sequence[str | Path] = (),
    inactivity_timeout: float = DEFAULT_INACTIVITY_TIMEOUT_SECONDS,
    command_timeout: float = DEFAULT_COMMAND_TIMEOUT_SECONDS,
) -> None:
    printable = " ".join(str(part) for part in command)
    print(f"[{step}] {printable}")
    started_at = time.perf_counter()
    env = os.environ.copy()
    env.setdefault("PYTHONUTF8", "1")
    env.setdefault("PYTHONIOENCODING", "utf-8")
    env.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")
    output_queue: queue.Queue[str | None] = queue.Queue()
    process = subprocess.Popen(
        [str(part) for part in command],
        cwd=str(cwd) if cwd else None,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    monitor_event(
        "command_started",
        command_step=step,
        command=printable,
        pid=process.pid,
        watched_paths=", ".join(str(path) for path in watch_paths),
        inactivity_timeout_seconds=inactivity_timeout,
        command_timeout_seconds=command_timeout,
    )

    def read_output() -> None:
        """独立读取子进程输出，避免主线程被 readline 永久阻塞。"""
        assert process.stdout is not None
        try:
            for line in process.stdout:
                output_queue.put(line.rstrip("\r\n"))
        finally:
            output_queue.put(None)

    reader = threading.Thread(target=read_output, name=f"monitor-{step}", daemon=True)
    reader.start()
    started_monotonic = time.monotonic()
    last_progress_at = started_monotonic
    last_warning_at = 0.0
    last_heartbeat_at = started_monotonic
    signatures = _path_signatures(watch_paths)
    output_closed = False
    termination_reason = ""
    while not output_closed or process.poll() is None:
        current_monotonic = time.monotonic()
        current_signatures = _path_signatures(watch_paths)
        if current_signatures != signatures:
            changed_paths = [
                path for path, value in current_signatures.items()
                if signatures.get(path) != value
            ]
            signatures = current_signatures
            last_progress_at = current_monotonic
            monitor_event(
                "file_activity",
                command_step=step,
                paths=", ".join(changed_paths),
            )

        command_elapsed = current_monotonic - started_monotonic
        idle_seconds = current_monotonic - last_progress_at
        if current_monotonic - last_heartbeat_at >= HEARTBEAT_INTERVAL_SECONDS:
            monitor_event(
                "command_heartbeat",
                command_step=step,
                pid=process.pid,
                duration_seconds=round(command_elapsed, 1),
                idle_seconds=round(idle_seconds, 1),
                process_alive=process.poll() is None,
                watched_paths=", ".join(str(path) for path in watch_paths),
            )
            last_heartbeat_at = current_monotonic
        if command_elapsed >= command_timeout:
            termination_reason = "command_timeout"
            monitor_event(
                "command_timeout",
                command_step=step,
                timeout_seconds=command_timeout,
                duration_seconds=round(command_elapsed, 1),
            )
            _terminate_process_tree(process)
            break
        if idle_seconds >= inactivity_timeout:
            termination_reason = "inactivity_timeout"
            monitor_event(
                "inactivity_timeout",
                command_step=step,
                idle_seconds=round(idle_seconds, 1),
                duration_seconds=round(command_elapsed, 1),
            )
            _terminate_process_tree(process)
            break
        try:
            line = output_queue.get(timeout=1.0)
        except queue.Empty:
            if idle_seconds >= 60 and idle_seconds - last_warning_at >= 60:
                print(f"[{step}] warning: no progress for {idle_seconds:.0f}s", flush=True)
                monitor_event(
                    "inactivity_warning",
                    command_step=step,
                    idle_seconds=round(idle_seconds, 1),
                    watched_paths=", ".join(str(path) for path in watch_paths),
                )
                last_warning_at = idle_seconds
            continue
        if line is None:
            output_closed = True
            continue
        print(line, flush=True)
        last_progress_at = time.monotonic()
        monitor_event("command_output", command_step=step, message=line)

    reader.join(timeout=2)
    return_code = process.wait()
    elapsed = time.perf_counter() - started_at
    print(f"[{step}] elapsed: {elapsed:.1f}s")
    if termination_reason:
        monitor_event(
            "command_failed",
            command_step=step,
            pid=process.pid,
            return_code=return_code,
            duration_seconds=round(elapsed, 1),
            reason=termination_reason,
        )
        raise RuntimeError(
            f"Step stopped ({step}) because {termination_reason} after {elapsed:.1f}s"
        )
    monitor_event(
        "command_finished" if return_code == 0 else "command_failed",
        command_step=step,
        pid=process.pid,
        return_code=return_code,
        duration_seconds=round(elapsed, 1),
    )
    if return_code != 0:
        raise RuntimeError(f"Step failed ({step}) with exit code {return_code}")
