from __future__ import annotations

import json
import multiprocessing
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from shared.runtime_monitor import (
    MONITOR_LOG_ENV,
    MONITOR_PATH_ENV,
    MONITOR_RUN_STARTED_ENV,
    MONITOR_WRITE_LOCK_SUFFIX,
    initialize_monitor,
    monitor_event,
)


def _write_monitor_events(
    jsonl_path: str,
    text_log_path: str,
    run_started: str,
    worker_index: int,
    event_count: int,
    start_event,
) -> None:
    """在独立进程中写入足够长的事件，放大并发追加冲突。"""
    os.environ[MONITOR_PATH_ENV] = jsonl_path
    os.environ[MONITOR_LOG_ENV] = text_log_path
    os.environ[MONITOR_RUN_STARTED_ENV] = run_started
    if not start_event.wait(timeout=10):
        raise RuntimeError("并发写测试未收到启动信号")

    payload = "并发监控内容" * 512
    for sequence in range(event_count):
        monitor_event(
            "concurrent_monitor_event",
            worker=worker_index,
            sequence=sequence,
            message=f"{payload}-{worker_index}-{sequence}",
        )


class RuntimeMonitorTests(unittest.TestCase):
    def test_concurrent_processes_append_complete_lines(self):
        worker_count = 6
        events_per_worker = 40
        expected_count = worker_count * events_per_worker

        with patch.dict(os.environ, {}, clear=False), tempfile.TemporaryDirectory() as temp:
            temp_path = Path(temp)
            jsonl_path = temp_path / "workflow_monitor.jsonl"
            text_log_path = temp_path / "workflow_monitor.log"
            initialize_monitor(jsonl_path, text_log_path)
            run_started = os.environ[MONITOR_RUN_STARTED_ENV]

            context = multiprocessing.get_context("spawn")
            start_event = context.Event()
            processes = [
                context.Process(
                    target=_write_monitor_events,
                    args=(
                        str(jsonl_path),
                        str(text_log_path),
                        run_started,
                        worker_index,
                        events_per_worker,
                        start_event,
                    ),
                )
                for worker_index in range(worker_count)
            ]
            for process in processes:
                process.start()
            start_event.set()
            for process in processes:
                process.join(timeout=30)

            alive_processes = [process for process in processes if process.is_alive()]
            for process in alive_processes:
                process.terminate()
                process.join(timeout=5)
            self.assertEqual(alive_processes, [], "并发写子进程未按时退出")
            for process in processes:
                self.assertEqual(process.exitcode, 0)

            jsonl_lines = jsonl_path.read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(jsonl_lines), expected_count)
            payloads = [json.loads(line) for line in jsonl_lines]
            self.assertEqual(
                {(item["worker"], item["sequence"]) for item in payloads},
                {
                    (worker_index, sequence)
                    for worker_index in range(worker_count)
                    for sequence in range(events_per_worker)
                },
            )
            self.assertTrue(
                all(item["event"] == "concurrent_monitor_event" for item in payloads)
            )

            text_lines = text_log_path.read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(text_lines), expected_count)
            self.assertTrue(
                all("concurrent_monitor_event" in line for line in text_lines)
            )
            self.assertTrue(
                Path(str(jsonl_path) + MONITOR_WRITE_LOCK_SUFFIX).exists()
            )


if __name__ == "__main__":
    unittest.main()
