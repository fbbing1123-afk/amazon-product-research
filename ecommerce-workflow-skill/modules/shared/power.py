from __future__ import annotations

import ctypes
import os
from collections.abc import Callable, Iterator
from contextlib import contextmanager
from dataclasses import dataclass

from shared.runtime_monitor import monitor_event


ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001


@dataclass
class WorkflowPowerHoldState:
    """记录当前工作流防休眠请求的状态。"""

    platform: str
    supported: bool
    active: bool = False
    request_result: int | None = None
    release_result: int | None = None
    error: str = ""


def _set_thread_execution_state(flags: int) -> int:
    """调用 Windows API 更新当前线程的执行状态。"""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    setter = kernel32.SetThreadExecutionState
    setter.argtypes = [ctypes.c_uint]
    setter.restype = ctypes.c_uint
    return int(setter(flags))


@contextmanager
def workflow_power_hold(
    *,
    os_name: str | None = None,
    set_execution_state: Callable[[int], int] | None = None,
) -> Iterator[WorkflowPowerHoldState]:
    """在工作流运行期间阻止 Windows 因系统空闲自动睡眠。"""
    platform_name = os_name if os_name is not None else os.name
    supported = platform_name == "nt"
    state = WorkflowPowerHoldState(
        platform=platform_name,
        supported=supported,
    )
    setter = set_execution_state or _set_thread_execution_state

    if supported:
        try:
            state.request_result = setter(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
            state.active = state.request_result != 0
            if not state.active:
                state.error = "SetThreadExecutionState 返回 0，未建立防休眠请求"
        except Exception as exc:
            state.error = f"{type(exc).__name__}: {exc}"

    monitor_event(
        "workflow_power_hold_started",
        platform=state.platform,
        supported=state.supported,
        active=state.active,
        request_result=state.request_result,
        error=state.error,
    )
    if state.active:
        print("[workflow] power hold: active (Windows 系统自动睡眠已阻止)", flush=True)
    elif state.supported:
        print(f"[workflow] power hold: unavailable ({state.error})", flush=True)
    else:
        print("[workflow] power hold: skipped (当前系统不是 Windows)", flush=True)

    try:
        yield state
    finally:
        release_success = False
        if state.active:
            try:
                state.release_result = setter(ES_CONTINUOUS)
                release_success = state.release_result != 0
                if not release_success:
                    state.error = "SetThreadExecutionState 恢复调用返回 0"
            except Exception as exc:
                state.error = f"{type(exc).__name__}: {exc}"
        monitor_event(
            "workflow_power_hold_finished",
            platform=state.platform,
            supported=state.supported,
            active=state.active,
            release_attempted=state.active,
            release_success=release_success,
            release_result=state.release_result,
            error=state.error,
        )
        if release_success:
            print("[workflow] power hold: released (Windows 电源状态已恢复)", flush=True)
        elif state.active:
            print(f"[workflow] power hold: release failed ({state.error})", flush=True)
        elif state.supported:
            print("[workflow] power hold: finished (未建立防休眠请求)", flush=True)
        else:
            print("[workflow] power hold: finished (当前系统无需恢复)", flush=True)
