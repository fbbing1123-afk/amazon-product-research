from __future__ import annotations

import unittest
from unittest.mock import patch

from shared.power import (
    ES_CONTINUOUS,
    ES_SYSTEM_REQUIRED,
    workflow_power_hold,
)


class WorkflowPowerHoldTests(unittest.TestCase):
    def test_windows_hold_is_released_when_body_raises(self):
        calls: list[int] = []

        def fake_setter(flags: int) -> int:
            calls.append(flags)
            return ES_CONTINUOUS

        with (
            patch("shared.power.monitor_event") as monitor,
            patch("builtins.print"),
            self.assertRaisesRegex(RuntimeError, "测试异常"),
        ):
            with workflow_power_hold(
                os_name="nt",
                set_execution_state=fake_setter,
            ) as state:
                self.assertTrue(state.active)
                raise RuntimeError("测试异常")

        self.assertEqual(
            calls,
            [ES_CONTINUOUS | ES_SYSTEM_REQUIRED, ES_CONTINUOUS],
        )
        self.assertEqual(
            [call.args[0] for call in monitor.call_args_list],
            ["workflow_power_hold_started", "workflow_power_hold_finished"],
        )
        self.assertTrue(monitor.call_args_list[0].kwargs["active"])
        self.assertTrue(monitor.call_args_list[1].kwargs["release_success"])

    def test_non_windows_is_noop(self):
        calls: list[int] = []

        def fake_setter(flags: int) -> int:
            calls.append(flags)
            return ES_CONTINUOUS

        with (
            patch("shared.power.monitor_event") as monitor,
            patch("builtins.print"),
            workflow_power_hold(
                os_name="posix",
                set_execution_state=fake_setter,
            ) as state,
        ):
            self.assertFalse(state.supported)
            self.assertFalse(state.active)

        self.assertEqual(calls, [])
        self.assertFalse(monitor.call_args_list[0].kwargs["active"])
        self.assertFalse(monitor.call_args_list[1].kwargs["release_attempted"])

    def test_failed_windows_request_does_not_attempt_release(self):
        calls: list[int] = []

        def fake_setter(flags: int) -> int:
            calls.append(flags)
            return 0

        with (
            patch("shared.power.monitor_event") as monitor,
            patch("builtins.print"),
            workflow_power_hold(
                os_name="nt",
                set_execution_state=fake_setter,
            ) as state,
        ):
            self.assertFalse(state.active)
            self.assertIn("返回 0", state.error)

        self.assertEqual(calls, [ES_CONTINUOUS | ES_SYSTEM_REQUIRED])
        self.assertFalse(monitor.call_args_list[1].kwargs["release_attempted"])


if __name__ == "__main__":
    unittest.main()
