import base64
import importlib.util
import sys
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).with_name("sellersprite_playwright_runner.py")
SPEC = importlib.util.spec_from_file_location("sellersprite_playwright_runner_test", SCRIPT_PATH)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


def png_header(width: int, height: int) -> str:
    data = b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR" + width.to_bytes(4, "big") + height.to_bytes(4, "big")
    return base64.b64encode(data).decode("ascii")


class ChartExportDimensionTests(unittest.TestCase):
    def test_reads_expected_png_dimensions(self):
        encoded = png_header(MODULE.TREND_EXPORT_PNG_WIDTH, MODULE.TREND_EXPORT_PNG_HEIGHT)

        self.assertEqual(
            MODULE.chart_png_dimensions(encoded),
            (MODULE.TREND_EXPORT_PNG_WIDTH, MODULE.TREND_EXPORT_PNG_HEIGHT),
        )

    def test_accepts_data_url_prefix(self):
        encoded = "data:image/png;base64," + png_header(640, 480)

        self.assertEqual(MODULE.chart_png_dimensions(encoded), (640, 480))

    def test_rejects_non_png_data(self):
        encoded = base64.b64encode(b"not a png").decode("ascii")

        with self.assertRaises(ValueError):
            MODULE.chart_png_dimensions(encoded)


if __name__ == "__main__":
    unittest.main()
