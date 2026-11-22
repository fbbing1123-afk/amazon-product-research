from __future__ import annotations

import unittest

from shared.marketplaces import _amazon_confirmation_button, _first_visible_locator


class FakeElement:
    def __init__(self, *, visible: bool):
        self.visible = visible
        self.clicked = False

    def is_visible(self, *, timeout: int) -> bool:
        return self.visible

    def click(self, *, timeout: int) -> None:
        self.clicked = True


class FakeLocatorCollection:
    def __init__(self, elements: list[FakeElement]):
        self.elements = elements

    def count(self) -> int:
        return len(self.elements)

    def nth(self, index: int) -> FakeElement:
        return self.elements[index]


class FakePage:
    def __init__(
        self,
        elements: list[FakeElement] | dict[str, list[FakeElement]],
    ):
        self.elements = elements
        self.selectors: list[str] = []

    def locator(self, selector: str) -> FakeLocatorCollection:
        self.selectors.append(selector)
        if isinstance(self.elements, dict):
            return FakeLocatorCollection(self.elements.get(selector, []))
        return FakeLocatorCollection(self.elements)


class MarketplaceTests(unittest.TestCase):
    def test_first_visible_locator_skips_hidden_duplicate_id(self):
        hidden = FakeElement(visible=False)
        visible = FakeElement(visible=True)

        selected = _first_visible_locator(FakePage([hidden, visible]), "#duplicate")

        self.assertIs(selected, visible)

    def test_confirmation_click_uses_visible_duplicate(self):
        hidden = FakeElement(visible=False)
        visible = FakeElement(visible=True)
        page = FakePage([hidden, visible])

        confirmation = _amazon_confirmation_button(page)
        self.assertIs(confirmation, visible)
        confirmation.click(timeout=5000)

        self.assertFalse(hidden.clicked)
        self.assertTrue(visible.clicked)
        self.assertEqual(
            page.selectors,
            ["#GLUXConfirmClose, #GLUXConfirmClose-announce"],
        )

    def test_confirmation_falls_back_to_visible_done_after_success(self):
        hidden_confirmation = FakeElement(visible=False)
        visible_done = FakeElement(visible=True)
        page = FakePage(
            {
                "#GLUXConfirmClose, #GLUXConfirmClose-announce": [hidden_confirmation],
                "button[name='glowDoneButton'], input[name='glowDoneButton']": [visible_done],
            }
        )

        confirmation = _amazon_confirmation_button(page)

        self.assertIs(confirmation, visible_done)


if __name__ == "__main__":
    unittest.main()
