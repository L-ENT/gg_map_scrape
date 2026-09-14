import unittest
from unittest.mock import Mock, patch

import app


class BrowserPacingTests(unittest.TestCase):
    def make_browser(self):
        stop = Mock()
        stop.is_set.return_value = False
        stop.wait.return_value = False
        checkpoint = Mock()
        first, second = Mock(), Mock()
        factory = patch.object(app, "build_driver", side_effect=[first, second])
        factory.start()
        self.addCleanup(factory.stop)
        return app.PacedBrowser(False, stop, checkpoint), stop, checkpoint, first, second

    def test_restart_saves_closes_waits_and_restores_original_page_once(self):
        browser, stop, checkpoint, first, second = self.make_browser()
        browser.last_url = "https://www.google.com/maps/place/example"
        order = []
        checkpoint.side_effect = lambda: order.append("save")
        first.quit.side_effect = lambda: order.append("close")
        stop.wait.side_effect = lambda seconds: order.append(seconds) or False
        with patch.object(app, "maybe_accept_google_consent"), patch.object(app.random, "uniform", return_value=5):
            self.assertTrue(browser.restart_after_captcha(Mock()))
            self.assertTrue(browser.restart_after_captcha(Mock()))
        self.assertEqual(order, ["save", "close", 60, 5])
        second.get.assert_called_once_with(browser.last_url)
        second.quit.assert_not_called()
        self.assertIs(browser.raw, second)

    def test_stop_during_cooldown_does_not_reopen(self):
        browser, stop, checkpoint, first, second = self.make_browser()
        stop.wait.return_value = True
        self.assertFalse(browser.restart_after_captcha(Mock()))
        checkpoint.assert_called_once()
        first.quit.assert_called_once()
        stop.wait.assert_called_once_with(60)
        second.get.assert_not_called()
        self.assertIs(browser.raw, first)

    def test_navigation_delay_can_be_cancelled(self):
        browser, stop, _, first, _ = self.make_browser()
        stop.wait.return_value = True
        with patch.object(app.random, "uniform", return_value=7.5) as delay:
            with self.assertRaises(app.WebDriverException):
                browser.get("https://www.google.com/search?q=example")
        delay.assert_called_once_with(3, 8)
        stop.wait.assert_called_once_with(7.5)
        first.get.assert_not_called()

    def test_cleanup_blank_page_has_no_delay(self):
        browser, stop, _, first, _ = self.make_browser()
        browser.get("about:blank")
        stop.wait.assert_not_called()
        first.get.assert_called_once_with("about:blank")
