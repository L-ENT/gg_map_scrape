import unittest
from unittest.mock import Mock, patch

import app


class BrowserPacingTests(unittest.TestCase):
    def make_browser(self):
        stop = Mock()
        stop.is_set.return_value = False
        stop.wait.return_value = False
        checkpoint = Mock()
        drivers = [Mock() for _ in range(app.CAPTCHA_AUTO_RESTART_LIMIT + 1)]
        factory = patch.object(app, "build_driver", side_effect=drivers)
        factory.start()
        self.addCleanup(factory.stop)
        return app.PacedBrowser(False, stop, checkpoint), stop, checkpoint, drivers

    def test_first_four_captchas_restart_then_the_fifth_stays_open(self):
        browser, stop, checkpoint, drivers = self.make_browser()
        browser.last_url = "https://www.google.com/maps/place/example"
        with patch.object(app, "maybe_accept_google_consent"), patch.object(app.random, "uniform", return_value=5):
            for _ in range(app.CAPTCHA_AUTO_RESTART_LIMIT):
                self.assertTrue(browser.restart_after_captcha(Mock()))
            self.assertTrue(browser.restart_after_captcha(Mock()))
        self.assertEqual(checkpoint.call_count, app.CAPTCHA_AUTO_RESTART_LIMIT)
        for closed_driver in drivers[:-1]:
            closed_driver.quit.assert_called_once()
        for reopened_driver in drivers[1:]:
            reopened_driver.get.assert_called_once_with(browser.last_url)
        self.assertIs(browser.raw, drivers[-1])

    def test_stop_during_cooldown_does_not_reopen(self):
        browser, stop, checkpoint, drivers = self.make_browser()
        first, second = drivers[:2]
        stop.wait.return_value = True
        self.assertFalse(browser.restart_after_captcha(Mock()))
        checkpoint.assert_called_once()
        first.quit.assert_called_once()
        stop.wait.assert_called_once_with(60)
        second.get.assert_not_called()
        self.assertIs(browser.raw, first)

    def test_wait_loop_requests_manual_verification_on_fifth_captcha(self):
        browser, _, checkpoint, _ = self.make_browser()
        messages = []
        visible = [True, True, True, True, True, True, False]
        with (
            patch.object(app, "captcha_is_visible", side_effect=visible),
            patch.object(app, "maybe_accept_google_consent"),
            patch.object(app.random, "uniform", return_value=5),
            patch.object(app.time, "sleep"),
        ):
            self.assertTrue(app.wait_for_manual_captcha(browser, messages.append))
        self.assertEqual(checkpoint.call_count, app.CAPTCHA_AUTO_RESTART_LIMIT)
        self.assertTrue(any("lần thứ 5" in message for message in messages))
        self.assertTrue(messages[-1].startswith("Đã xác minh CAPTCHA"))

    def test_navigation_delay_can_be_cancelled(self):
        browser, stop, _, drivers = self.make_browser()
        first = drivers[0]
        stop.wait.return_value = True
        with patch.object(app.random, "uniform", return_value=7.5) as delay:
            with self.assertRaises(app.WebDriverException):
                browser.get("https://www.google.com/search?q=example")
        delay.assert_called_once_with(3, 8)
        stop.wait.assert_called_once_with(7.5)
        first.get.assert_not_called()

    def test_cleanup_blank_page_has_no_delay(self):
        browser, stop, _, drivers = self.make_browser()
        first = drivers[0]
        browser.get("about:blank")
        stop.wait.assert_not_called()
        first.get.assert_called_once_with("about:blank")
