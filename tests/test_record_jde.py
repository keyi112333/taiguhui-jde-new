import unittest
from unittest.mock import patch

from record_jde import record


class RecordingTests(unittest.TestCase):
    @patch("record_jde.JDELoginTask")
    @patch("record_jde.BrowserManager")
    @patch("record_jde.Config")
    def test_record_only_after_login(self, config, manager, login):
        events = []
        login.return_value.run.side_effect = lambda: events.append("login") or True
        manager.return_value.get_page.return_value.pause.side_effect = lambda: events.append("record")
        record("JDE")
        self.assertEqual(events, ["login", "record"])
        manager.assert_called_once_with(headless=False)
        manager.return_value.close.assert_called_once()

    @patch("record_jde.JDELoginTask")
    @patch("record_jde.BrowserManager")
    @patch("record_jde.Config")
    def test_failed_login_does_not_record(self, config, manager, login):
        login.return_value.run.return_value = False
        with self.assertRaises(RuntimeError):
            record("JDE")
        manager.return_value.get_page.return_value.pause.assert_not_called()
        manager.return_value.close.assert_called_once()
