"""无需访问 JDE 的流程验证。"""
import os
from pathlib import Path
import sys
import unittest
from unittest.mock import MagicMock, patch

ROOT =Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))

from workflows.jde_flow.main_task import JDEMainTask


@patch("workflows.jde_flow.login_task.read_windows_credential", return_value=("test", "test"))
class JDEFlowTests(unittest.TestCase):
    def setUp(self):
        self.env = patch.dict(os.environ, {}, clear=True)
        self.env.start()
        self.addCleanup(self.env.stop)
        self.config = MagicMock()
        values = {"JDE配置.jde_" + key: value for key, value in
                  (("url", "https://example.invalid"), ("credential_name", "JDE_TEST"))}
        self.config.get.side_effect = lambda key, default=None: values.get(key, default)

    @patch("workflows.jde_flow.login_task.jde_login")
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_login_and_cleanup(self, manager, login, credential):
        task = JDEMainTask(self.config)
        self.assertTrue(task.run())
        login.assert_called_once()
        manager.return_value.close.assert_called_once()
        self.assertTrue(task.context.get("jde_logged_in"))

    @patch("workflows.jde_flow.login_task.jde_login", side_effect=RuntimeError("login failed"))
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_failure_screenshot_and_cleanup(self, manager, login, credential):
        task = JDEMainTask(self.config)
        self.assertFalse(task.run())
        self.assertEqual(task.status, "failed")
        manager.return_value.get_page.return_value.screenshot.assert_called_once()
        manager.return_value.close.assert_called_once()

    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_missing_credentials_does_not_start_browser(self, manager, credential):
        self.config.get.side_effect = lambda key, default=None: default
        self.assertFalse(JDEMainTask(self.config).run())
        manager.assert_not_called()

    @patch("workflows.jde_flow.login_task.jde_login")
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_reads_named_windows_credential(self, manager, login, credential):
        self.assertTrue(JDEMainTask(self.config).run())
        credential.assert_called_once_with("JDE_TEST")
        self.assertEqual(login.call_args.kwargs["jde_user"], "test")
        self.assertEqual(login.call_args.kwargs["jde_pass"], "test")


if __name__ == "__main__":
    unittest.main()
