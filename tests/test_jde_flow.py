"""无需访问 JDE 的流程验证。"""
import os
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from workflows.jde_flow.main_task import JDEMainTask, parse_shared_path


RUN_ROW = {"window凭据": "JDE_TEST", "凭据账号": "test",
           "共享盘路径": r"C:\共享\太古汇 2501"}


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
        self.config.get_rows.side_effect = lambda name: [dict(RUN_ROW)] if name == "登录配置" else []
        prepare = patch.object(JDEMainTask, "prepare", return_value={})
        prepare.start()
        self.addCleanup(prepare.stop)

    @patch("workflows.jde_flow.main_task.PropertyTaxFlow")
    @patch("workflows.jde_flow.login_task.jde_login")
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_login_and_cleanup(self, manager, login, flow, credential):
        task = JDEMainTask(self.config)
        self.assertTrue(task.run())
        login.assert_called_once()
        flow.return_value.run.assert_called_once()
        manager.return_value.close.assert_called_once()
        self.assertTrue(task.context.get("jde_logged_in"))

    @patch("workflows.jde_flow.main_task.PropertyTaxFlow")
    @patch("workflows.jde_flow.login_task.jde_login", side_effect=RuntimeError("login failed"))
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_failure_screenshot_and_cleanup(self, manager, login, flow, credential):
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

    @patch("workflows.jde_flow.main_task.PropertyTaxFlow")
    @patch("workflows.jde_flow.login_task.jde_login")
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_reads_named_windows_credential(self, manager, login, flow, credential):
        self.assertTrue(JDEMainTask(self.config).run())
        credential.assert_called_once_with("JDE_TEST")
        self.assertEqual(login.call_args.kwargs["jde_user"], "test")
        self.assertEqual(login.call_args.kwargs["jde_pass"], "test")

    @patch("workflows.jde_flow.main_task.JDEMainTask.prepare")
    @patch("workflows.jde_flow.main_task.BrowserManager")
    def test_empty_login_sheet_blocks_run(self, manager, prepare, credential):
        self.config.get_rows.side_effect = lambda name: []
        task = JDEMainTask(self.config)
        self.assertFalse(task.run())
        manager.assert_not_called()
        prepare.assert_not_called()


class FakeConfig:
    """prepare 需要的最小配置接口。"""

    def __init__(self, values, output_dir):
        self.values = dict(values)
        self.output_dir = output_dir

    def get(self, key, default=None):
        return self.values.get(key, default)

    def set(self, key, value):
        self.values[key] = value

    def get_rows(self, name):
        return []

    def get_path(self, kind, filename=""):
        return self.output_dir


class PrepareTests(unittest.TestCase):
    """登录配置一行 -> 公司代码、模板、数据源目录。"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.shared = self.root / "太古汇 2501"
        self.shared.mkdir(parents=True)
        (self.shared / "房产税计税表 太古汇.xlsx").write_bytes(b"fake")
        self.output = self.root / "03_输出"
        self.output.mkdir()

    def build(self, values=None):
        base = {"JDE配置.jde_url": "https://example.invalid",
                "房产税配置.会计期间": "202609"}
        base.update(values or {})
        self.config = FakeConfig(base, self.output)
        return JDEMainTask(self.config)

    @patch("workflows.jde_flow.main_task.read_windows_credential",
           return_value=("test", "secret"))
    def test_prepare_derives_company_and_paths(self, credential):
        task = self.build()
        info = task.prepare({"window凭据": "JDE_TEST", "凭据账号": "test",
                             "共享盘路径": str(self.shared)})
        self.assertEqual(info["short_name"], "太古汇")
        self.assertEqual(info["company"], "02501")
        for key in ("明细账自公司", "明细账至公司", "分录公司"):
            self.assertEqual(self.config.get(f"房产税配置.{key}"), "02501")
        self.assertEqual(self.config.get("房产税配置.公司"), "太古汇")
        self.assertEqual(Path(self.config.get("房产税配置.模板文件")),
                         self.shared / "房产税计税表 太古汇.xlsx")
        self.assertEqual(Path(self.config.get("房产税配置.输出目录")), self.shared)
        self.assertEqual(info["data_dir"].name, "太古汇 2501 202609")
        self.assertTrue(info["data_dir"].is_dir())
        self.assertEqual(self.config.get("JDE配置.jde_credential_name"), "JDE_TEST")

    @patch("workflows.jde_flow.main_task.read_windows_credential",
           return_value=("test", "secret"))
    def test_prepare_defaults_to_natural_month(self, credential):
        task = self.build({"房产税配置.会计期间": None})
        from datetime import date
        info = task.prepare({"window凭据": "JDE_TEST", "凭据账号": "test",
                             "共享盘路径": str(self.shared)})
        self.assertEqual(info["data_dir"].name,
                         f"太古汇 2501 {date.today():%Y%m}")

    def test_prepare_rejects_missing_columns_and_template(self):
        task = self.build()
        with self.assertRaises(Exception):
            task.prepare({"window凭据": "JDE_TEST", "凭据账号": "test", "共享盘路径": ""})

        (self.shared / "房产税计税表 太古汇.xlsx").unlink()
        with self.assertRaises(Exception):
            task.prepare({"window凭据": "JDE_TEST", "凭据账号": "test",
                          "共享盘路径": str(self.shared)})

    def test_parse_shared_path(self):
        path, short, code = parse_shared_path(r"N:\Finance\房产税\太古汇 2501")
        self.assertEqual((short, code), ("太古汇", "2501"))
        for bad in (r"N:\房产税\太古汇", r"N:\房产税\太古汇 2501 01", r"N:\房产税\太古汇 abc"):
            with self.assertRaises(Exception):
                parse_shared_path(bad)


if __name__ == "__main__":
    unittest.main()
