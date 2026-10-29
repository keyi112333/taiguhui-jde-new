"""读取 Windows 凭据登录 JDE，随后通过 Inspector 录制后续操作。"""
import argparse
from pathlib import Path

from config import Config
from core.context import TaskContext
from log.loggingConfig import logger
from web.browser import BrowserManager
from workflows.jde_flow.login_task import JDELoginTask


def record(credential_name=None):
    root = Path(__file__).resolve().parent
    config = Config()
    config.load(str(root / "data/UserConfig.xlsx"), str(root / "data/Config.xlsx"))
    if credential_name:
        # 仅本次录制覆盖凭据名，不修改配置文件。
        config._loader._config.setdefault("JDE配置", {})["jde_credential_name"] = credential_name
    if not config.get("JDE配置.jde_credential_name"):
        raise ValueError("请填写 Config.xlsx 中的 jde_credential_name，或传入 --credential-name")

    context = TaskContext()
    context.config = config
    # Inspector 录制需要有界面浏览器，不使用业务配置中的 headless。
    browser = BrowserManager(headless=False)
    context.set_browser("JDE", browser)
    try:
        browser.start()
        if not JDELoginTask(context).run():
            raise RuntimeError("JDE自动登录失败，未启动录制")
        logger.info("登录成功。在 Inspector 中点击 Record 录制后续操作，并复制生成的 Python 代码。")
        logger.info("录制完成并复制代码后，点击 Resume 结束本次会话。")
        browser.get_page().pause()
    finally:
        context.close_all_browsers()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="凭据自动登录 JDE 后打开 Playwright 录制器")
    parser.add_argument("--credential-name", help="Windows 通用凭据名称；默认读取 Config.xlsx")
    args = parser.parse_args()
    record(args.credential_name)
