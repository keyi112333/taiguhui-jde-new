"""通过框架配置和浏览器管理器运行 JDE 模块。"""

from config.config_loader import ConfigError
from core.context import TaskContext
from core.task_base import TaskBase
from log.loggingConfig import logger
from web.browser import BrowserManager
from workflows.jde_flow.login_task import JDELoginTask
from workflows.jde_flow.property_tax import PropertyTaxFlow


class JDEMainTask(TaskBase):
    def __init__(self, config):
        context = TaskContext()
        context.config = config
        super().__init__("JDE业务流程", context)

    def execute(self) -> bool:
        config = self.context.config
        if not config.get("JDE配置.jde_url", ""):
            raise ConfigError("请配置 JDE配置.jde_url")
        if not config.get("JDE配置.jde_credential_name", ""):
            raise ConfigError("请配置 JDE配置.jde_credential_name")

        headless = config.get("JDE配置.headless", False)
        if isinstance(headless, str):
            if headless.strip().lower() not in ("true", "false", "1", "0", "yes", "no"):
                raise ConfigError("JDE配置.headless 必须为 true 或 false")
            headless = headless.strip().lower() in ("true", "1", "yes")
        browser = BrowserManager(headless=bool(headless))
        self.context.set_browser("JDE", browser)
        try:
            browser.start()
            login_task = JDELoginTask(self.context)
            if not login_task.run():
                raise RuntimeError("JDE login task failed")
            PropertyTaxFlow(self.context).run()
            return True
        except Exception:
            page = browser.get_page()
            if page is not None:
                try:
                    page.screenshot(path=str(config.get_path("screenshot", "jde_failure.png")))
                except Exception:
                    logger.warning("JDE错误截图保存失败")
            raise
        finally:
            self.context.close_all_browsers()
