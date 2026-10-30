"""按「登录配置」逐行运行：每行一个账号+共享盘目录，完整跑一遍流程。"""

from datetime import date
from pathlib import Path

from config.config_loader import ConfigError
from core.context import TaskContext
from core.task_base import TaskBase
from log.loggingConfig import logger
from utils.credentials import read_windows_credential
from utils.helpers import get_timestamp
from web.browser import BrowserManager
from workflows.jde_flow.login_task import JDELoginTask
from workflows.jde_flow.property_tax import PropertyTaxFlow, month_bounds


RUN_SHEET = "登录配置"
REQUIRED_COLUMNS = ("window凭据", "凭据账号", "共享盘路径")


def parse_shared_path(value):
    """从共享盘路径末段目录（如“太古汇 2501”）解析公司简称与公司代码。"""
    path = Path(str(value).strip())
    tokens = path.name.split()
    if len(tokens) != 2 or not tokens[1].isdigit():
        raise ConfigError(
            f"共享盘路径末段需为“公司简称 公司代码”（如 太古汇 2501），当前：{path.name}"
        )
    return path, tokens[0], tokens[1]


def read_headless(value):
    if isinstance(value, str):
        text = value.strip().lower()
        if text not in ("true", "false", "1", "0", "yes", "no"):
            raise ConfigError("JDE配置.headless 必须为 true 或 false")
        return text in ("true", "1", "yes")
    return bool(value)


class JDEMainTask(TaskBase):
    def __init__(self, config):
        context = TaskContext()
        context.config = config
        super().__init__("JDE业务流程", context)

    def execute(self) -> bool:
        config = self.context.config
        if not config.get("JDE配置.jde_url", ""):
            raise ConfigError("请配置 JDE配置.jde_url")
        runs = config.get_rows(RUN_SHEET)
        if not runs:
            raise ConfigError(
                f"{RUN_SHEET} 页没有数据行，请填写 "
                f"{'、'.join(REQUIRED_COLUMNS)}；每行会完整跑一遍流程"
            )
        failures = []
        for index, run in enumerate(runs, 1):
            label = f"[{index}/{len(runs)}] {run.get('共享盘路径')}"
            logger.info(f"开始执行 {label}")
            try:
                self.run_one(run)
                logger.info(f"{label} 执行完成")
            except Exception as exc:
                logger.exception(f"{label} 执行失败：{exc}")
                failures.append(f"{label}：{exc}")
        if failures:
            raise ConfigError("以下账号执行失败：" + "；".join(failures))
        return True

    def prepare(self, run):
        """校验本行配置并把公司代码、路径写入房产税配置，返回本次运行信息。"""
        config = self.context.config
        values = {name: str(run.get(name) or "").strip() for name in REQUIRED_COLUMNS}
        missing = [name for name, value in values.items() if not value]
        if missing:
            raise ConfigError(f"{RUN_SHEET} 缺少：{', '.join(missing)}")

        shared_path, short_name, code = parse_shared_path(values["共享盘路径"])
        if not shared_path.is_dir():
            raise ConfigError(f"共享盘路径不存在：{shared_path}")
        template = shared_path / f"房产税计税表 {short_name}.xlsx"
        if not template.is_file():
            raise ConfigError(f"共享盘缺少模板：{template}")

        username, _ = read_windows_credential(values["window凭据"])
        if username.split("\\")[-1].strip().lower() != values["凭据账号"].lower():
            logger.warning(
                f"Windows 凭据 {values['window凭据']} 的用户名 {username} "
                f"与「凭据账号」{values['凭据账号']} 不一致，请确认数据归属"
            )

        period_value = config.get("房产税配置.会计期间")
        first_day, _ = month_bounds(period_value or date.today())
        period = f"{first_day:%Y%m}"
        company = code.zfill(5)
        stamp = get_timestamp()
        data_dir = config.get_path("output") / f"{short_name} {code} {period}"
        data_dir.mkdir(parents=True, exist_ok=True)

        overrides = {
            "JDE配置.jde_credential_name": values["window凭据"],
            "房产税配置.公司": short_name,
            "房产税配置.模板文件": str(template),
            "房产税配置.输出目录": str(shared_path),
            "房产税配置.明细账自公司": company,
            "房产税配置.明细账至公司": company,
            "房产税配置.分录公司": company,
            # 明细账不写后缀：下载时按 JDE 实际文件名补 .xlsx。
            "房产税配置.明细账文件": str(data_dir / f"JDE明细账 {stamp}"),
            "房产税配置.R55GL1005文件": str(data_dir / f"R55GL1005_{stamp}.xlsx"),
        }
        for key, value in overrides.items():
            config.set(key, value)
        logger.info(
            f"运行参数：凭据={values['window凭据']} 账号={values['凭据账号']} "
            f"公司={company} 共享盘={shared_path} 数据源={data_dir}"
        )
        return {
            "credential": values["window凭据"],
            "company": company,
            "short_name": short_name,
            "shared_path": shared_path,
            "template": template,
            "data_dir": data_dir,
        }

    def run_one(self, run):
        info = self.prepare(run)
        config = self.context.config
        browser = BrowserManager(headless=read_headless(config.get("JDE配置.headless", False)))
        self.context.set_browser("JDE", browser)
        try:
            browser.start()
            if not JDELoginTask(self.context).run():
                raise RuntimeError("JDE 登录失败")
            PropertyTaxFlow(self.context).run()
        except Exception:
            page = browser.get_page()
            if page is not None:
                suffix = "_".join(filter(None, (info.get("short_name"), info.get("company"))))
                try:
                    page.screenshot(
                        path=str(config.get_path(
                            "screenshot", f"jde_failure_{suffix}.png" if suffix else "jde_failure.png"
                        ))
                    )
                except Exception:
                    logger.warning("JDE错误截图保存失败")
            raise
        finally:
            self.context.close_all_browsers()
