#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""登录 JDE，定位并点击“按科目帐的帐户查询”（真实环境手动探针）。

运行：python tests/manual_probe_account_query.py
"""
import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from config import Config
from core.context import TaskContext
from log.loggingConfig import logger
from web.browser import BrowserManager
from web.JDE.jde_function import MENU_IFRAME, _enter_total_ledger
from workflows.jde_flow.login_task import JDELoginTask


LABEL = "按科目帐的帐户查询"


def probe_and_click(page, timeout):
    scopes = [("主文档", page)]
    scopes.extend(
        (f"iframe[{frame.name or 'anonymous'}]", frame)
        for frame in page.frames
        if frame != page.main_frame
    )

    candidates = []
    for scope_name, scope in scopes:
        links = scope.get_by_role("link", name=LABEL)
        count = links.count()
        visible_indexes = [
            index for index in range(count) if links.nth(index).is_visible()
        ]
        details = []
        for index in range(count):
            link = links.nth(index)
            details.append(
                f"#{index + 1}(visible={link.is_visible()}, "
                f"enabled={link.is_enabled()}, href={link.get_attribute('href')!r})"
            )
        logger.info(
            f"探测 {scope_name}：匹配数={count}，"
            f"可见数={len(visible_indexes)}"
            + (f"，详情={'；'.join(details)}" if details else "")
        )
        candidates.extend(
            (scope_name, links.nth(index)) for index in visible_indexes
        )

    if not candidates:
        raise RuntimeError(f"登录并进入总账页后，未发现可见选项：{LABEL}")

    scope_name, target = candidates[0]
    logger.info(f"探针准备点击：{scope_name}中的{LABEL}")
    target.click(timeout=timeout)
    logger.info(f"探针点击完成：{scope_name}中的{LABEL}")

    page.wait_for_timeout(1500)
    frame = page.frame_locator(MENU_IFRAME)
    textbox_count = frame.get_by_role("textbox").count()
    logger.info(
        f"点击后页面状态：URL={page.url}，"
        f"菜单 iframe 文本框数量={textbox_count}"
    )


def main():
    parser = argparse.ArgumentParser(
        description="登录 JDE 并探测/点击“按科目帐的帐户查询”"
    )
    parser.add_argument(
        "--credential-name",
        help="Windows 通用凭据名；默认读取 Config.xlsx",
    )
    parser.add_argument(
        "--headless", action="store_true", help="无界面运行（默认显示浏览器）"
    )
    parser.add_argument(
        "--timeout", type=int, default=20000, help="页面操作超时毫秒数，默认 20000"
    )
    args = parser.parse_args()

    config = Config()
    config.load(
        str(ROOT / "data" / "UserConfig.xlsx"),
        str(ROOT / "data" / "Config.xlsx"),
    )
    credential_name = args.credential_name or config.get(
        "JDE配置.jde_credential_name"
    )
    if not credential_name:
        raise SystemExit(
            "缺少 JDE配置.jde_credential_name：请填写 Config.xlsx "
            "或传入 --credential-name"
        )
    if args.timeout <= 0:
        raise SystemExit("--timeout 必须大于 0")
    if args.credential_name:
        config._loader._config.setdefault("JDE配置", {})[
            "jde_credential_name"
        ] = args.credential_name

    context = TaskContext()
    context.config = config
    browser = BrowserManager(headless=args.headless)
    context.set_browser("JDE", browser)

    try:
        browser.start()
        if not JDELoginTask(context).run():
            raise RuntimeError("JDE 登录失败，无法执行菜单探针")
        logger.info("登录成功，准备进入总账页签")
        page = browser.get_page()
        _enter_total_ledger(page.frame_locator(MENU_IFRAME), args.timeout)
        logger.info("已进入总账页签，开始探测菜单选项")
        probe_and_click(page, args.timeout)
        logger.info("菜单选项探针完成")
    finally:
        context.close_all_browsers()


if __name__ == "__main__":
    main()
