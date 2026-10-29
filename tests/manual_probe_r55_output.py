#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""诊断脚本：检查“可单击图标 查看输出”链接的定位与点击后的页面跳转。

只做观测，不修改业务流程，也不保存下载文件。手动执行，不参与 unittest 收集：
    python tests/manual_probe_r55_output.py
    python tests/manual_probe_r55_output.py --wait 15
    python tests/manual_probe_r55_output.py --no-click        # 只探测，不点击
    python tests/manual_probe_r55_output.py --index 1         # 点第 2 个可见候选

关注点：
    1. 过滤前后各 frame 中同名链接的匹配数、可见性、坐标、所属行内容；
    2. 被点击的候选到底是哪一行，与手动看到的行是否一致；
    3. 点击后是新开页面、隐藏 frame 加载页面，还是直接触发下载。
"""
import argparse
from pathlib import Path
import re
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from config import Config
from core.context import TaskContext
from log.loggingConfig import logger
from web.browser import BrowserManager
from web.JDE.jde_function import (
    MENU_IFRAME,
    _click_checkbox,
    _click_menu_link,
    _enter_total_ledger,
)
from workflows.jde_flow.login_task import JDELoginTask


LABEL = "可单击图标 查看输出"
JOB_NAME = "R55GL1005_TKHXLS001"


def _frames_named(page):
    for frame in page.frames:
        name = "主文档" if frame == page.main_frame else (
            f"iframe[{frame.name or 'anonymous'}]"
        )
        yield name, frame


def _frame_url(frame):
    try:
        return frame.url or "(空)"
    except Exception as exc:
        return f"(取不到: {exc})"


def probe_output_links(page):
    """返回所有匹配链接的观测信息，包含隐藏候选。"""
    candidates = []
    for scope_name, frame in _frames_named(page):
        try:
            links = frame.get_by_role("link", name=LABEL)
            count = links.count()
        except Exception as exc:
            logger.warning(f"探测 {scope_name} 失败：{exc}")
            continue

        logger.info(
            f"探测 {scope_name}（{_frame_url(frame)[:90]}）：匹配数={count}"
        )
        for index in range(count):
            link = links.nth(index)
            info = {"scope": scope_name, "frame": frame, "index": index,
                    "link": link}
            try:
                info["visible"] = link.is_visible()
            except Exception:
                info["visible"] = None
            try:
                info["enabled"] = link.is_enabled()
            except Exception:
                info["enabled"] = None
            try:
                info["box"] = link.bounding_box()
            except Exception:
                info["box"] = None
            try:
                info["href"] = link.get_attribute("href")
            except Exception:
                info["href"] = None
            try:
                info["row"] = " ".join(
                    link.locator("xpath=ancestor::tr[1]")
                    .inner_text(timeout=1000)
                    .split()
                )
            except Exception:
                info["row"] = ""
            candidates.append(info)

            box = info["box"] or {}
            x = box.get("x")
            y = box.get("y")
            logger.info(
                f"  候选#{index} visible={info['visible']} "
                f"enabled={info['enabled']} "
                f"x={'%.0f' % x if x is not None else '-'} "
                f"y={'%.0f' % y if y is not None else '-'} "
                f"href={info['href']!r} 行内容={info['row'][:120]!r}"
            )
    return candidates


def click_and_watch(page, candidate, timeout, wait_seconds):
    """点击候选并记录下载、新开页面、frame URL 变化。"""
    context = page.context
    downloads = []
    new_pages = []
    frame_urls_before = {frame: _frame_url(frame) for frame in page.frames}
    page_url_before = page.url

    def on_download(download):
        try:
            owner = download.page.url
        except Exception:
            owner = "(未知)"
        downloads.append({
            "filename": download.suggested_filename,
            "url": download.url,
            "owner": owner,
        })

    def on_page(opened_page):
        new_pages.append(opened_page)
        try:
            opened_page.on("download", on_download)
        except Exception as exc:
            logger.warning(f"为新页面注册下载监听失败：{exc}")

    for existing in context.pages:
        existing.on("download", on_download)
    context.on("page", on_page)

    logger.info(
        f"点击 {candidate['scope']} 的候选#{candidate['index']}"
        f"（行内容={candidate['row'][:80]!r}）"
    )
    try:
        candidate["link"].click(timeout=timeout)
    finally:
        context.remove_listener("page", on_page)

    deadline = time.monotonic() + wait_seconds
    while time.monotonic() < deadline and not downloads:
        page.wait_for_timeout(500)

    logger.info("=== 点击结果 ===")
    if downloads:
        for item in downloads:
            logger.info(
                f"触发下载：{item['filename']} 来源页面={item['owner'][:90]} "
                f"url={item['url'][:140]}"
            )
    else:
        logger.warning(
            f"点击后 {wait_seconds} 秒内没有下载事件；"
            "说明进入的是中介页而不是文件下载"
        )

    logger.info(f"新开页面数={len(new_pages)}")
    for opened_page in new_pages:
        try:
            logger.info(f"  新页面 URL={opened_page.url[:200]}")
            logger.info(
                f"  新页面标题={opened_page.title()[:80]!r}"
            )
            media = opened_page.get_by_role("link", name="click here")
            logger.info(f"  新页面中 'click here' 链接数={media.count()}")
        except Exception as exc:
            logger.warning(f"  读取新页面失败：{exc}")

    logger.info(f"原页面 URL={page_url_before[:140]}")
    now_url = page.url
    if now_url != page_url_before:
        logger.warning(f"原页面 URL 已变化：{now_url[:140]}")

    for frame, before in frame_urls_before.items():
        after = _frame_url(frame)
        if after != before:
            logger.info(
                f"  frame 变化：{frame.name or 'anonymous'} "
                f"{before[:70]} -> {after[:70]}"
            )

    return downloads, new_pages


def prepare_submission(page, company, timeout):
    """复现 download_r55gl1005 的提交流程，停在“已提交作业”界面。"""
    home = page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()

    frame = page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    _click_menu_link(page, "中方科目总账", timeout)

    version = frame.get_by_role("textbox", name="版本", exact=True)
    version.wait_for(state="visible", timeout=timeout)
    version.fill("TKHXLS001", timeout=timeout)
    version.press("Enter", timeout=timeout)

    _click_checkbox(frame, "选择所有提取的记录", timeout)
    frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click(timeout=timeout)
    _click_checkbox(frame, "数据选择", timeout)
    frame.locator("a").filter(has_text="提交(U)").click(timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    selection = frame.locator('input[name="E1"]')
    selection.wait_for(state="visible", timeout=timeout)
    selection.click(timeout=timeout)
    selection.click(timeout=timeout)
    frame.locator("td").filter(has_text=re.compile(r"^Company$")).click(
        timeout=timeout
    )
    selection.fill("", timeout=timeout)
    selection.fill(str(company).strip().lstrip("0") or "0", timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    frame.locator("a").filter(has_text="屏幕(F)").click(timeout=timeout)
    frame.get_by_role("cell", name="已提交作业", exact=True).click(
        timeout=timeout
    )
    return frame


def main():
    parser = argparse.ArgumentParser(
        description="诊断“可单击图标 查看输出”链接定位与点击结果"
    )
    parser.add_argument("--credential-name", help="Windows 通用凭据名；默认读配置")
    parser.add_argument("--company", help="公司代码；默认取 房产税配置.明细账自公司")
    parser.add_argument("--timeout", type=int, default=60000,
                        help="页面操作超时毫秒数，默认 60000")
    parser.add_argument("--wait", type=int, default=10,
                        help="点击后观察秒数，默认 10")
    parser.add_argument("--index", type=int, default=0,
                        help="点击第几个可见候选（从 0 开始），默认 0")
    parser.add_argument("--no-click", action="store_true",
                        help="只探测链接，不执行点击")
    parser.add_argument("--headless", action="store_true", help="无界面运行")
    parser.add_argument("--keep-open", action="store_true",
                        help="诊断结束后保留浏览器，按 Enter 再关闭")
    args = parser.parse_args()

    config = Config()
    config.load(
        str(ROOT / "data" / "UserConfig.xlsx"),
        str(ROOT / "data" / "Config.xlsx"),
    )
    company = args.company or config.get("房产税配置.明细账自公司")
    if not company or not str(company).strip():
        raise SystemExit("缺少公司代码：请传 --company 或配置 房产税配置.明细账自公司")
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
            raise SystemExit("JDE 登录失败，无法执行诊断")

        page = browser.get_page()
        logger.info(f"公司={company}，开始进入 R55GL1005 提交界面")
        frame = prepare_submission(page, company, args.timeout)

        logger.info("=== 过滤前（已提交作业列表原样） ===")
        before = probe_output_links(page)

        logger.info(f"=== 输入“作业”={JOB_NAME} 并回车过滤 ===")
        job_filter = frame.get_by_role("textbox", name="作业", exact=True)
        job_filter.wait_for(state="visible", timeout=args.timeout)
        job_filter.fill(JOB_NAME, timeout=args.timeout)
        job_filter.press("Enter", timeout=args.timeout)

        logger.info("=== 过滤后（立即探测，未额外等待） ===")
        after = probe_output_links(page)

        visible = [item for item in after if item["visible"]]
        logger.info(
            f"过滤前匹配总数={len(before)}，过滤后匹配总数={len(after)}，"
            f"其中可见={len(visible)}"
        )
        if len(visible) > 1:
            logger.warning(
                f"存在 {len(visible)} 个可见同名链接；"
                "程序取 .first 时可能点到与手动不同的行"
            )

        if args.no_click:
            logger.info("--no-click：跳过点击")
        elif not visible:
            logger.warning("没有可见候选，无法点击；请检查是否已进入“已提交作业”列表")
        else:
            if args.index >= len(visible):
                raise SystemExit(
                    f"--index {args.index} 超出可见候选数 {len(visible)}"
                )
            click_and_watch(page, visible[args.index], args.timeout, args.wait)

        if args.keep_open:
            input("诊断完成，按 Enter 关闭浏览器...")
    finally:
        context.close_all_browsers()


if __name__ == "__main__":
    main()
