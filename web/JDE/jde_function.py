
import functools
import re
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from utils.helpers import ensure_dir,get_timestamp

MENU_IFRAME = 'iframe[name="e1menuAppIframe"]'

OUTPUT_LINK_NAME = "可单击图标 查看输出"
JOB_FILTER_NAME = "作业"
JOB_NAME = "R55GL1005_TKHXLS001"
JOB_DONE_KEYWORDS = ("完成", "Done")
JOB_FAILED_KEYWORDS = ("错误", "失败", "Error")
JOB_STATUS_POLL_SECONDS = 120
JOB_STATUS_POLL_INTERVAL_SECONDS = 5


# 重试
def retry_on_failure(max_retries=3,delay=5):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt < max_retries-1:
                        logger.warning(f"{func.__name__} 执行失败，准备重试")
                        logger.info(f"等待 {delay} 秒后进行第 {attempt+2} 次尝试...")
                        time.sleep(delay)
                    else:
                        logger.info(f"{func.__name__} 最终执行失败，已重试 {max_retries} 次")
                        raise
            return None
        return wrapper
    return decorator


# 登录JDE
@retry_on_failure(max_retries=3,delay=5)
def jde_login(in_page ,jde_url ,jde_environment ,jde_role ,jde_user ,jde_pass):

    in_page.goto(jde_url)

    element_exists = in_page.locator("#drop_mainmenu")
    if element_exists.is_visible(timeout=5000):
        logger.info("登录成功")
    else:
        # Details
        element_exists = in_page.get_by_role("textbox", name="Environment")
        if not element_exists.is_visible(timeout=2000):

            in_page.get_by_role("link", name="Details").click()
            in_page.wait_for_timeout(500)

        if jde_environment:
            in_page.get_by_role("textbox", name="Environment").fill(jde_environment)
            in_page.wait_for_timeout(500)
        if jde_role:
            in_page.get_by_role("textbox", name="Role").fill(jde_role)
            in_page.wait_for_timeout(500)

        # User Password
        in_page.get_by_role("textbox", name="User ID").fill(jde_user)
        in_page.wait_for_timeout(500)

        in_page.get_by_role("textbox", name="Password").fill(jde_pass)
        in_page.wait_for_timeout(500)

        # Sign In
        in_page.get_by_role("button", name="Sign In").click()
        in_page.wait_for_timeout(3000)

        element_exists = in_page.locator("#drop_mainmenu")
        element_exists.wait_for(state="visible", timeout=20000)
        if element_exists.is_visible():
            logger.info("登录成功")
        else:
            raise Exception("登录失败，未进入JDE界面，请检查！")


def _enter_total_ledger(frame,timeout):
    """先进入所有页面，再进入总账页签（与录制脚本一致：页签在 iframe 内点击）"""
    select_page = frame.get_by_role("button", name="选择页面")
    select_page.wait_for(state="visible",timeout=timeout)
    select_page.click()
    all_pages_grid = frame.get_by_role("gridcell", name="所有页面")
    all_pages_grid.wait_for(state="visible",timeout=timeout)
    all_pages_grid.click()
    
    all_pages =frame.get_by_text("所有页面").first
    all_pages.click()
    if all_pages.is_visible():
        select_page.click()


    tab =frame.locator("#tabBar").get_by_text("总帐")
    tab.wait_for(state="visible",timeout=timeout)
    tab.click()


def _click_menu_link(in_page, name, timeout):
    """在主文档和当前 iframe 中查找可见菜单项并点击。"""
    deadline = time.monotonic() + timeout / 1000
    while True:
        for frame in in_page.frames:
            links = frame.get_by_role("link", name=name)
            for index in range(links.count()):
                link = links.nth(index)
                if link.is_visible():
                    logger.info(f"在页面/frame中找到菜单项并点击：{name}")
                    link.click(timeout=timeout)
                    return

        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise PlaywrightTimeoutError(
                f"等待 {timeout} 毫秒后，主文档和 iframe 中均未找到可见菜单项：{name}"
            )
        in_page.wait_for_timeout(min(500, max(1, int(remaining * 1000))))


def _click_account_query_link(in_page, timeout):
    _click_menu_link(in_page, "按科目帐的帐户查询", timeout)


def _click_checkbox(frame, name, timeout):
    """JDE 网格复选框由页面脚本维护状态，check() 的状态校验会误报，只做点击。"""
    checkbox = frame.get_by_role("checkbox", name=name).first
    checkbox.wait_for(state="visible", timeout=timeout)
    checkbox.click(timeout=timeout)
    logger.info(f"已点击勾选框：{name}")


def download_gl_movement(in_page,ledger_type,from_company,to_company,
                         start_date,end_date,destination,timeout=60000,
                         grid_timeout=300000,download_timeout=300000):

    for label,value in (("分类账类型",ledger_type),("自公司",from_company),("至公司",to_company),
                         ("起始日期",start_date),("至日期",end_date)):
        if value is None or not str(value).strip():
            raise ValueError(f"明细账缺少导出参数：{label}")


    frame =in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame,timeout)
    logger.info("打开按科目帐的帐户查询")
    _click_account_query_link(in_page, timeout)

    def textbox(name,exact=False):
        locator = frame.get_by_role("textbox",name = name, exact=exact)
        locator.wait_for(state="visible",timeout = timeout)
        return locator

    logger.info(f"填写查询条件：{from_company}-{to_company}{start_date}-{end_date}类型 {ledger_type}")
    textbox("分类帐类型 2", exact=True).fill(str(ledger_type).strip())
    textbox("自公司").fill(str(from_company).strip())
    textbox("至公司").fill(str(to_company).strip())
    textbox("起始日期").fill(str(start_date).strip())
    textbox("至日期").fill(str(end_date).strip())

    frame.get_by_role("button",name="查找 (Ctrl+Alt+I)").click()

    logger.info("查询已提交，跳转至提取记录末尾并等待网格加载")
    frame.get_by_role("button", name="转到结束").click(timeout=timeout)
    select_all = frame.get_by_role("checkbox", name="选择所有提取的记录")
    select_all.wait_for(state="visible", timeout=grid_timeout)
    select_all.check(timeout=grid_timeout)
    logger.info("网格记录已就绪并选择全部记录")
    frame.get_by_role(
        "button", name="输出网格数据 (Ctrl + Shift +E)"
    ).click(timeout=timeout)
    local_file = frame.get_by_role("radio", name="本地文件系统")
    local_file.wait_for(state="visible",timeout=timeout)
    local_file.check()
    frame.get_by_role("radio", name="输出到 Excel").check()
    frame.get_by_role("radio", name="输出到新工作簿 输出到新工作簿").check()
    frame.get_by_role("radio", name="所有").check()

    logger.info(f"已点击继续，等待明细账下载，最长 {download_timeout // 1000} 秒")
    with in_page.expect_download(timeout=download_timeout) as download_info:
        frame.get_by_role(
            "cell", name="继续 (Ctrl+Alt+E)", exact=True
        ).locator("a").click(timeout=download_timeout)
    download =download_info.value

    target = Path(destination)
    if not target.suffix:
        target=target.with_suffix(Path(download.suggested_filename).suffix or ".xlsx")
    ensure_dir(target.parent)
    download.save_as(str(target))
    logger.info(f"明细账已导出：{target}")
    return target


def _set_clipboard_text(text):
    """把整块 TSV 写入系统剪贴板，供 JDE 网格 Ctrl+V 粘贴。"""
    try:
        import win32clipboard
    except ImportError as exc:
        raise RuntimeError(f"缺少 pywin32，无法写入剪贴板：{exc}") from exc
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def create_journal_entry(in_page, company, description, gl_date, rows,
                         timeout=60000, discard=True):
    """录入日记帐分录并返回 JDE 生成的批号。

    rows 为 (帐号, 借项金额, 贷项金额, 备注) 序列，按 TSV 整块粘贴到第一个单元格，
    行数随数据动态变化。discard=True 时按录制脚本两次取消（开发环境不允许保存）。
    """
    if not rows:
        raise ValueError("分录没有可录入的数据行")
    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    logger.info("打开日记帐分录")
    _click_menu_link(in_page, "日记帐分录", timeout)

    frame.get_by_role("button", name="增加 (Ctrl+Alt+A)").click(timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    batch_box = frame.get_by_role("textbox", name="批号")
    batch_box.wait_for(state="visible", timeout=timeout)
    batch_box.click(timeout=timeout)
    batch = batch_box.input_value(timeout=timeout)
    logger.info(f"JDE 批号：{batch}")

    frame.get_by_role("textbox", name="Document Company").fill(
        str(company).strip(), timeout=timeout
    )
    frame.get_by_role("textbox", name="说明").fill(
        str(description).strip(), timeout=timeout
    )
    frame.get_by_role("textbox", name="总帐日期").fill(
        str(gl_date).strip(), timeout=timeout
    )

    lines = []
    for code, debit, credit, remark in rows:
        lines.append("\t".join("" if value is None else _cell_text(value)
                               for value in (code, debit, credit, remark)))
    _set_clipboard_text("\n".join(lines))
    logger.info(f"分录数据已写入剪贴板：{len(lines)} 行")
    frame.locator(".JSTextfield").first.click(timeout=timeout)
    in_page.wait_for_timeout(500)
    in_page.keyboard.press("Control+V")
    in_page.wait_for_timeout(1000)

    if discard:
        # 正式环境开发不允许保存，按录制脚本取消两次退出。
        for _ in range(2):
            frame.get_by_role("button", name="取消 (Ctrl+Alt+L)").click(timeout=timeout)
            in_page.wait_for_timeout(500)
        logger.info("已取消分录录入（未保存）")
    return batch


def _cell_text(value):
    """数值去掉多余小数位，避免粘贴时出现 1.146979552E7 之类的科学计数法。"""
    if isinstance(value, float):
        return format(value, "f").rstrip("0").rstrip(".")
    return str(value)


def _click_and_collect_download(in_page, link, timeout):
    """点击后会跳转页面触发下载；下载可能在原页面或新开页面中发起。"""
    context = in_page.context
    existing_pages = list(context.pages)
    collected = []

    def collect(download):
        collected.append(download)

    def watch(page):
        page.on("download", collect)

    for page in existing_pages:
        watch(page)
    context.on("page", watch)
    try:
        link.click(timeout=timeout)
        deadline = time.monotonic() + timeout / 1000
        while not collected:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise PlaywrightTimeoutError(
                    f"点击后 {timeout} 毫秒内未触发下载"
                )
            in_page.wait_for_timeout(min(500, max(1, int(remaining * 1000))))
        return collected[0]
    finally:
        context.remove_listener("page", watch)
        for page in context.pages:
            page.remove_listener("download", collect)


def _close_pages_opened_after(context, existing_pages):
    for page in context.pages:
        if page not in existing_pages:
            page.close()


def _row_text(link):
    try:
        return " ".join(
            link.locator("xpath=ancestor::tr[1]").inner_text(timeout=1000).split()
        )
    except Exception:
        return ""


def _job_number(row_text):
    match = re.search(r"R55GL1005\w*\s+(\d{5,10})\b", row_text)
    if match is None:
        match = re.search(r"\b(\d{5,10})\b", row_text)
    return match.group(1) if match else None


def _requery_submitted_jobs(frame, job_name, timeout=1500):
    """在“作业”框重新填入作业名并回车，让列表重新查询以更新状态。"""
    try:
        job_filter = frame.get_by_role("textbox", name=JOB_FILTER_NAME, exact=True)
        job_filter.fill(job_name, timeout=timeout)
        job_filter.press("Enter", timeout=timeout)
        return True
    except Exception as exc:
        logger.warning(f"重新查询作业状态失败：{exc}")
        return False


def _wait_latest_job_completed(frame, timeout,
                               poll_seconds=JOB_STATUS_POLL_SECONDS,
                               job_name=JOB_NAME):
    """等“已提交作业”最新一行（本次提交）变为完成，返回该行的输出链接。

    作业号每次提交都会变，因此按“列表第一行”确定目标，再用作业号跟踪该行；
    状态轮询总时长上限固定为 poll_seconds 秒，避免下载到未完成的作业。
    """
    links = frame.get_by_role("link", name=OUTPUT_LINK_NAME)
    deadline = time.monotonic() + timeout / 1000
    while links.count() == 0:
        if time.monotonic() >= deadline:
            raise PlaywrightTimeoutError(
                f"{timeout} 毫秒内“已提交作业”列表仍为空白"
            )
        time.sleep(0.5)

    first_row = _row_text(links.first)
    job_no = _job_number(first_row)
    if job_no is None:
        raise RuntimeError(f"无法从最新作业行解析作业号：{first_row!r}")
    logger.info(f"本次提交的作业号：{job_no}；当前状态：{first_row}")

    poll_deadline = time.monotonic() + poll_seconds
    last_row = first_row
    requery = False
    while True:
        if requery:
            _requery_submitted_jobs(frame, job_name)

        target = None
        for index in range(links.count()):
            link = links.nth(index)
            row = _row_text(link)
            if job_no in row:
                last_row = row
                target = link
                break

        if target is not None and any(k in last_row for k in JOB_DONE_KEYWORDS):
            logger.info(f"作业 {job_no} 已完成：{last_row}")
            return target

        if any(k in last_row for k in JOB_FAILED_KEYWORDS):
            raise RuntimeError(f"作业 {job_no} 状态异常：{last_row}")

        if time.monotonic() >= poll_deadline:
            raise PlaywrightTimeoutError(
                f"作业 {job_no} 在 {poll_seconds} 秒内未完成，"
                f"最后状态：{last_row}"
            )

        # 页面不会自动更新状态：先等待间隔，再在“作业”框重新查询，随即检查结果。
        requery = True
        remaining = poll_deadline - time.monotonic()
        time.sleep(min(JOB_STATUS_POLL_INTERVAL_SECONDS, max(0, remaining)))


def download_r55gl1005(in_page, company, destination, timeout=60000):
    """返回主页，提交 R55GL1005 报表并下载生成的 Excel 文件。"""
    home = in_page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()
    logger.info("已返回主页，重新选择页面并进入总帐")

    frame = in_page.frame_locator(MENU_IFRAME)
    _enter_total_ledger(frame, timeout)
    _click_menu_link(in_page, "中方科目总账", timeout)

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
    company_value = str(company).strip().lstrip("0") or "0"
    selection.fill("", timeout=timeout)
    selection.fill(company_value, timeout=timeout)
    frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click(timeout=timeout)

    frame.locator("a").filter(has_text="屏幕(F)").click(timeout=timeout)
    frame.get_by_role("cell", name="已提交作业", exact=True).click(
        timeout=timeout
    )

    job_filter = frame.get_by_role("textbox", name=JOB_FILTER_NAME, exact=True)
    job_filter.wait_for(state="visible", timeout=timeout)
    job_filter.fill(JOB_NAME, timeout=timeout)
    job_filter.press("Enter", timeout=timeout)
    # with page.expect_download() as download1_info:
    #     page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("link", name="可单击图标 查看输出").first.click()
    # download1 = download1_info.value
    # 本次提交的作业号每次都变，按“列表第一行 + 等待状态完成”定位，避免点到处理中的作业。
    output_link = _wait_latest_job_completed(frame, timeout)

    target = Path(destination)
    if not target.suffix:
        target = target.with_suffix(".xls")
    ensure_dir(target.parent)
    existing_pages = list(in_page.context.pages)
    download = _click_and_collect_download(in_page, output_link, timeout)
    download.save_as(str(target))
    # 点击后跳转出的页面已完成下载，直接关闭。
    _close_pages_opened_after(in_page.context, existing_pages)
    logger.info(f"R55GL1005 已导出：{target}")
    return target
    



# 下载 TB
def download_TB():
    logger.info("")
