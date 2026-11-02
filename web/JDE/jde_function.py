
import functools
import re
import time
from pathlib import Path

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from log.loggingConfig import logger
from utils.helpers import ensure_dir,get_timestamp

MENU_IFRAME_NAME = "e1menuAppIframe"
MENU_IFRAME = f'iframe[name="{MENU_IFRAME_NAME}"]'
# 日记帐分录录入前需要选中的分录方案（JDE #gtab0_1 下拉值）。
JOURNAL_SCHEME_OPTION = "8d5afde055f149ea973fb07587dafa28"
# 粘贴前必须存在的列及顺序（TSV 按该顺序整块粘贴）。
JOURNAL_COLUMNS = ("帐号", "借项金额", "贷项金额", "备注")

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


def _focus_journal_grid(frame, in_page, timeout):
    """把焦点移到分录网格首格。

    表头字段（批号/公司/说明/日期）与网格单元都用 .JSTextfield，
    直接点第一个会把焦点留在日期栏，粘贴就会写进日期。这里只认
    “网格行内的输入框”获得焦点，确认不了就抛错，绝不盲目粘贴。
    """
    focused_selector = "tr[id] input.JSTextfield:focus"
    for selector in ("tr[id^='G0_'] input.JSTextfield", "#G0_1 input.JSTextfield"):
        cells = frame.locator(selector)
        if cells.count() == 0:
            continue
        cells.first.click(timeout=timeout)
        in_page.wait_for_timeout(300)
        if frame.locator(focused_selector).count() > 0:
            logger.info(f"焦点已移到分录网格：{selector}")
            return
    # 网格 id 不确定时，用 Tab 顺序进入网格，每步校验焦点。
    for _ in range(12):
        in_page.keyboard.press("Tab")
        in_page.wait_for_timeout(200)
        if frame.locator(focused_selector).count() > 0:
            logger.info("已用 Tab 把焦点移入分录网格")
            return
    raise RuntimeError("未能把焦点移到分录网格，已跳过粘贴，避免写入日期栏")


def _normalize_label(text):
    return re.sub(r"\s+", "", str(text or ""))


def _role_header_labels(frame):
    """读取网格 role=columnheader 的表头文字。"""
    labels = []
    try:
        headers = frame.get_by_role("columnheader")
        for index in range(headers.count()):
            item = headers.nth(index)
            try:
                if not item.is_visible():
                    continue
                label = _normalize_label(item.inner_text())
            except Exception:
                continue
            if label:
                labels.append(label)
    except Exception:
        return []
    return labels


def _th_header_labels(frame):
    """标准 table 表头兜底。"""
    labels = []
    try:
        headers = frame.locator("th")
        for index in range(headers.count()):
            item = headers.nth(index)
            try:
                if not item.is_visible():
                    continue
                label = _normalize_label(item.inner_text())
            except Exception:
                continue
            if label:
                labels.append(label)
    except Exception:
        return []
    return labels


def _app_frame(in_page):
    """返回 JDE 菜单 iframe 的真实 Frame，用于按坐标读取整行表头。"""
    for candidate in in_page.frames:
        try:
            if candidate.name == MENU_IFRAME_NAME:
                return candidate
        except Exception:
            continue
    return None


def _position_header_labels(frame, in_page):
    """按表头文字定位整行，返回该行从左到右的完整列名（用于连续校验）。"""
    inputs = frame.locator("#G0_1 input.JSTextfield")
    if inputs.count() == 0:
        inputs = frame.locator("tr[id^='G0_'] input.JSTextfield")
    if inputs.count() == 0:
        return []
    try:
        anchor = inputs.first.bounding_box()
    except Exception:
        return []
    if not anchor:
        return []
    boxes = []
    for name in JOURNAL_COLUMNS:
        locator = frame.get_by_text(name)
        found = []
        for index in range(min(locator.count(), 20)):
            item = locator.nth(index)
            try:
                if not item.is_visible():
                    continue
                box = item.bounding_box()
            except Exception:
                continue
            if not box or box["width"] <= 0 or box["height"] <= 0:
                continue
            if box["y"] + box["height"] <= anchor["y"] + 2:
                found.append(box)
        if not found:
            return []
        found.sort(key=lambda box: box["y"] + box["height"])
        boxes.append(found[-1])
    tops = [box["y"] for box in boxes]
    if max(tops) - min(tops) > 40:
        return []
    real_frame = _app_frame(in_page)
    if real_frame is None:
        return []
    top = min(box["y"] for box in boxes)
    bottom = max(box["y"] + box["height"] for box in boxes)
    left = min(box["x"] for box in boxes)
    right = max(box["x"] + box["width"] for box in boxes)
    try:
        items = real_frame.evaluate(
            """
            ({top, bottom, left, right}) => {
                const result = [];
                for (const node of document.querySelectorAll('td, th, div, span, a, label')) {
                    const rect = node.getBoundingClientRect();
                    if (rect.width <= 0 || rect.height <= 0) continue;
                    const centerY = rect.top + rect.height / 2;
                    if (centerY < top || centerY > bottom) continue;
                    if (rect.right <= left || rect.left >= right) continue;
                    let hasChildText = false;
                    for (const child of node.children) {
                        if ((child.innerText || '').trim()) { hasChildText = true; break; }
                    }
                    if (hasChildText) continue;
                    const text = (node.innerText || '').replace(/\\s+/g, '');
                    if (!text || text.length > 30) continue;
                    result.push({text: text, x: rect.left, w: rect.width});
                }
                return result;
            }
            """,
            {"top": top, "bottom": bottom, "left": left + 1, "right": right - 1},
        )
    except Exception:
        return []
    items.sort(key=lambda item: (item["x"], item["w"]))
    labels, seen = [], set()
    for item in items:
        key = (item["text"], round(item["x"]))
        if key in seen:
            continue
        seen.add(key)
        labels.append(item["text"])
    return labels


def _journal_header_labels(frame, in_page):
    """返回能找到全部目标列的表头顺序；找不到返回空列表。"""
    for labels in (_role_header_labels(frame), _th_header_labels(frame),
                   _position_header_labels(frame, in_page)):
        if labels and all(name in labels for name in JOURNAL_COLUMNS):
            return labels
    return []


def _verify_journal_columns(frame, in_page, timeout):
    """粘贴前校验网格列顺序，不通过则中止，避免数据写入错误列。"""
    wait_timeout = min(timeout, 5000)
    try:
        frame.locator("#G0_1 input.JSTextfield").first.wait_for(
            state="visible", timeout=wait_timeout
        )
    except Exception:
        try:
            frame.locator("tr[id^='G0_'] input.JSTextfield").first.wait_for(
                state="visible", timeout=wait_timeout
            )
        except Exception:
            pass
    labels = _journal_header_labels(frame, in_page)
    if not labels:
        raise RuntimeError("无法读取日记帐分录网格表头，未执行粘贴")
    width = len(JOURNAL_COLUMNS)
    matched = any(
        tuple(labels[start:start + width]) == JOURNAL_COLUMNS
        for start in range(len(labels) - width + 1)
    )
    if not matched:
        missing = [name for name in JOURNAL_COLUMNS if name not in labels]
        if missing:
            raise RuntimeError(
                f"分录网格缺少列：{'、'.join(missing)}；当前表头：{' / '.join(labels)}"
            )
        raise RuntimeError(
            f"分录网格列顺序不符（四列必须连续）：期望 {'、'.join(JOURNAL_COLUMNS)}，"
            f"当前表头：{' / '.join(labels)}"
        )
    logger.info(f"分录网格列顺序校验通过：{' / '.join(labels)}")


def _click_cancel(frame, timeout):
    """在当前 frame 点击一次取消，失败返回 False。"""
    try:
        button = frame.get_by_role("button", name="取消 (Ctrl+Alt+L)").first
        if button.count() == 0 or not button.is_visible():
            return False
        button.click(timeout=timeout)
        return True
    except Exception as exc:
        logger.warning(f"点击取消失败：{exc}")
        return False


def _click_cancel_anywhere(in_page, timeout, prefer_new=False):
    """跨页面/frame 查找并点击取消；找不到返回 False。"""
    pages = list(in_page.context.pages)
    if prefer_new:
        pages = [page for page in pages if page != in_page] + [in_page]
    for page in pages:
        try:
            frames = page.frames
        except Exception:
            continue
        for frame in frames:
            try:
                buttons = frame.get_by_role("button", name="取消 (Ctrl+Alt+L)")
                for index in range(min(buttons.count(), 5)):
                    button = buttons.nth(index)
                    if button.is_visible():
                        button.click(timeout=timeout)
                        return True
            except Exception:
                continue
    return False


def _discard_entry(frame, in_page, timeout):
    """取消分录录入：第二次取消可能不存在或出现在新页面，按尽力而为处理。"""
    if not _click_cancel(frame, timeout):
        _click_cancel_anywhere(in_page, timeout)
    in_page.wait_for_timeout(800)
    if _click_cancel_anywhere(in_page, min(timeout, 5000), prefer_new=True):
        logger.info("已点击第二次取消")
    else:
        logger.info("未发现第二次取消按钮，跳过")


def create_journal_entry(in_page, company, description, gl_date, rows,
                         scheme_option=JOURNAL_SCHEME_OPTION,
                         timeout=60000, discard=True):
    """录入日记帐分录并返回 JDE 生成的批号。

    rows 为 (帐号, 借项金额, 贷项金额, 备注) 序列，按 TSV 整块粘贴到第一个单元格，
    行数随数据动态变化；粘贴前先选分录方案并校验网格列顺序。
    discard=True 时按录制脚本取消退出（第二次取消可能不存在或位于新页面）。
    """
    if not rows:
        raise ValueError("分录没有可录入的数据行")
    # 先回主页，再重选页面进入总帐，避免停留在上一模块的页面状态。
    home = in_page.get_by_title("主页")
    home.wait_for(state="visible", timeout=timeout)
    home.click()
    logger.info("已返回主页，重新选择页面并进入总帐")
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
    # 粘贴前先选好分录方案，否则网格不接受整块粘贴；选完等网格重绘。
    scheme = frame.locator("#gtab0_1")
    scheme.wait_for(state="visible", timeout=timeout)
    scheme.select_option(scheme_option, timeout=timeout)
    logger.info(f"已选择分录方案：{scheme_option}")
    in_page.wait_for_timeout(1000)
    # 列顺序不符时不粘贴，先取消录入再报错，避免数据写入错误列。
    try:
        _verify_journal_columns(frame, in_page, timeout)
    except Exception:
        _discard_entry(frame, in_page, timeout)
        raise
    _focus_journal_grid(frame, in_page, timeout)
    in_page.keyboard.press("Control+V")
    in_page.wait_for_timeout(1000)

    if discard:
        # 正式环境开发不允许保存；第二次取消可能不存在或在其他页面。
        _discard_entry(frame, in_page, timeout)
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
