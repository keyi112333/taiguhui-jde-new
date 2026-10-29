import re
from playwright.sync_api import Page, expect


def test_example(page: Page) -> None:
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="选择页面").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("gridcell", name="所有页面").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_text("所有页面").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="选择页面").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#tabBar").get_by_text("总帐").click()
    page.get_by_role("link", name="按科目帐的帐户查询").click()

    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#tabBar").get_by_text("总帐").click()
    page.get_by_role("link", name="按科目帐的帐户查询").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="分类帐类型 2", exact=True).click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="分类帐类型 2", exact=True).press("CapsLock")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="分类帐类型 2", exact=True).fill("AA")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="分类帐类型 2", exact=True).press("CapsLock")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="自公司").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="自公司").fill("02501")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="至公司").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="至公司").fill("02504")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="起始日期").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="起始日期").fill("10/01/2026")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="至日期").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="至日期").fill("10/31/2026")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="查找 (Ctrl+Alt+I)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="选择所有提取的记录").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="输出网格数据 (Ctrl + Shift +E)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="本地文件系统").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="输出到 Excel").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="输出到新工作簿 输出到新工作簿").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="所有").check()
    with page.expect_download() as download_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("cell", name="继续 (Ctrl+Alt+E)", exact=True).locator("a").click()
    download = download_info.value


    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="转到结束").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="选择所有提取的记录").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="输出网格数据 (Ctrl + Shift +E)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="本地文件系统").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="输出到 Excel").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="输出到新工作簿 输出到新工作簿").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("radio", name="所有").check()
    with page.expect_download() as download_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("cell", name="继续 (Ctrl+Alt+E)", exact=True).locator("a").click()
    download = download_info.value