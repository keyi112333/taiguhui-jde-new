import re
from playwright.sync_api import Page, expect


def test_example(page: Page) -> None:
    page.get_by_title("主页").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#tabBar").get_by_text("总帐").click()
    page.get_by_role("link", name="中方科目总账").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#G0_1_R9 > .JSSelectGrid > .cellCentered > input").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="行:").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="数据选择").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("a").filter(has_text="提交(U)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("input[name=\"E1\"]").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("input[name=\"E1\"]").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("td").filter(has_text=re.compile(r"^Company$")).click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("input[name=\"E1\"]").fill("2501")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("a").filter(has_text="屏幕(F)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_text("已提交作业").click()
    with page.expect_download() as download_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("link", name="R55GL1005_TKHXLS001").click()
    download = download_info.value
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_text("IFrame support is required.If").click()
    with page.expect_download() as download1_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("link", name="RD55GL1005_EN_511120_26559_208.xls").click()
    download1 = download1_info.value


    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#tabBar").get_by_text("总帐").click()
    page.get_by_role("link", name="中方科目总账").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#G0_1_R9 > .JSSelectGrid > .cellCentered > input").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_text("TKHXLS001").click()   #需要定位到这个版本再点下面的勾选，下面的勾选要对应这个
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="行:").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="数据选择").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("a").filter(has_text="提交(U)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("input[name=\"E1\"]").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("input[name=\"E1\"]").fill("2501")    #这一步需要清空输入框再输入
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("a").filter(has_text="屏幕(F)").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("cell", name="已提交作业", exact=True).click()
    with page.expect_download() as download_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("link", name="R55GL1005_TKHXLS001").first.click() #可能存在多个，选择最上面的
    download = download_info.value
    with page.expect_download() as download1_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("link", name="click here").click()  #点击后会下载文件
    download1 = download1_info.value
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("table").filter(has_text="If file download does not").click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="关闭").click()


    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("CapsLock")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).fill("TKHXLS001")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("Enter")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="行:").check()

    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).click()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("CapsLock")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).fill("TKHXLS001")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("textbox", name="版本", exact=True).press("Enter")
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("checkbox", name="选择所有提取的记录").check()
    page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("button", name="选择 (Ctrl+Alt+S)").click()

    with page.expect_download() as download1_info:
        page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.get_by_role("link", name="R55GL1005_TKHXLS001").first.click()
    download1 = download1_info.value