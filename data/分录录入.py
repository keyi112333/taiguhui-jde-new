#执行前提：更新表金额核对通过，且在重算阶段需要将房产税（从租）计算底稿中的G-J的表格（具体可以看模板文件）的账号、借项金额、贷项金额、备注这几列下的值（不要标题）存为可复制黏贴的excel表格格式（行数是动态的，要人工操作的从1excel文件复制然后再黏贴）

page.get_by_title("主页").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator("#tabBar").get_by_text("总帐").click()
page.get_by_role("link", name="日记帐分录").click()    #直接复用另外两个页面总帐选择页面点击页面的操作
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="增加 (Ctrl+Alt+A)").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="确定 (Ctrl+Alt+O)").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="批号").click()   #需要记录这个值回写会更新表：房产税（从租）计算底稿中“批号：”单元格右边第一个单元格
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="Document Company").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="Document Company").fill("02501")   #先默认02501
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="说明").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="说明").fill("计提2026年9月从租房产税")   #读取更新表格
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="总帐日期").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("textbox", name="总帐日期").fill("10/29/2026")  #运行当天的实时日期
page.locator("iframe[name=\"e1menuAppIframe\"]").content_frame.locator("#gtab0_1").select_option("8d5afde055f149ea973fb07587dafa28")
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.locator(".JSTextfield").first.click()   #这个是第一个单元格，点击后黏贴我们前面提到的房产税（从租）计算底稿中的G-J的表格数据

page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="取消 (Ctrl+Alt+L)").click()
page.locator("iframe[name=\"elmenuAppIframe\"]").content_frame.get_by_role("button", name="取消 (Ctrl+Alt+L)").click()  #正式环境开发不可以保存所以要取消