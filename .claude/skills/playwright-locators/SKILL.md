---
name: playwright-locators
description: |
  Playwright 元素定位最佳实践。用于编写稳定、可维护的自动化脚本。

  TRIGGER when: 编写 Playwright 自动化脚本、定位网页元素、调试选择器问题、创建 Page Object。

  DO NOT trigger when: 使用 Selenium、Puppeteer、或其他自动化工具。
---

# Playwright 元素定位最佳实践

## 选择器优先级

从优到劣：

```python
# 1. data-testid（最稳定）- 专为测试设计
page.locator('[data-testid="login-btn"]')

# 2. role + name（语义化）
page.get_by_role("button", name="登录")
page.get_by_role("textbox", name="用户名")

# 3. label/placeholder（表单元素）
page.get_by_label("邮箱")
page.get_by_placeholder("请输入密码")

# 4. text 内容
page.get_by_text("Download Vendor List")

# 5. CSS 选择器（语义化 class/id）
page.locator('#submit-btn')

# 6. XPath（最后选择）
page.locator('//button[contains(text(),"登录")]')
```

## 选择器速查表

| 元素类型 | 推荐选择器 |
|---------|-----------|
| 按钮 | `get_by_role("button", name="xx")` |
| 输入框 | `get_by_label("xx")` 或 `get_by_placeholder("xx")` |
| 链接 | `get_by_role("link", name="xx")` |
| 下拉框 | `get_by_label("xx")` |
| 复选框 | `get_by_role("checkbox", name="xx")` |
| 图标按钮 | `get_by_label("xx")` (aria-label) |
| 表格行 | `data-testid` + 链式定位 |

## 避免使用

```python
# ❌ 结构依赖 - 页面结构变化就失效
page.locator('div:nth-child(3) > span:nth-child(2)')

# ❌ 动态 class - 自动生成的 class 会变
page.locator('.css-1a2b3c4d')

# ❌ 完整路径 - 太脆弱
page.locator('/html/body/div[2]/form/button')
```

## 稳定定位技巧

```python
# 组合条件
page.locator('button').filter(has_text="下载").first

# 模糊文本匹配
page.locator('text=/下载/i')  # 正则，忽略大小写

# 等待元素可见再操作
locator = page.locator('[data-testid="btn"]')
locator.wait_for(state="visible")
locator.click()

# 链式定位（先定位容器）
table = page.locator('[data-testid="data-table"]')
table.locator('tr').first.locator('td').nth(2)
```

## Page Object 示例

```python
from playwright.sync_api import Page, Locator

class LoginPage:
    def __init__(self, page: Page):
        self.page = page
        # 使用语义化选择器
        self.email = page.get_by_placeholder("Email")
        self.password = page.get_by_placeholder("Password")
        self.submit = page.get_by_role("button", name="Log In")

    def login(self, email: str, password: str):
        self.email.fill(email)
        self.password.fill(password)
        self.submit.click()
        self.page.wait_for_load_state("networkidle")
```

## 选择器集中管理

```python
class AppPage:
    SELECTORS = {
        "login_btn": "button[type='submit']",
        "vendors_menu": "text=Vendors",
        "download_btn": "text=Download Vendor List",
    }

    def click_download(self):
        self.page.locator(self.SELECTORS["download_btn"]).click()
```

## 调试技巧

```bash
# 代码生成器 - 自动生成选择器
playwright codegen https://example.com

# 慢动作执行
browser = p.chromium.launch(headless=False, slow_mo=1000)
```

```python
# 高亮元素
page.locator('button').highlight()

# 查看匹配数量
print(page.locator('button').count())

# 截图调试
page.screenshot(path="debug.png")
```

## 常见问题

| 问题 | 解决方案 |
|------|---------|
| 元素不可点击 | `locator.wait_for(state="visible")` |
| 有多个匹配 | `.first` 或 `.filter()` 缩小范围 |
| 动态加载 | `page.wait_for_load_state("networkidle")` |
| iframe 内元素 | `page.frame_locator('iframe').locator('...')` |
| 新窗口 | `page.wait_for_event('popup')` |