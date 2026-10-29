#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
页面操作基类
"""
from typing import Optional
from log.loggingConfig import logger


class BasePage:
    """页面操作基类

    提供通用的页面操作方法：点击、输入、等待、截图等
    """

    def __init__(self, page):
        self.page = page

    def goto(self, url: str) -> None:
        """访问页面"""
        logger.info(f"访问页面: {url}")
        self.page.goto(url)

    def click(self, selector: str, timeout: int = 10000) -> None:
        """点击元素"""
        logger.info(f"点击元素: {selector}")
        self.page.click(selector, timeout=timeout)

    def fill(self, selector: str, value: str, timeout: int = 10000) -> None:
        """输入文本"""
        logger.info(f"输入文本: {selector} -> {value}")
        self.page.fill(selector, value, timeout=timeout)

    def type(self, selector: str, value: str, delay: int = 50, timeout: int = 10000) -> None:
        """逐字符输入（模拟真实输入）"""
        logger.info(f"逐字符输入: {selector}")
        self.page.type(selector, value, delay=delay, timeout=timeout)

    def wait_for_selector(self, selector: str, timeout: int = 30000) -> None:
        """等待元素出现"""
        logger.info(f"等待元素: {selector}")
        self.page.wait_for_selector(selector, timeout=timeout)

    def wait_for_url(self, url: str, timeout: int = 30000) -> None:
        """等待URL变化"""
        logger.info(f"等待URL: {url}")
        self.page.wait_for_url(url, timeout=timeout)

    def get_text(self, selector: str) -> str:
        """获取元素文本"""
        return self.page.locator(selector).text_content()

    def get_value(self, selector: str) -> str:
        """获取输入框值"""
        return self.page.input_value(selector)

    def screenshot(self, path: str) -> None:
        """截图"""
        logger.info(f"截图保存: {path}")
        self.page.screenshot(path=path)

    def is_visible(self, selector: str) -> bool:
        """判断元素是否可见"""
        return self.page.locator(selector).is_visible()

    def select_option(self, selector: str, value: str) -> None:
        """选择下拉框选项"""
        logger.info(f"选择选项: {selector} -> {value}")
        self.page.select_option(selector, value)

    def check(self, selector: str) -> None:
        """勾选复选框"""
        logger.info(f"勾选: {selector}")
        self.page.check(selector)

    def uncheck(self, selector: str) -> None:
        """取消勾选"""
        logger.info(f"取消勾选: {selector}")
        self.page.uncheck(selector)

    def wait(self, timeout: int = 1000) -> None:
        """等待指定时间"""
        self.page.wait_for_timeout(timeout)