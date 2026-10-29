#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Windows 客户端应用管理器（uiautomation封装）
"""
import subprocess
import time
from typing import Optional
from log.loggingConfig import logger

try:
    import uiautomation as auto
    HAS_UIAUTOMATION = True
except ImportError:
    HAS_UIAUTOMATION = False
    logger.warning("需要安装 uiautomation: pip install uiautomation")


class ClientManager:
    """Windows 客户端应用管理器

    封装 uiautomation 库，管理应用程序窗口
    支持启动新应用和连接已有窗口
    """

    def __init__(self, app_path: str = None, window_name: str = None, window_class: str = None):
        """初始化应用管理器

        Args:
            app_path: 应用程序路径（用于启动新应用）
            window_name: 窗口标题/名称（用于定位窗口）
            window_class: 窗口类名（用于定位窗口）
        """
        self.app_path = app_path
        self.window_name = window_name
        self.window_class = window_class
        self.window = None          # 主窗口对象
        self.process = None         # 进程对象

    def start(self, timeout: int = 30000) -> bool:
        """启动应用程序并定位主窗口

        Args:
            timeout: 等待窗口出现的超时时间（毫秒）

        Returns:
            bool: 是否成功启动并定位窗口
        """
        if not HAS_UIAUTOMATION:
            raise ImportError("需要安装 uiautomation: pip install uiautomation")

        if not self.app_path:
            logger.error("未指定应用程序路径")
            return False

        logger.info(f"启动应用程序: {self.app_path}")

        try:
            # 启动应用程序
            self.process = subprocess.Popen(self.app_path)
            logger.info(f"应用程序进程已启动，PID: {self.process.pid}")

            # 等待窗口出现
            self.window = self._wait_for_window(timeout)

            if self.window:
                logger.info("应用程序启动成功，已定位主窗口")
                return True
            else:
                logger.error("应用程序启动后未能定位主窗口")
                return False

        except Exception as e:
            logger.error(f"启动应用程序失败: {e}")
            return False

    def connect(self, window_name: str = None, window_class: str = None,
                timeout: int = 10000) -> bool:
        """连接已存在的应用程序窗口

        Args:
            window_name: 窗口标题/名称（可选，覆盖初始化时的设置）
            window_class: 窗口类名（可选，覆盖初始化时的设置）
            timeout: 查找窗口的超时时间（毫秒）

        Returns:
            bool: 是否成功连接窗口
        """
        if not HAS_UIAUTOMATION:
            raise ImportError("需要安装 uiautomation: pip install uiautomation")

        # 使用传入参数或初始化时的设置
        name = window_name or self.window_name
        class_name = window_class or self.window_class

        logger.info(f"连接已有窗口: name={name}, class={class_name}")

        self.window = self._find_window(name, class_name, timeout)

        if self.window:
            logger.info("成功连接窗口")
            return True
        else:
            logger.error("未能找到目标窗口")
            return False

    def close(self) -> None:
        """关闭应用程序"""
        if self.window:
            try:
                self.window.Close()
                logger.info("窗口已关闭")
            except Exception as e:
                logger.error(f"关闭窗口失败: {e}")

        if self.process:
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
                logger.info("应用程序进程已终止")
            except Exception as e:
                logger.error(f"终止进程失败: {e}")
                try:
                    self.process.kill()
                except:
                    pass

        self.window = None
        self.process = None

    def get_window(self):
        """获取主窗口对象"""
        return self.window

    def _wait_for_window(self, timeout: int) -> Optional[auto.WindowControl]:
        """等待窗口出现

        Args:
            timeout: 超时时间（毫秒）

        Returns:
            WindowControl 或 None
        """
        start_time = time.time()
        timeout_sec = timeout / 1000

        while time.time() - start_time < timeout_sec:
            window = self._find_window(self.window_name, self.window_class, 1000)
            if window:
                return window
            time.sleep(0.5)

        return None

    def _find_window(self, name: str = None, class_name: str = None,
                     timeout: int = 10000) -> Optional[auto.WindowControl]:
        """查找窗口

        Args:
            name: 窗口名称/标题
            class_name: 窗口类名
            timeout: 查找超时时间（毫秒）

        Returns:
            WindowControl 或 None
        """
        # 构建查找参数
        search_params = {}
        if name:
            search_params['Name'] = name
        if class_name:
            search_params['ClassName'] = class_name

        try:
            # 使用 uiautomation 查找窗口
            window = auto.WindowControl(**search_params)
            if window.Exists(maxSearchSeconds=timeout/1000):
                return window
            return None
        except Exception as e:
            logger.error(f"查找窗口失败: {e}")
            return None