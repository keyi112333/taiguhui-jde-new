#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
截图工具（通用系统截图）
"""
from datetime import datetime
from pathlib import Path
from typing import Optional
from log.loggingConfig import logger

try:
    from PIL import ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    logger.warning("需要安装 Pillow: pip install Pillow")


def take_screenshot(save_dir: str, filename: Optional[str] = None) -> Optional[str]:
    """系统截图（截取整个屏幕）

    Args:
        save_dir: 保存目录
        filename: 文件名（可选，默认使用时间戳）

    Returns:
        str: 截图保存路径，失败返回None
    """
    if not HAS_PIL:
        logger.error("Pillow未安装，无法截图")
        return None

    try:
        save_path = Path(save_dir)
        save_path.mkdir(parents=True, exist_ok=True)

        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshot_{timestamp}.png"

        screenshot_path = save_path / filename

        # 截取整个屏幕
        img = ImageGrab.grab()
        img.save(str(screenshot_path))

        logger.info(f"截图已保存: {screenshot_path}")
        return str(screenshot_path)
    except Exception as e:
        logger.error(f"截图失败: {e}")
        return None


def take_region_screenshot(save_dir: str, bbox: tuple, filename: Optional[str] = None) -> Optional[str]:
    """截取屏幕指定区域

    Args:
        save_dir: 保存目录
        bbox: 截图区域 (left, top, right, bottom)
        filename: 文件名（可选）

    Returns:
        str: 截图保存路径
    """
    if not HAS_PIL:
        logger.error("Pillow未安装，无法截图")
        return None

    try:
        save_path = Path(save_dir)
        save_path.mkdir(parents=True, exist_ok=True)

        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"region_{timestamp}.png"

        screenshot_path = save_path / filename

        # 截取指定区域
        img = ImageGrab.grab(bbox=bbox)
        img.save(str(screenshot_path))

        logger.info(f"区域截图已保存: {screenshot_path}")
        return str(screenshot_path)
    except Exception as e:
        logger.error(f"区域截图失败: {e}")
        return None


def take_error_screenshot(save_dir: str, error_name: str) -> Optional[str]:
    """错误时自动截图

    Args:
        save_dir: 保存目录（通常是错误截图目录）
        error_name: 错误名称（用于文件命名）

    Returns:
        str: 截图保存路径
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"error_{error_name}_{timestamp}.png"
    return take_screenshot(save_dir, filename)