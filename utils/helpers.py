#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具模块
"""
import functools
import time
from datetime import datetime
from pathlib import Path
from typing import Callable

from log.loggingConfig import logger


def retry(max_attempts: int = 3, delay: float = 1.0):
    """重试装饰器"""
    def decorator(func: Callable) -> Callable:
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_attempts):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt < max_attempts - 1:
                        logger.warning(f"{func.__name__} 失败: {e}, 重试中...")
                        time.sleep(delay)
                    else:
                        raise
        return wrapper
    return decorator


def log_execution(func: Callable) -> Callable:
    """日志装饰器"""
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        logger.info(f"开始: {func.__name__}")
        try:
            result = func(*args, **kwargs)
            logger.info(f"完成: {func.__name__}")
            return result
        except Exception as e:
            logger.error(f"失败: {func.__name__} - {e}")
            raise
    return wrapper


def ensure_dir(path: Path) -> Path:
    """确保目录存在"""
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_timestamp() -> str:
    """获取时间戳"""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def format_time(seconds: float) -> str:
    """格式化时间显示

    Args:
        seconds: 秒数

    Returns:
        格式化的时间字符串，如 "5.23秒"、"2分30.50秒"、"1时30分45.00秒"
    """
    if seconds < 60:
        return f"{seconds:.2f}秒"
    elif seconds < 3600:
        minutes = int(seconds // 60)
        secs = seconds % 60
        return f"{minutes}分{secs:.2f}秒"
    else:
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = seconds % 60
        return f"{hours}时{minutes}分{secs:.2f}秒"