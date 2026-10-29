#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件操作工具
"""
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional
from log.loggingConfig import logger


def copy_file(src: str, dst: str) -> bool:
    """复制文件"""
    try:
        shutil.copy2(src, dst)
        logger.info(f"复制文件: {src} -> {dst}")
        return True
    except Exception as e:
        logger.error(f"复制文件失败: {e}")
        return False


def move_file(src: str, dst: str) -> bool:
    """移动文件"""
    try:
        shutil.move(src, dst)
        logger.info(f"移动文件: {src} -> {dst}")
        return True
    except Exception as e:
        logger.error(f"移动文件失败: {e}")
        return False


def delete_file(path: str) -> bool:
    """删除文件"""
    try:
        Path(path).unlink()
        logger.info(f"删除文件: {path}")
        return True
    except Exception as e:
        logger.error(f"删除文件失败: {e}")
        return False


def ensure_dir(path: str) -> Path:
    """确保目录存在"""
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def get_file_list(dir_path: str, pattern: str = "*") -> list:
    """获取目录下的文件列表"""
    p = Path(dir_path)
    if not p.exists():
        return []
    return list(p.glob(pattern))


def get_latest_file(dir_path: str, pattern: str = "*") -> Optional[Path]:
    """获取目录下最新的文件"""
    files = get_file_list(dir_path, pattern)
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)


def backup_file(src: str, backup_dir: str) -> Optional[str]:
    """备份文件"""
    try:
        src_path = Path(src)
        if not src_path.exists():
            return None

        backup_path = Path(backup_dir)
        backup_path.mkdir(parents=True, exist_ok=True)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{src_path.stem}_{timestamp}{src_path.suffix}"
        backup_file_path = backup_path / backup_name

        shutil.copy2(src_path, backup_file_path)
        logger.info(f"备份文件: {src} -> {backup_file_path}")
        return str(backup_file_path)
    except Exception as e:
        logger.error(f"备份文件失败: {e}")
        return None


def file_exists(path: str) -> bool:
    """检查文件是否存在"""
    return Path(path).exists()


def get_file_size(path: str) -> int:
    """获取文件大小（字节）"""
    try:
        return Path(path).stat().st_size
    except:
        return 0