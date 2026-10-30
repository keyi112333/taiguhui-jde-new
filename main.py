#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RPA框架入口
"""
import time
from pathlib import Path

from log.loggingConfig import logger
from config.config import USER_CONFIG_PATH, SYSTEM_CONFIG_PATH
from config import config
from init.init import initEnv
from workflows.jde_flow.main_task import JDEMainTask
from utils.helpers import format_time
from utils.period_dialog import ask_accounting_period


def main():
    """主函数"""
    # 记录流程开始时间
    flow_start_time = time.time()

    # 加载配置
    config_path = Path(USER_CONFIG_PATH)
    if not config_path.is_absolute():
        config_path = Path(__file__).resolve().parent / config_path
    if config_path.exists():
        system_path = Path(SYSTEM_CONFIG_PATH)
        if not system_path.is_absolute():
            system_path = Path(__file__).resolve().parent / system_path
        config.load(str(config_path), str(system_path))
        config.ensure_directories()
    else:
        logger.error(f"配置文件不存在，请检查路径是否正确: {config_path}")
        raise Exception(f"配置文件不存在，请检查路径是否正确: {config_path}")

    # 运行前弹窗输入会计期间：填写值优先，留空或超时按自然月运行。
    period = ask_accounting_period()
    config.set("房产税配置.会计期间", period or "")
    if period:
        logger.info(f"本次运行会计期间：{period}（弹窗输入）")
    else:
        logger.info("未输入会计期间，按自然月运行")

    #初始化环境
    initEnv()

    # 创建并运行主任务
    logger.info(f"{config.project_name}开始执行")

    main_task = JDEMainTask(config)

    if main_task.run():
        flow_end_time = time.time()
        elapsed_time = flow_end_time - flow_start_time
        logger.info(f"{config.project_name}完成执行, 总耗时: {format_time(elapsed_time)}")
        return True
    else:
        flow_end_time = time.time()
        elapsed_time = flow_end_time - flow_start_time
        logger.error(f"{config.project_name}执行失败，状态: {main_task.status}, 总耗时: {format_time(elapsed_time)}")
        return False


if __name__ == '__main__':
    raise SystemExit(0 if main() else 1)
