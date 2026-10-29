#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
任务管理器（可选的工具类）
"""
from typing import List
from core.task_base import TaskBase


class TaskManager(TaskBase):
    """任务编排器（可选使用）

    如果需要简单的串行执行，可以使用此类。
    如果需要更灵活的控制，直接继承 TaskBase 自由编写 execute 方法。
    """

    def __init__(self, name: str, context):
        super().__init__(name, context)
        self._subtasks: List[TaskBase] = []

    def execute(self) -> bool:
        """按顺序执行所有子任务（简单串行模式）"""
        for subtask in self._subtasks:
            if not subtask.run():
                return False
        return True

    def add_subtask(self, subtask: TaskBase) -> None:
        """添加子任务"""
        self._subtasks.append(subtask)

    def get_subtask_status(self) -> dict:
        """获取所有子任务状态"""
        return {subtask.name: subtask.status for subtask in self._subtasks}