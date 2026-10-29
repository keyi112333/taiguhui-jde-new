#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
窗口操作基类
"""
import time
from typing import Optional
from log.loggingConfig import logger

try:
    import uiautomation as auto
    HAS_UIAUTOMATION = True
except ImportError:
    HAS_UIAUTOMATION = False
    logger.warning("需要安装 uiautomation: pip install uiautomation")


class BaseWindow:
    """窗口操作基类

    提供通用的 Windows 控件操作方法
    支持多种定位方式：Name/Text、ControlType、ClassName、索引
    """

    # 控件类型映射（uiautomation 控件类名）
    CONTROL_TYPES = {
        'button': auto.ButtonControl,
        'edit': auto.EditControl,
        'text': auto.TextControl,
        'list': auto.ListControl,
        'listitem': auto.ListItemControl,
        'tree': auto.TreeControl,
        'treeitem': auto.TreeItemControl,
        'combobox': auto.ComboBoxControl,
        'checkbox': auto.CheckBoxControl,
        'radiobutton': auto.RadioButtonControl,
        'tab': auto.TabControl,
        'tabitem': auto.TabItemControl,
        'menu': auto.MenuControl,
        'menuitem': auto.MenuItemControl,
        'toolbar': auto.ToolBarControl,
        'pane': auto.PaneControl,
        'document': auto.DocumentControl,
        'image': auto.ImageControl,
        'hyperlink': auto.HyperlinkControl,
        'slider': auto.SliderControl,
        'spinner': auto.SpinnerControl,
        'datagrid': auto.DataGridControl,
        'dataitem': auto.DataItemControl,
        'scrollbar': auto.ScrollBarControl,
        'thumb': auto.ThumbControl,
        'titlebar': auto.TitleBarControl,
        'statusbar': auto.StatusBarControl,
        'window': auto.WindowControl,
        'group': auto.GroupControl,
        'splitbutton': auto.SplitButtonControl,
        'progressbar': auto.ProgressBarControl,
    }

    def __init__(self, window):
        """初始化窗口操作基类

        Args:
            window: 窗口对象（WindowControl 或 ClientManager）
        """
        # 支持传入 ClientManager 或直接的窗口对象
        if hasattr(window, 'get_window'):
            self.window = window.get_window()
        else:
            self.window = window

    def find_control(self, control_type: str = None,
                     name: str = None,
                     class_name: str = None,
                     index: int = None,
                     depth: int = None) -> Optional[auto.Control]:
        """查找控件

        支持多种定位方式组合：
        - control_type: 控件类型如 'Button', 'Edit', 'List'
        - name: 控件显示名称/文本
        - class_name: 控件类名
        - index: 在父容器中的索引位置
        - depth: 搜索深度

        Args:
            control_type: 控件类型字符串
            name: 控件名称/文本
            class_name: 控件类名
            index: 索引位置
            depth: 搜索深度

        Returns:
            Control 对象或 None
        """
        if not HAS_UIAUTOMATION:
            raise ImportError("需要安装 uiautomation: pip install uiautomation")

        search_params = {}

        # 处理控件类型
        if control_type:
            control_type_lower = control_type.lower()
            if control_type_lower in self.CONTROL_TYPES:
                control_class = self.CONTROL_TYPES[control_type_lower]
            else:
                # 尝试直接使用传入的类型名
                control_class = getattr(auto, f'{control_type}Control', None)
                if not control_class:
                    logger.warning(f"未知的控件类型: {control_type}")
                    control_class = auto.Control

        else:
            control_class = auto.Control

        # 其他参数
        if name:
            search_params['Name'] = name
        if class_name:
            search_params['ClassName'] = class_name
        if depth:
            search_params['searchDepth'] = depth

        try:
            # 在窗口中查找控件
            if index is not None:
                # 使用索引查找
                controls = self.window.GetChildren()
                filtered = []
                for ctrl in controls:
                    match = True
                    if name and ctrl.Name != name:
                        match = False
                    if class_name and ctrl.ClassName != class_name:
                        match = False
                    if control_type:
                        ctrl_type = ctrl.ControlTypeName
                        if control_type_lower not in ctrl_type.lower():
                            match = False
                    if match:
                        filtered.append(ctrl)

                if index < len(filtered):
                    return filtered[index]
                return None
            else:
                # 直接查找
                control = control_class(self.window, **search_params)
                if control.Exists(maxSearchSeconds=1):
                    return control
                return None

        except Exception as e:
            logger.error(f"查找控件失败: {e}")
            return None

    def click(self, control, timeout: int = 10000) -> None:
        """点击控件

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            logger.info(f"点击控件: {ctrl.Name or ctrl.ControlTypeName}")
            ctrl.Click()
        else:
            raise Exception("无法找到要点击的控件")

    def double_click(self, control, timeout: int = 10000) -> None:
        """双击控件

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            logger.info(f"双击控件: {ctrl.Name or ctrl.ControlTypeName}")
            ctrl.DoubleClick()
        else:
            raise Exception("无法找到要双击的控件")

    def input_text(self, control, text: str, timeout: int = 10000,
                   clear_first: bool = True) -> None:
        """输入文本

        Args:
            control: 控件对象或定位参数字典
            text: 要输入的文本
            timeout: 等待时间（毫秒）
            clear_first: 是否先清空现有内容
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            logger.info(f"输入文本: {text}")
            if clear_first:
                ctrl.SetValue('')
            ctrl.SetValue(text)
        else:
            raise Exception("无法找到要输入的控件")

    def get_text(self, control, timeout: int = 10000) -> str:
        """获取控件文本

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）

        Returns:
            控件文本内容
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            return ctrl.Name or ''
        return ''

    def wait_for_control(self, control_type: str = None,
                         name: str = None,
                         class_name: str = None,
                         index: int = None,
                         timeout: int = 30000) -> Optional[auto.Control]:
        """等待控件出现

        Args:
            control_type: 控件类型
            name: 控件名称
            class_name: 控件类名
            index: 索引位置
            timeout: 超时时间（毫秒）

        Returns:
            Control 对象或 None
        """
        logger.info(f"等待控件: type={control_type}, name={name}")

        start_time = time.time()
        timeout_sec = timeout / 1000

        while time.time() - start_time < timeout_sec:
            ctrl = self.find_control(control_type, name, class_name, index)
            if ctrl:
                logger.info("控件已出现")
                return ctrl
            time.sleep(0.5)

        logger.warning(f"等待控件超时: type={control_type}, name={name}")
        return None

    def screenshot(self, path: str) -> None:
        """窗口截图

        Args:
            path: 截图保存路径
        """
        logger.info(f"截图保存: {path}")
        if self.window:
            self.window.CaptureToImage(path)

    def is_visible(self, control, timeout: int = 10000) -> bool:
        """判断控件是否可见

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）

        Returns:
            是否可见
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            return ctrl.IsVisible
        return False

    def select_item(self, control, item: str, timeout: int = 10000) -> None:
        """选择列表/树形控件项

        Args:
            control: 控件对象或定位参数字典
            item: 要选择的项名称
            timeout: 等待时间（毫秒）
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            logger.info(f"选择项: {item}")
            # 尝试 Select 方法
            if hasattr(ctrl, 'Select'):
                ctrl.Select(item)
            else:
                # 对于 ListControl，查找子项并点击
                item_ctrl = ctrl.GetChildByName(item)
                if item_ctrl:
                    item_ctrl.Click()
                else:
                    raise Exception(f"找不到项: {item}")
        else:
            raise Exception("无法找到列表控件")

    def expand(self, control, timeout: int = 10000) -> None:
        """展开树节点/下拉框

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            logger.info(f"展开控件: {ctrl.Name or ctrl.ControlTypeName}")
            if hasattr(ctrl, 'Expand'):
                ctrl.Expand()
            elif hasattr(ctrl, 'Click'):  # ComboBox 可能需要点击展开
                ctrl.Click()
        else:
            raise Exception("无法找到要展开的控件")

    def collapse(self, control, timeout: int = 10000) -> None:
        """折叠树节点/下拉框

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）
        """
        ctrl = self._get_control(control, timeout)
        if ctrl:
            logger.info(f"折叠控件: {ctrl.Name or ctrl.ControlTypeName}")
            if hasattr(ctrl, 'Collapse'):
                ctrl.Collapse()
        else:
            raise Exception("无法找到要折叠的控件")

    def send_keys(self, keys: str) -> None:
        """发送快捷键到窗口

        Args:
            keys: 快捷键字符串，如 '{Ctrl}{C}', '{Enter}', 'Hello'
        """
        logger.info(f"发送快捷键: {keys}")
        if self.window:
            auto.SendKeys(keys)

    def wait(self, timeout: int = 1000) -> None:
        """等待指定时间

        Args:
            timeout: 等待时间（毫秒）
        """
        time.sleep(timeout / 1000)

    def _get_control(self, control, timeout: int = 10000):
        """获取控件对象

        支持传入控件对象或定位参数字典

        Args:
            control: 控件对象或定位参数字典
            timeout: 等待时间（毫秒）

        Returns:
            Control 对象
        """
        if isinstance(control, dict):
            return self.wait_for_control(**control, timeout=timeout)
        elif hasattr(control, 'Exists'):  # 是 uiautomation 控件
            return control
        else:
            return self.find_control(control)