---
name: rpa-framework-guard
description: RPA框架使用规范守卫。当用户在这个RPA框架项目中开发、修改代码时触发。规范核心模块保护、扩展方式指导。TRIGGER when: 用户请求修改core/、web/、data_processing/、utils/、config/、log/、init/目录下的代码，或询问如何扩展框架功能。
---

# RPA框架使用规范

## 核心原则

**保护框架核心，业务代码外延**

框架核心模块已稳定，业务开发应在 `workflows/` 和 `web/<系统名>/` 中进行。

## 模块保护级别

| 模块 | 保护级别 | 说明 |
|------|----------|------|
| `core/` | 禁止修改 | 任务基类、上下文、编排器 |
| `config/` | 禁止修改 | 配置加载、路径管理 |
| `log/` | 禁止修改 | 日志配置 |
| `init/` | 谨慎修改 | 环境初始化 |
| `utils/` | 可扩展 | 可添加新工具函数 |
| `web/browser.py` | 禁止修改 | 浏览器管理基类 |
| `web/base_page.py` | 谨慎修改 | 页面操作基类 |
| `web/<系统名>/` | 自由修改 | 各系统的业务模块 |
| `data_processing/` | 可扩展 | 可添加新数据处理类 |
| `workflows/` | 自由修改 | 业务流程目录 |

## 正确的扩展方式

### 添加新系统
在 `web/` 下创建新目录，如 `web/new_system/`

### 添加新流程
在 `workflows/` 下创建新目录，如 `workflows/my_flow/`

### 添加新工具
在 `utils/` 下添加新文件，如 `utils/my_tool.py`

### 添加通用工具函数
通用工具函数（如时间格式化、字符串处理等）统一放到 `utils/helpers.py`，便于复用

### 添加数据处理
在 `data_processing/` 下添加新文件

## 修改核心模块的流程

如果确实需要修改 禁止修改 或 谨慎修改 标记的模块：

1. **先询问用户**：说明修改原因、影响范围、替代方案
2. **等待用户确认**：用户同意后方可执行
3. **记录修改**：更新README.md中的变更说明

## 违规检测

当检测到以下行为时，应拒绝并引导正确方式：

- 直接修改 `core/task_base.py` -> 在 `workflows/` 中继承扩展
- 直接修改 `web/browser.py` -> 在 `web/<系统>/` 中封装调用
- 直接修改 `config/config.py` -> 使用 `config.get()` 获取配置