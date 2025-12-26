[![English](https://img.shields.io/badge/lang-English-blue.svg)](README.en.md)
[![Chinese](https://img.shields.io/badge/lang-简体中文-red.svg)](README.zh-CN.md)
[![Japanese](https://img.shields.io/badge/lang-日本語-green.svg)](README.md)

----

# 项目分析：Office 文件质量检查工具

## 项目简介

本项目是一款基于 Python 的 GUI 应用程序，旨在从 Office 文档（特别支持 `.docx`、`.xlsx`、`.xlsm` 和 `.pptx` 格式）中提取内容。该工具采用 Tkinter 构建用户界面，并利用专业的 Python 库进行文档解析。

**核心功能包括：**

* **文本提取：** 从 Office 文档中提取纯文本内容并保存为 Markdown (`.md`) 文件。提取的文件将按时间戳组织在用户指定的输出路径下的 `PlainText_YYYYMMDDHHMMSS` 目录中。
* **超链接提取：** 从 Office 文档中提取所有超链接并保存为 `.csv` 文件（包含特定表头和扩展名信息）。这些文件将组织在时间戳命名的 `HyperLinks_YYYYMMDDHHMMSS` 目录中。
* **批量 URL 打开：** 支持从生成的 `.csv` 文件中批量打开超链接。该功能内置了强大的全局去重机制，确保每个唯一的 URL 在所有处理过的文档中仅被打开一次。在打开之前，系统会生成一个包含处理摘要的 HTML 页面，并在默认浏览器中预览。

**核心技术栈：**

* **编程语言：** Python
* **GUI 框架：** Tkinter
* **文档解析库：** `python-docx`, `openpyxl`, `python-pptx`
* **其他依赖：** `csv`, `os`, `sys`, `threading`, `webbrowser`, `datetime`

**项目架构：**

* **GUI 层：** `gui.py` —— 处理用户交互、窗口管理、事件绑定以及进度显示。
* **核心逻辑层：** `core/` 目录包含：
* `extractor.py`：实现 Office 文件解析、文本提取和超链接提取的具体逻辑。
* `url_opener.py`：负责读取 `.csv` 文件、执行 URL 去重以及在浏览器中打开链接。



## 构建与运行

**1. 安装依赖：**

要配置项目环境并安装所有必要的 Python 包，请在终端中运行以下命令：

```bash
pip install -r requirements.txt

```

**2. 运行程序：**

执行主脚本启动 GUI 应用程序：

```bash
python gui.py

```

启动后将打开图形化界面，您可以选择输入和输出目录并开始提取流程。

## 开发规范

* **代码组织：** 项目采用关注点分离（SoC）的设计原则，核心业务逻辑位于 `core/` 目录，用户界面逻辑位于 `gui.py`。
* **注释与文档：** 代码包含详尽的注释（主要为日语），解释了函数、类及特定逻辑块的用途。方法均使用了 Docstrings。
* **错误处理：** 针对文件操作、库导入和数据处理过程，利用 `try-except` 块实现了稳健的错误处理。面向用户的错误通过 Tkinter 的 `messagebox` 提示，详细信息会重定向到控制台/日志窗口（`sys.stderr`）。
* **用户反馈：** GUI 通过进度条 (`ttk.Progressbar`) 提供视觉反馈，并设有专门的日志区域捕捉 `stdout` 和 `stderr` 输出。
* **异步操作：** 对于耗时较长的提取任务，使用 `threading.Thread` 进行异步处理，防止 GUI 界面假死，确保流畅的用户体验。
* **命名约定：** 提取的文件遵循统一的命名方案：
* **文本：** `PlainText_[原文件名]_[原扩展名].md`（例如：`PlainText_Report_docx.md`）
* **超链接：** `Urls_[原文件名]_[原扩展名].csv`（例如：`Urls_Report_docx.csv`）
* 输出目录带有时间戳（例如：`PlainText_20251227103000`）。


* **URL 去重：** `url_opener.py` 模块实现了文件内和全局的 URL 去重，避免重复打开相同的链接。