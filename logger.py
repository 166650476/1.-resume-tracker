"""统一日志与全局异常处理模块。

功能：
    1. 同时输出日志到 logs 目录下的文件和控制台（控制台仅显示警告及以上）；
    2. 提供 INFO / WARNING / ERROR 三级日志快捷函数；
    3. 记录用户操作、数据库变更和异常堆栈；
    4. 安装全局异常钩子，避免未捕获异常导致程序静默退出。

所有日志写入失败都不会影响主业务，程序会降级为仅控制台输出。
"""

from __future__ import annotations

import logging
import sqlite3
import sys
import threading
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any

import config

_LOGGER: logging.Logger | None = None


def _build_logger() -> logging.Logger:
    """创建并配置全局 Logger。

    日志文件写入失败时降级为只输出到控制台，避免日志系统本身导致程序崩溃。
    """
    logger = logging.getLogger(config.APP_LOGGER_NAME)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    # 防止模块被重复导入时重复添加 handler
    if logger.handlers:
        return logger

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(name)s | %(module)s:%(lineno)d | %(message)s",
        datefmt=config.DATETIME_FORMAT,
    )

    try:
        config.LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_file = config.LOG_DIR / (
            f"{config.LOG_FILE_PREFIX}{datetime.now():%Y%m%d}{config.LOG_FILE_SUFFIX}"
        )
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=config.LOG_MAX_BYTES,
            backupCount=config.LOG_BACKUP_COUNT,
            encoding="utf-8",
        )
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except (OSError, PermissionError):
        # 日志目录不可写时仍要保证程序可以继续使用
        logger.addHandler(logging.NullHandler())

    try:
        console_handler = logging.StreamHandler(stream=sys.stderr)
        console_handler.setLevel(logging.WARNING)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    except Exception:  # pragma: no cover - 极端环境下流不可用
        logger.addHandler(logging.NullHandler())

    return logger


def get_logger() -> logging.Logger:
    """获取全局 Logger（第一次调用时初始化）。"""
    global _LOGGER
    if _LOGGER is None:
        _LOGGER = _build_logger()
    return _LOGGER


def log_info(message: str, *args: Any) -> None:
    """记录普通信息。"""
    try:
        get_logger().info(message, *args)
    except Exception:
        pass


def log_warning(message: str, *args: Any) -> None:
    """记录警告信息。"""
    try:
        get_logger().warning(message, *args)
    except Exception:
        pass


def log_error(message: str, *args: Any, exc: BaseException | None = None) -> None:
    """记录错误信息。

    Args:
        message: 中文错误描述。
        *args: 传递给 logging 的格式化参数。
        exc: 可选异常对象；提供时会同时记录异常堆栈。
    """
    try:
        if exc is None:
            get_logger().error(message, *args)
        else:
            get_logger().error(
                message,
                *args,
                exc_info=(type(exc), exc, exc.__traceback__),
            )
    except Exception:
        pass


def log_exception(context: str, exc: BaseException) -> None:
    """记录带堆栈的异常信息。

    Args:
        context: 发生异常的业务场景，例如“新增投递记录”。
        exc: 捕获到的异常对象。
    """
    log_error("%s失败：%s", context, exc, exc=exc)


def log_operation(action: str, detail: str = "") -> None:
    """记录用户操作日志。"""
    if detail:
        log_info("用户操作 | %s | %s", action, detail)
    else:
        log_info("用户操作 | %s", action)


def log_db_change(
    operation: str,
    table: str,
    record_id: int | str | None = None,
    detail: str = "",
) -> None:
    """记录数据库增删改日志。"""
    record_text = f"记录ID={record_id}" if record_id is not None else "记录ID=未知"
    if detail:
        log_info("数据库变更 | %s | %s | %s | %s", operation, table, record_text, detail)
    else:
        log_info("数据库变更 | %s | %s | %s", operation, table, record_text)


def friendly_error(exc: BaseException) -> str:
    """把技术异常转换为适合展示给普通用户的中文提示。

    Args:
        exc: 捕获到的异常。

    Returns:
        不包含英文堆栈的简洁中文提示。
    """
    message = str(exc).strip()

    if isinstance(exc, KeyboardInterrupt):
        return "操作已取消。"
    if isinstance(exc, FileNotFoundError):
        return message or "文件不存在，请检查路径是否正确。"
    if isinstance(exc, PermissionError):
        return "没有文件访问权限，请关闭正在占用的文件后重试。"
    if isinstance(exc, UnicodeError):
        return "文件编码无法识别，请将文件另存为 UTF-8 或 GBK 后重试。"
    if isinstance(exc, sqlite3.Error):
        return f"数据库操作失败：{message or '数据库文件可能被占用或损坏。'}"
    if isinstance(exc, (ValueError, RuntimeError, LookupError)):
        return message or "输入或操作有误，请检查后重试。"
    if message:
        return f"操作失败：{message}"
    return "系统发生未知异常，详情已写入 logs 目录中的日志文件。"


def install_global_exception_handler() -> None:
    """安装主线程与子线程的全局异常钩子。

    该钩子只作为最后一道防线。主程序菜单内部还会逐次捕获业务异常，
    因此一般操作失败不会导致程序退出。
    """

    def _global_excepthook(
        exc_type: type[BaseException],
        exc_value: BaseException,
        exc_traceback: Any,
    ) -> None:
        """处理主线程未捕获异常并记录完整堆栈。"""
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        log_error(
            "未捕获的全局异常，程序可能终止",
            exc=exc_value,
        )
        print("\n[系统提示] 程序遇到未预期异常，日志已保存到 logs 目录。")
        print(f"[系统提示] {friendly_error(exc_value)}")

    def _thread_excepthook(args: threading.ExceptHookArgs) -> None:
        """处理子线程未捕获异常。"""
        if args.exc_value is None:
            return
        log_error(
            "子线程未捕获异常 | 线程=%s",
            args.thread.name if args.thread else "未知",
            exc=args.exc_value,
        )

    sys.excepthook = _global_excepthook
    threading.excepthook = _thread_excepthook


def get_log_file_path() -> Path:
    """返回当天日志文件的路径，供界面展示。"""
    return config.LOG_DIR / (
        f"{config.LOG_FILE_PREFIX}{datetime.now():%Y%m%d}{config.LOG_FILE_SUFFIX}"
    )
