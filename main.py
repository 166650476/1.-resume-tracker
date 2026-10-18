"""简历批量生成与投递管理系统 —— 中文命令行主程序。

程序采用分层菜单：
    1. 个人基础信息管理；
    2. 简历模板管理（新建、预览、修改、复制、生成/批量生成 Word）；
    3. 投递记录管理（筛选、修改、状态更新、批量导入、txt 备份）；
    4. 数据分析与报表（统计、图表、Excel）；
    5. 数据库备份与恢复。

所有用户输入都会经过长度、类型或格式校验；业务异常会被捕获并写入日志，
因此普通输入错误不会导致程序闪退。
"""

from __future__ import annotations

import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

import analysis
import config
import db
import logger
import resume_generator
import tracker

_UNCHANGED = object()


# ---------------------------------------------------------------------------
# 命令行显示工具
# ---------------------------------------------------------------------------
def _display_width(text: str) -> int:
    """计算中文/英文混合文本在终端中的显示宽度。"""
    width = 0
    for char in text:
        width += 2 if unicodedata.east_asian_width(char) in ("W", "F") else 1
    return width


def _truncate_text(text: str, max_width: int) -> str:
    """按终端显示宽度截断文本，超出部分用省略号表示。"""
    if _display_width(text) <= max_width:
        return text
    if max_width <= 1:
        return "…"
    result = ""
    current_width = 0
    for char in text:
        char_width = 2 if unicodedata.east_asian_width(char) in ("W", "F") else 1
        if current_width + char_width > max_width - 1:
            break
        result += char
        current_width += char_width
    return result + "…"


def _format_cell(value: Any) -> str:
    """把单元格值转换为单行文本。"""
    if value is None:
        return ""
    return str(value).replace("\r\n", " / ").replace("\n", " / ").strip()


def _print_table(
    headers: list[str],
    rows: list[tuple[Any, ...] | list[Any]],
    max_widths: list[int] | None = None,
) -> None:
    """打印自适应宽度的中文终端表格。"""
    if max_widths is None:
        max_widths = [32] * len(headers)
    if len(max_widths) != len(headers):
        max_widths = [32] * len(headers)

    normalized_rows = [
        [_format_cell(value) for value in row]
        for row in rows
    ]
    widths: list[int] = []
    for index, header in enumerate(headers):
        natural_width = _display_width(header)
        for row in normalized_rows:
            if index < len(row):
                natural_width = max(natural_width, _display_width(row[index]))
        widths.append(
            max(_display_width(header), min(natural_width, max_widths[index]))
        )

    def render_row(values: list[str]) -> str:
        """把一行数据按列宽对齐后拼接为终端文本。"""
        cells = []
        for index, value in enumerate(values):
            content = _truncate_text(value, widths[index])
            padding = " " * (widths[index] - _display_width(content))
            cells.append(content + padding)
        return " | ".join(cells)

    print(render_row(headers))
    print("-+-".join("-" * width for width in widths))
    if not normalized_rows:
        print("（暂无记录）")
        return
    for row in normalized_rows:
        padded = row + [""] * (len(headers) - len(row))
        print(render_row(padded[: len(headers)]))


def _pause() -> None:
    """暂停并等待用户按回车，避免输出一闪而过。"""
    try:
        input("\n按回车键继续...")
    except EOFError:
        pass


def _strip_quotes(text: str) -> str:
    """去掉 Windows 拖拽文件时自动添加的引号。"""
    return text.strip().strip('"').strip("'")


# ---------------------------------------------------------------------------
# 输入校验工具
# ---------------------------------------------------------------------------
def _prompt_text(
    prompt: str,
    default: str | None = None,
    required: bool = False,
    max_length: int | None = None,
) -> str:
    """读取单行文本，支持默认值、必填与长度限制。"""
    while True:
        raw = input(prompt).strip()
        if not raw:
            if default is not None:
                return default
            if not required:
                return ""
            print("[提示] 该项不能为空，请重新输入。")
            continue
        if max_length is not None and len(raw) > max_length:
            print(f"[提示] 输入长度不能超过 {max_length} 个字符。")
            continue
        return raw


def _prompt_multiline(
    prompt: str,
    required: bool = False,
    show_tip: bool = True,
) -> str:
    """读取多行文本，单独输入 END 结束。

    以 ``-``、``*``、``•`` 开头的行会在 Word 中生成项目符号。
    """
    if show_tip:
        print("提示：逐行输入，单独输入 END 结束；以 - 开头的行会生成项目符号。")
    print(prompt)
    lines: list[str] = []
    while True:
        line = input("> ").rstrip()
        if line.strip().upper() == "END":
            break
        lines.append(line)

    content = "\n".join(lines).strip()
    if required and not content:
        print("[提示] 内容不能为空，请重新输入。")
        return _prompt_multiline(prompt, required=True, show_tip=False)
    return content


def _prompt_int(
    prompt: str,
    minimum: int = 1,
    allow_blank: bool = False,
    default: int | None = None,
) -> int | None:
    """读取整数，并校验最小值和空值规则。"""
    while True:
        raw = input(prompt).strip()
        if not raw:
            if default is not None:
                return default
            if allow_blank:
                return None
            print("[提示] 请输入整数。")
            continue
        try:
            value = int(raw)
        except ValueError:
            print("[提示] 输入格式错误，请输入整数。")
            continue
        if value < minimum:
            print(f"[提示] 请输入不小于 {minimum} 的整数。")
            continue
        return value


def _prompt_float(
    prompt: str,
    default: float,
    minimum: float,
    maximum: float,
) -> float:
    """读取浮点数，留空时使用默认值。"""
    while True:
        raw = input(prompt).strip()
        if not raw:
            return default
        try:
            value = float(raw)
        except ValueError:
            print("[提示] 输入格式错误，请输入数字。")
            continue
        if not minimum <= value <= maximum:
            print(f"[提示] 数值必须在 {minimum:g} 到 {maximum:g} 之间。")
            continue
        return value


def _prompt_date(
    prompt: str,
    default: str | None = None,
    required: bool = False,
    allow_blank: bool = True,
) -> str:
    """读取 YYYY-MM-DD 日期。"""
    while True:
        raw = input(prompt).strip()
        if not raw:
            if default is not None:
                return default
            if allow_blank and not required:
                return ""
            print("[提示] 日期不能为空。")
            continue
        try:
            parsed = datetime.strptime(raw, config.DATE_FORMAT).date()
        except ValueError:
            print("[提示] 日期格式错误，请使用 YYYY-MM-DD，例如 2026-07-01。")
            continue
        return parsed.strftime(config.DATE_FORMAT)


def _prompt_yes_no(prompt: str, default: bool = False) -> bool:
    """读取是/否确认。"""
    suffix = "Y/n" if default else "y/N"
    while True:
        raw = input(f"{prompt} [{suffix}]: ").strip().lower()
        if not raw:
            return default
        if raw in ("y", "yes", "是"):
            return True
        if raw in ("n", "no", "否"):
            return False
        print("[提示] 请输入 y 或 n。")


def _prompt_status(default: str | None = config.DEFAULT_STATUS) -> str | None:
    """显示内置状态列表并读取选择。

    留空时返回 default；default 为 None 时表示“保持原状态/不筛选”。
    """
    print("可选投递状态：")
    for index, status in enumerate(config.STATUS_OPTIONS, start=1):
        print(f"  {index}. {status}")
    while True:
        raw = input("请选择状态（可输入序号或状态名称，回车使用默认）: ").strip()
        if not raw:
            return default
        if raw.isdigit():
            index = int(raw)
            if 1 <= index <= len(config.STATUS_OPTIONS):
                return config.STATUS_OPTIONS[index - 1]
            print("[提示] 状态序号超出范围。")
            continue
        if raw in config.STATUS_OPTIONS:
            return raw
        print("[提示] 状态无效，请从列表中选择。")


def _run_action(action: Callable[..., Any], *args: Any, **kwargs: Any) -> Any:
    """执行菜单动作并捕获异常，保证返回菜单而不是退出程序。"""
    try:
        return action(*args, **kwargs)
    except KeyboardInterrupt:
        raise
    except Exception as exc:
        logger.log_exception(f"执行操作“{action.__name__}”", exc)
        print(f"\n[错误] {logger.friendly_error(exc)}")
        _pause()
        return None


# ---------------------------------------------------------------------------
# 记录选择与详情展示
# ---------------------------------------------------------------------------
def _print_personal_table(records: list[dict[str, Any]]) -> None:
    """显示个人基础信息列表。"""
    rows = [
        (
            record["id"],
            record["name"],
            record["phone"],
            record["email"],
            record.get("education", ""),
            record.get("school", ""),
            record.get("job_intention", ""),
        )
        for record in records
    ]
    _print_table(
        ["ID", "姓名", "手机号", "邮箱", "学历", "学校", "求职意向"],
        rows,
        max_widths=[5, 10, 14, 24, 8, 18, 18],
    )


def _choose_personal_info_id(allow_blank: bool = False) -> int | None:
    """列出个人信息并让用户选择有效 ID。"""
    records = db.list_personal_infos()
    if not records:
        print("[提示] 暂无个人基础信息，请先到“个人基础信息管理”中新增。")
        return None
    _print_personal_table(records)
    while True:
        record_id = _prompt_int(
            "请输入个人信息 ID（回车取消）: ",
            allow_blank=allow_blank,
        )
        if record_id is None:
            return None
        if db.get_personal_info(record_id) is not None:
            return record_id
        print("[提示] 该 ID 不存在，请重新输入。")


def _print_template_table(records: list[dict[str, Any]]) -> None:
    """显示简历模板列表。"""
    rows = [
        (
            record["id"],
            record["template_name"],
            record.get("personal_name", ""),
            record.get("target_position", ""),
            record.get("font_name", ""),
            record.get("font_size", ""),
            record.get("updated_at", ""),
        )
        for record in records
    ]
    _print_table(
        ["ID", "模板名称", "姓名", "目标岗位", "字体", "字号", "更新时间"],
        rows,
        max_widths=[5, 24, 10, 18, 12, 6, 19],
    )


def _choose_template_id(allow_blank: bool = False) -> int | None:
    """列出简历模板并让用户选择有效 ID。"""
    records = db.list_resume_templates()
    if not records:
        print("[提示] 暂无简历模板，请先新建或复制一个模板。")
        return None
    _print_template_table(records)
    while True:
        record_id = _prompt_int(
            "请输入简历模板 ID（回车取消）: ",
            allow_blank=allow_blank,
        )
        if record_id is None:
            return None
        if db.get_resume_template(record_id) is not None:
            return record_id
        print("[提示] 该模板 ID 不存在，请重新输入。")


def _print_application_table(records: list[dict[str, Any]]) -> None:
    """显示投递记录列表。"""
    rows = [
        (
            record["id"],
            record.get("company_name", ""),
            record.get("position", ""),
            record.get("apply_date", ""),
            record.get("status", ""),
            record.get("template_name") or "未关联",
            record.get("notes", ""),
            record.get("updated_at", ""),
        )
        for record in records
    ]
    _print_table(
        ["ID", "公司", "岗位", "投递日期", "状态", "关联模板", "备注", "更新时间"],
        rows,
        max_widths=[5, 18, 16, 11, 9, 16, 26, 19],
    )


def _choose_application_id(allow_blank: bool = False) -> int | None:
    """列出投递记录并让用户选择有效 ID。"""
    records = db.list_applications()
    if not records:
        print("[提示] 暂无投递记录，请先新增。")
        return None
    _print_application_table(records)
    while True:
        record_id = _prompt_int(
            "请输入投递记录 ID（回车取消）: ",
            allow_blank=allow_blank,
        )
        if record_id is None:
            return None
        if db.get_application(record_id) is not None:
            return record_id
        print("[提示] 该投递记录 ID 不存在，请重新输入。")


def _print_personal_detail(record: dict[str, Any]) -> None:
    """打印个人基础信息详情。"""
    print(config.SUB_SEPARATOR)
    print(f"个人信息 ID：{record['id']}")
    fields = [
        ("姓名", "name"),
        ("手机号", "phone"),
        ("邮箱", "email"),
        ("性别", "gender"),
        ("出生日期", "birth_date"),
        ("所在城市", "city"),
        ("求职意向", "job_intention"),
        ("学历", "education"),
        ("学校", "school"),
        ("专业", "major"),
        ("毕业年份", "graduation_year"),
        ("创建时间", "created_at"),
        ("更新时间", "updated_at"),
    ]
    for label, key in fields:
        print(f"{label}：{record.get(key) or '（未填写）'}")
    print(config.SUB_SEPARATOR)


def _print_application_detail(record: dict[str, Any]) -> None:
    """打印单条投递记录详情。"""
    print(config.SUB_SEPARATOR)
    print(f"投递记录 ID：{record['id']}")
    print(f"公司名称：{record.get('company_name', '')}")
    print(f"岗位名称：{record.get('position', '')}")
    print(f"投递日期：{record.get('apply_date', '')}")
    print(f"投递状态：{record.get('status', '')}")
    print(f"关联模板：{record.get('template_name') or '未关联'}")
    print("备注：")
    print(record.get("notes") or "（无）")
    print(f"创建时间：{record.get('created_at', '')}")
    print(f"更新时间：{record.get('updated_at', '')}")
    print(config.SUB_SEPARATOR)


# ---------------------------------------------------------------------------
# 个人基础信息菜单
# ---------------------------------------------------------------------------
def _add_personal_info() -> None:
    """交互式新增个人基础信息。"""
    print("\n【新增个人基础信息】带 * 的项为必填。")
    data = {
        "name": _prompt_text("* 姓名: ", required=True, max_length=config.NAME_MAX_LEN),
        "phone": _prompt_text("* 手机号（11 位）: ", required=True),
        "email": _prompt_text("* 邮箱: ", required=True, max_length=config.EMAIL_MAX_LEN),
        "gender": _prompt_text(
            f"性别（{'/'.join(config.GENDER_OPTIONS)}，可留空）: "
        ),
        "birth_date": _prompt_date(
            "出生日期（YYYY-MM-DD，可留空）: ",
            required=False,
            allow_blank=True,
        ),
        "city": _prompt_text("所在城市（可留空）: ", max_length=config.CITY_MAX_LEN),
        "job_intention": _prompt_text(
            "求职意向（可留空）: ",
            max_length=config.JOB_INTENTION_MAX_LEN,
        ),
        "education": _prompt_text(
            f"学历（{'/'.join(config.EDUCATION_OPTIONS)}，可留空）: "
        ),
        "school": _prompt_text(
            "学校（可留空）: ", max_length=config.SCHOOL_MAX_LEN
        ),
        "major": _prompt_text(
            "专业（可留空）: ", max_length=config.MAJOR_MAX_LEN
        ),
        "graduation_year": _prompt_text("毕业年份（如 2026，可留空）: "),
    }
    record_id = db.create_personal_info(data)
    print(f"\n[成功] 个人信息已保存，ID={record_id}。")
    _pause()


def _list_personal_infos() -> None:
    """查看全部个人基础信息。"""
    records = db.list_personal_infos()
    print(f"\n【个人基础信息列表】共 {len(records)} 条")
    _print_personal_table(records)
    logger.log_operation("查看个人信息列表", f"数量={len(records)}")
    _pause()


def _show_personal_info_detail() -> None:
    """查看单条个人基础信息详情。"""
    record_id = _choose_personal_info_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_personal_info(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的个人信息不存在。")
    _print_personal_detail(record)
    logger.log_operation("查看个人信息详情", f"ID={record_id}")
    _pause()


def _update_personal_info() -> None:
    """修改个人基础信息，回车保留原值。"""
    record_id = _choose_personal_info_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_personal_info(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的个人信息不存在。")

    _print_personal_detail(record)
    print("修改提示：直接回车保留原值。")
    data = {
        "name": _prompt_text(
            f"姓名 [{record['name']}]: ", default=record["name"], required=True
        ),
        "phone": _prompt_text(
            f"手机号 [{record['phone']}]: ", default=record["phone"], required=True
        ),
        "email": _prompt_text(
            f"邮箱 [{record['email']}]: ", default=record["email"], required=True
        ),
        "gender": _prompt_text(
            f"性别 [{record.get('gender') or '未填写'}]: ",
            default=record.get("gender") or "",
        ),
        "birth_date": _prompt_date(
            f"出生日期 [{record.get('birth_date') or '未填写'}]: ",
            default=record.get("birth_date") or None,
            required=False,
            allow_blank=True,
        ),
        "city": _prompt_text(
            f"所在城市 [{record.get('city') or '未填写'}]: ",
            default=record.get("city") or "",
        ),
        "job_intention": _prompt_text(
            f"求职意向 [{record.get('job_intention') or '未填写'}]: ",
            default=record.get("job_intention") or "",
        ),
        "education": _prompt_text(
            f"学历 [{record.get('education') or '未填写'}]: ",
            default=record.get("education") or "",
        ),
        "school": _prompt_text(
            f"学校 [{record.get('school') or '未填写'}]: ",
            default=record.get("school") or "",
        ),
        "major": _prompt_text(
            f"专业 [{record.get('major') or '未填写'}]: ",
            default=record.get("major") or "",
        ),
        "graduation_year": _prompt_text(
            f"毕业年份 [{record.get('graduation_year') or '未填写'}]: ",
            default=record.get("graduation_year") or "",
        ),
    }
    db.update_personal_info(record_id, data)
    print("\n[成功] 个人信息已更新。")
    _pause()


def _delete_personal_info() -> None:
    """删除个人基础信息。"""
    record_id = _choose_personal_info_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_personal_info(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的个人信息不存在。")
    _print_personal_detail(record)
    if not _prompt_yes_no("确认删除这条个人信息吗？", default=False):
        print("[提示] 已取消删除。")
        return
    db.delete_personal_info(record_id)
    print("\n[成功] 个人信息已删除。")
    _pause()


def personal_info_menu() -> None:
    """个人基础信息管理子菜单。"""
    while True:
        print(config.MENU_SEPARATOR)
        print("个人基础信息管理")
        print(config.SUB_SEPARATOR)
        print("1. 新增个人信息")
        print("2. 查看全部个人信息")
        print("3. 查看个人信息详情")
        print("4. 修改个人信息")
        print("5. 删除个人信息")
        print("0. 返回主菜单")
        choice = input("请选择功能: ").strip()

        if choice == "0":
            return
        if choice == "1":
            _run_action(_add_personal_info)
        elif choice == "2":
            _run_action(_list_personal_infos)
        elif choice == "3":
            _run_action(_show_personal_info_detail)
        elif choice == "4":
            _run_action(_update_personal_info)
        elif choice == "5":
            _run_action(_delete_personal_info)
        else:
            print("[提示] 无效选项，请输入 0-5。")


# ---------------------------------------------------------------------------
# 简历模板菜单
# ---------------------------------------------------------------------------
def _prompt_section_edit(label: str, current: str) -> Any:
    """编辑一个多行区块，返回新内容、空字符串或 _UNCHANGED。"""
    print(f"\n【{label}】当前内容：")
    print(current if current else "（空）")
    while True:
        choice = input("回车保留 / 输入 C 清空 / 输入 E 重新编辑: ").strip().lower()
        if not choice:
            return _UNCHANGED
        if choice == "c":
            return ""
        if choice == "e":
            return _prompt_multiline(
                f"请输入新的{label}内容：",
                show_tip=False,
            )
        print("[提示] 请输入 C、E 或直接回车。")


def _add_resume_template() -> None:
    """交互式新建简历模板。"""
    print("\n【新建简历模板】")
    personal_id = _choose_personal_info_id(allow_blank=True)
    if personal_id is None:
        _pause()
        return

    template_name = _prompt_text(
        "* 模板名称: ",
        required=True,
        max_length=config.TEMPLATE_NAME_MAX_LEN,
    )
    target_position = _prompt_text(
        "* 目标岗位: ",
        required=True,
        max_length=config.TARGET_POSITION_MAX_LEN,
    )

    print(
        "\n接下来填写简历区块。每个区块可留空；"
        "以 - 开头的行会生成项目符号，逐行输入 END 结束。"
    )
    education = _prompt_multiline("教育背景：", show_tip=False)
    internship = _prompt_multiline("实习经历：", show_tip=False)
    project = _prompt_multiline("项目经历：", show_tip=False)
    skills = _prompt_multiline("专业技能：", show_tip=False)
    awards = _prompt_multiline("荣誉奖项：", show_tip=False)
    self_evaluation = _prompt_multiline("自我评价：", show_tip=False)

    print("\n排版设置（直接回车使用默认值）：")
    font_name = _prompt_text(
        f"字体名称 [{config.DEFAULT_FONT_NAME}]: ",
        default=config.DEFAULT_FONT_NAME,
        max_length=config.FONT_NAME_MAX_LEN,
    )
    font_size = _prompt_float(
        f"字号 [{config.DEFAULT_FONT_SIZE}，范围 {config.FONT_SIZE_MIN:g}-{config.FONT_SIZE_MAX:g}]: ",
        default=config.DEFAULT_FONT_SIZE,
        minimum=config.FONT_SIZE_MIN,
        maximum=config.FONT_SIZE_MAX,
    )
    line_spacing = _prompt_float(
        f"行距 [{config.DEFAULT_LINE_SPACING}，范围 {config.LINE_SPACING_MIN:g}-{config.LINE_SPACING_MAX:g}]: ",
        default=config.DEFAULT_LINE_SPACING,
        minimum=config.LINE_SPACING_MIN,
        maximum=config.LINE_SPACING_MAX,
    )
    paragraph_indent = _prompt_float(
        f"首行缩进厘米 [{config.DEFAULT_PARAGRAPH_INDENT_CM}，范围 0-5]: ",
        default=config.DEFAULT_PARAGRAPH_INDENT_CM,
        minimum=config.PARAGRAPH_INDENT_MIN,
        maximum=config.PARAGRAPH_INDENT_MAX,
    )

    data = {
        "template_name": template_name,
        "personal_info_id": personal_id,
        "target_position": target_position,
        "education_experience": education,
        "internship_experience": internship,
        "project_experience": project,
        "skills": skills,
        "awards": awards,
        "self_evaluation": self_evaluation,
        "font_name": font_name,
        "font_size": font_size,
        "line_spacing": line_spacing,
        "paragraph_indent": paragraph_indent,
    }
    template_id = db.create_resume_template(data)
    print(f"\n[成功] 简历模板已创建，ID={template_id}。")
    _pause()


def _list_resume_templates() -> None:
    """查看全部简历模板。"""
    records = db.list_resume_templates()
    print(f"\n【简历模板列表】共 {len(records)} 条")
    _print_template_table(records)
    logger.log_operation("查看简历模板列表", f"数量={len(records)}")
    _pause()


def _preview_resume_template() -> None:
    """在命令行预览简历文本。"""
    template_id = _choose_template_id(allow_blank=True)
    if template_id is None:
        return
    resume_generator.preview_resume(template_id)
    _pause()


def _edit_resume_template() -> None:
    """修改简历模板内容与排版。"""
    template_id = _choose_template_id(allow_blank=True)
    if template_id is None:
        return
    template = db.get_resume_template(template_id)
    if template is None:
        raise LookupError(f"ID 为 {template_id} 的简历模板不存在。")

    print(config.SUB_SEPARATOR)
    print(f"正在修改模板：{template['template_name']}（ID={template_id}）")
    print("标量字段直接回车保留原值，多行区块可保留、清空或重新编辑。")
    print(config.SUB_SEPARATOR)

    data: dict[str, Any] = {
        "template_name": _prompt_text(
            f"模板名称 [{template['template_name']}]: ",
            default=template["template_name"],
            required=True,
            max_length=config.TEMPLATE_NAME_MAX_LEN,
        ),
        "target_position": _prompt_text(
            f"目标岗位 [{template['target_position']}]: ",
            default=template["target_position"],
            required=True,
            max_length=config.TARGET_POSITION_MAX_LEN,
        ),
        "font_name": _prompt_text(
            f"字体名称 [{template['font_name']}]: ",
            default=template["font_name"],
            required=True,
        ),
        "font_size": _prompt_float(
            f"字号 [{template['font_size']}，范围 {config.FONT_SIZE_MIN:g}-{config.FONT_SIZE_MAX:g}]: ",
            default=float(template["font_size"]),
            minimum=config.FONT_SIZE_MIN,
            maximum=config.FONT_SIZE_MAX,
        ),
        "line_spacing": _prompt_float(
            f"行距 [{template['line_spacing']}，范围 {config.LINE_SPACING_MIN:g}-{config.LINE_SPACING_MAX:g}]: ",
            default=float(template["line_spacing"]),
            minimum=config.LINE_SPACING_MIN,
            maximum=config.LINE_SPACING_MAX,
        ),
        "paragraph_indent": _prompt_float(
            f"首行缩进厘米 [{template['paragraph_indent']}，范围 0-5]: ",
            default=float(template["paragraph_indent"]),
            minimum=config.PARAGRAPH_INDENT_MIN,
            maximum=config.PARAGRAPH_INDENT_MAX,
        ),
    }

    for field, label in config.RESUME_SECTION_LABELS:
        new_value = _prompt_section_edit(label, template.get(field) or "")
        if new_value is not _UNCHANGED:
            data[field] = new_value

    db.update_resume_template(template_id, data)
    print("\n[成功] 简历模板已更新。")
    _pause()


def _copy_resume_template() -> None:
    """复制已有模板，快速派生新模板。"""
    source_id = _choose_template_id(allow_blank=True)
    if source_id is None:
        return
    new_name = _prompt_text(
        "新模板名称（直接回车自动生成“原名称（副本）”）: ",
        max_length=config.TEMPLATE_NAME_MAX_LEN,
    )
    new_id, actual_name = resume_generator.copy_template(
        source_id,
        new_name if new_name else None,
    )
    print(f"\n[成功] 模板已复制，新模板 ID={new_id}，名称：{actual_name}")
    print("[提示] 可继续使用“修改模板”调整新模板内容。")
    _pause()


def _generate_resume_document() -> None:
    """为指定模板生成 Word 简历。"""
    template_id = _choose_template_id(allow_blank=True)
    if template_id is None:
        return
    output_path = resume_generator.generate_resume(template_id)
    print(f"\n[成功] Word 简历已生成：\n{output_path}")
    _pause()


def _batch_generate_resume_documents() -> None:
    """批量生成全部模板的 Word 简历。"""
    results = resume_generator.batch_generate_all()
    if not results:
        print("[提示] 暂无可生成的简历模板。")
        _pause()
        return

    success_count = 0
    print("\n【批量生成结果】")
    for result in results:
        if result["path"] is not None:
            success_count += 1
            print(f"[成功] {result['template_name']} -> {result['path']}")
        else:
            print(f"[失败] {result['template_name']}：{result['error']}")
    print(f"\n共 {len(results)} 个模板，成功 {success_count} 个，失败 {len(results) - success_count} 个。")
    _pause()


def _delete_resume_template() -> None:
    """删除简历模板。"""
    template_id = _choose_template_id(allow_blank=True)
    if template_id is None:
        return
    template = db.get_resume_template(template_id)
    if template is None:
        raise LookupError(f"ID 为 {template_id} 的简历模板不存在。")
    print(
        f"\n即将删除模板：{template['template_name']}（ID={template_id}）\n"
        "删除后已生成的 Word 文件不受影响，关联投递记录会保留但取消模板关联。"
    )
    if not _prompt_yes_no("确认删除吗？", default=False):
        print("[提示] 已取消删除。")
        return
    db.delete_resume_template(template_id)
    print("\n[成功] 简历模板已删除。")
    _pause()


def resume_template_menu() -> None:
    """简历模板管理子菜单。"""
    while True:
        print(config.MENU_SEPARATOR)
        print("简历模板管理")
        print(config.SUB_SEPARATOR)
        print("1. 新建简历模板")
        print("2. 查看全部模板")
        print("3. 预览模板文本")
        print("4. 修改模板内容和排版")
        print("5. 复制模板（快速派生）")
        print("6. 生成单份 Word 简历")
        print("7. 批量生成全部 Word 简历")
        print("8. 删除简历模板")
        print("0. 返回主菜单")
        choice = input("请选择功能: ").strip()

        if choice == "0":
            return
        if choice == "1":
            _run_action(_add_resume_template)
        elif choice == "2":
            _run_action(_list_resume_templates)
        elif choice == "3":
            _run_action(_preview_resume_template)
        elif choice == "4":
            _run_action(_edit_resume_template)
        elif choice == "5":
            _run_action(_copy_resume_template)
        elif choice == "6":
            _run_action(_generate_resume_document)
        elif choice == "7":
            _run_action(_batch_generate_resume_documents)
        elif choice == "8":
            _run_action(_delete_resume_template)
        else:
            print("[提示] 无效选项，请输入 0-8。")


# ---------------------------------------------------------------------------
# 投递记录菜单
# ---------------------------------------------------------------------------
def _choose_optional_template_id() -> int | None:
    """新增投递记录时选择关联模板，留空表示不关联。"""
    records = db.list_resume_templates()
    if not records:
        print("[提示] 暂无简历模板，可不关联模板继续。")
        return None
    print("\n可关联的简历模板：")
    _print_template_table(records)
    while True:
        raw = input("关联模板 ID（回车不关联）: ").strip()
        if not raw:
            return None
        if not raw.isdigit():
            print("[提示] 模板 ID 必须是正整数。")
            continue
        template_id = int(raw)
        if db.get_resume_template(template_id) is not None:
            return template_id
        print("[提示] 该模板 ID 不存在。")


def _prompt_template_change(current_id: int | None) -> Any:
    """修改投递记录时调整模板关联，返回新 ID、None 或 _UNCHANGED。"""
    records = db.list_resume_templates()
    if not records:
        print("[提示] 暂无简历模板，当前关联保持不变。")
        return _UNCHANGED
    print("\n可关联的简历模板：")
    _print_template_table(records)
    print(f"当前关联模板：{current_id if current_id else '未关联'}")
    while True:
        raw = input("新模板 ID（回车保留 / 输入 0 取消关联）: ").strip()
        if not raw:
            return _UNCHANGED
        if raw == "0":
            return None
        if not raw.isdigit():
            print("[提示] 模板 ID 必须是正整数或 0。")
            continue
        template_id = int(raw)
        if db.get_resume_template(template_id) is not None:
            return template_id
        print("[提示] 该模板 ID 不存在。")


def _add_application_record() -> None:
    """交互式新增投递记录。"""
    print("\n【新增投递记录】")
    company_name = _prompt_text(
        "* 公司名称: ",
        required=True,
        max_length=config.COMPANY_NAME_MAX_LEN,
    )
    position = _prompt_text(
        "* 岗位名称: ",
        required=True,
        max_length=config.POSITION_NAME_MAX_LEN,
    )
    apply_date = _prompt_date(
        f"投递日期 [{date.today():%Y-%m-%d}]（回车使用今天）: ",
        default=date.today().strftime(config.DATE_FORMAT),
        required=True,
        allow_blank=False,
    )
    print("\n请选择投递状态（回车默认为“待投递”）：")
    status = _prompt_status(default=config.DEFAULT_STATUS)
    if status is None:
        status = config.DEFAULT_STATUS
    template_id = _choose_optional_template_id()
    notes = _prompt_multiline(
        "备注（可记录面试问题、薪资沟通等，可留空）：",
        required=False,
    )

    record_id = tracker.add_application(
        company_name=company_name,
        position=position,
        apply_date=apply_date,
        status=status,
        notes=notes,
        resume_template_id=template_id,
    )
    print(f"\n[成功] 投递记录已新增，ID={record_id}。")
    _pause()


def _query_applications_records() -> None:
    """按多个条件筛选投递记录。"""
    print("\n【筛选投递记录】所有条件均可留空。")
    company = _prompt_text("公司名称（模糊匹配，可留空）: ")
    position = _prompt_text("岗位名称（模糊匹配，可留空）: ")
    print("投递状态筛选：")
    status = _prompt_status(default="")
    start_date = _prompt_date(
        "开始日期（YYYY-MM-DD，可留空）: ",
        required=False,
        allow_blank=True,
    )
    end_date = _prompt_date(
        "结束日期（YYYY-MM-DD，可留空）: ",
        required=False,
        allow_blank=True,
    )

    filters = {
        "company_name": company,
        "position": position,
        "status": status or "",
        "start_date": start_date,
        "end_date": end_date,
    }
    records = tracker.query_applications(filters)
    print(f"\n【查询结果】共 {len(records)} 条")
    _print_application_table(records)
    _pause()


def _show_application_record_detail() -> None:
    """查看单条投递记录完整内容。"""
    record_id = _choose_application_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_application(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的投递记录不存在。")
    _print_application_detail(record)
    logger.log_operation("查看投递记录详情", f"ID={record_id}")
    _pause()


def _update_application_record() -> None:
    """修改投递记录的全部可编辑字段。"""
    record_id = _choose_application_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_application(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的投递记录不存在。")

    print("\n修改提示：文本和日期字段直接回车保留原值。")
    company_name = _prompt_text(
        f"公司名称 [{record['company_name']}]: ",
        default=record["company_name"],
        required=True,
    )
    position = _prompt_text(
        f"岗位名称 [{record['position']}]: ",
        default=record["position"],
        required=True,
    )
    apply_date = _prompt_date(
        f"投递日期 [{record['apply_date']}]: ",
        default=record["apply_date"],
        required=True,
    )
    print(f"当前状态：{record['status']}（回车保持不变）")
    new_status = _prompt_status(default=None)
    new_notes = _prompt_section_edit("备注", record.get("notes") or "")
    template_change = _prompt_template_change(record.get("resume_template_id"))

    data: dict[str, Any] = {
        "company_name": company_name,
        "position": position,
        "apply_date": apply_date,
    }
    if new_status is not None:
        data["status"] = new_status
    if new_notes is not _UNCHANGED:
        data["notes"] = new_notes
    if template_change is not _UNCHANGED:
        data["resume_template_id"] = template_change

    tracker.update_application(record_id, data)
    print("\n[成功] 投递记录已更新。")
    _pause()


def _update_application_status_record() -> None:
    """更新投递状态，并可追加面试问题、薪资沟通等备注。"""
    record_id = _choose_application_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_application(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的投递记录不存在。")

    _print_application_detail(record)
    print("新状态（回车保持不变）：")
    new_status = _prompt_status(default=None)
    note_append = _prompt_multiline(
        "追加备注（可记录面试问题、薪资沟通等；留空则不追加）：",
        required=False,
    )

    if new_status is None and not note_append:
        print("[提示] 未修改状态，也没有追加备注。")
        _pause()
        return

    tracker.update_status(
        record_id,
        new_status or record["status"],
        note_append,
    )
    print("\n[成功] 投递状态/备注已更新。")
    _pause()


def _delete_application_record() -> None:
    """删除投递记录。"""
    record_id = _choose_application_id(allow_blank=True)
    if record_id is None:
        return
    record = db.get_application(record_id)
    if record is None:
        raise LookupError(f"ID 为 {record_id} 的投递记录不存在。")
    _print_application_detail(record)
    if not _prompt_yes_no("确认删除这条投递记录吗？", default=False):
        print("[提示] 已取消删除。")
        return
    tracker.delete_application(record_id)
    print("\n[成功] 投递记录已删除。")
    _pause()


def _batch_import_applications() -> None:
    """从 CSV/Excel 批量导入投递记录。"""
    print("\n【批量导入投递记录】")
    print("支持 .csv / .xlsx / .xls 文件。必需列：公司名称、岗位、投递日期。")
    print("可选列：投递状态、备注、简历模板ID。列名也可使用英文。")
    raw_path = _prompt_text(
        "请输入导入文件完整路径（可将文件拖入终端）: ",
        required=True,
    )
    file_path = Path(_strip_quotes(raw_path)).expanduser()
    summary = tracker.batch_import_applications(file_path)

    print("\n【导入结果】")
    print(f"文件：{summary['file']}")
    print(f"有效数据：{summary['total']} 条")
    print(f"成功导入：{summary['success_count']} 条")
    print(f"重复跳过：{summary['duplicate_count']} 条")
    print(f"校验失败：{summary['failed_count']} 条")
    print(f"空白行跳过：{summary['blank_count']} 条")
    if summary["errors"]:
        print("\n前 10 条错误明细：")
        type_names = {"duplicate": "重复", "validation": "校验失败"}
        for error in summary["errors"][:10]:
            print(
                f"  第 {error['source_row']} 行 "
                f"[{type_names.get(error['type'], error['type'])}] "
                f"{error['message']}"
            )
    _pause()


def _export_applications_txt() -> None:
    """导出全部投递记录为 txt 备份。"""
    output_path = tracker.export_applications_txt()
    if output_path is None:
        print("[提示] 暂无投递记录，未生成 txt 文件。")
    else:
        print(f"\n[成功] 投递记录 txt 备份已生成：\n{output_path}")
    _pause()


def application_menu() -> None:
    """投递记录管理子菜单。"""
    while True:
        print(config.MENU_SEPARATOR)
        print("投递记录管理")
        print(config.SUB_SEPARATOR)
        print("1. 新增投递记录")
        print("2. 筛选查询投递记录")
        print("3. 查看单条投递详情")
        print("4. 修改投递记录")
        print("5. 更新状态 / 追加备注")
        print("6. 删除投递记录")
        print("7. 批量导入投递记录（CSV/Excel）")
        print("8. 导出投递记录 txt 备份")
        print("0. 返回主菜单")
        choice = input("请选择功能: ").strip()

        if choice == "0":
            return
        if choice == "1":
            _run_action(_add_application_record)
        elif choice == "2":
            _run_action(_query_applications_records)
        elif choice == "3":
            _run_action(_show_application_record_detail)
        elif choice == "4":
            _run_action(_update_application_record)
        elif choice == "5":
            _run_action(_update_application_status_record)
        elif choice == "6":
            _run_action(_delete_application_record)
        elif choice == "7":
            _run_action(_batch_import_applications)
        elif choice == "8":
            _run_action(_export_applications_txt)
        else:
            print("[提示] 无效选项，请输入 0-8。")


# ---------------------------------------------------------------------------
# 数据分析菜单
# ---------------------------------------------------------------------------
def _print_statistics_panel(stats: dict[str, Any], issues: list[str]) -> None:
    """打印统计数量、转化率和数据异常提示。"""
    print("\n【求职数据概览】")
    print(f"投递记录总量：{stats['total']}")
    print(f"待投递：{stats['pending']}")
    print(f"已投递（不含待投递）：{stats['submitted']}")
    print(f"笔试到达数量：{stats['written']}")
    print(f"面试到达数量：{stats['interview']}")
    print(f"Offer 数量：{stats['offer']}")
    print(f"拒绝/落选数量：{stats['rejected']}")

    print("\n【阶段转化率】")
    for label, value in stats["conversion_rates"].items():
        print(f"{label}：{value:.2f}%")

    print("\n【数据异常提示】")
    if issues:
        for issue in issues:
            print(f"- {issue}")
    else:
        print("- 未发现明显数据异常。")


def _show_statistics() -> None:
    """查看投递统计概览。"""
    stats = analysis.get_statistics()
    issues = analysis.detect_data_issues()
    _print_statistics_panel(stats, issues)
    _pause()


def _generate_charts_menu() -> None:
    """生成统计图表。"""
    charts = analysis.generate_charts()
    if not charts:
        print("[提示] 暂无投递数据，未生成图表。请先新增投递记录。")
        _pause()
        return
    print("\n[成功] 图表已生成：")
    for label, path in charts.items():
        chart_name = "投递状态分布饼图" if label == "status_chart" else "每月投递数量柱状图"
        print(f"  {chart_name}：{path}")
    _pause()


def _export_excel_menu() -> None:
    """导出 Excel 投递报表。"""
    output_path = analysis.export_applications_excel()
    if output_path is None:
        print("[提示] 暂无投递数据，未生成 Excel 报表。")
    else:
        print(f"\n[成功] Excel 投递报表已生成：\n{output_path}")
        print("包含工作表：投递记录、统计概览、状态分布。")
    _pause()


def _comprehensive_report() -> None:
    """一次性生成统计、图表和 Excel 报表。"""
    stats = analysis.get_statistics()
    issues = analysis.detect_data_issues()
    _print_statistics_panel(stats, issues)
    if stats["total"] == 0:
        print("\n[提示] 暂无数据，无法生成图表和 Excel。")
        _pause()
        return

    charts = analysis.generate_charts()
    if charts:
        print("\n[成功] 图表已生成：")
        for path in charts.values():
            print(f"  {path}")
    else:
        print("\n[提示] 图表未生成。")

    output_path = analysis.export_applications_excel()
    if output_path:
        print(f"\n[成功] Excel 报表已生成：\n{output_path}")
    else:
        print("\n[提示] Excel 报表未生成。")
    _pause()


def analysis_menu() -> None:
    """数据分析与报表子菜单。"""
    while True:
        print(config.MENU_SEPARATOR)
        print("数据分析与报表")
        print(config.SUB_SEPARATOR)
        print("1. 查看投递统计与转化率")
        print("2. 生成统计图表（饼图 + 柱状图）")
        print("3. 导出 Excel 投递报表")
        print("4. 一键生成综合报告（统计 + 图表 + Excel）")
        print("0. 返回主菜单")
        choice = input("请选择功能: ").strip()

        if choice == "0":
            return
        if choice == "1":
            _run_action(_show_statistics)
        elif choice == "2":
            _run_action(_generate_charts_menu)
        elif choice == "3":
            _run_action(_export_excel_menu)
        elif choice == "4":
            _run_action(_comprehensive_report)
        else:
            print("[提示] 无效选项，请输入 0-4。")


# ---------------------------------------------------------------------------
# 备份与恢复菜单
# ---------------------------------------------------------------------------
def _format_file_size(size_bytes: int) -> str:
    """把字节数转换为易读大小。"""
    size = float(size_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


def _show_backup_table(backups: list[Path]) -> None:
    """显示备份文件列表。"""
    rows = []
    for index, path in enumerate(backups, start=1):
        try:
            stat = path.stat()
            modified = datetime.fromtimestamp(stat.st_mtime).strftime(
                config.DATETIME_FORMAT
            )
            size_text = _format_file_size(stat.st_size)
        except OSError:
            modified = "未知"
            size_text = "未知"
        rows.append((index, path.name, size_text, modified))
    _print_table(
        ["序号", "备份文件名", "大小", "修改时间"],
        rows,
        max_widths=[5, 46, 10, 19],
    )


def _backup_database_now() -> None:
    """立即备份数据库到 export 文件夹。"""
    backup_path = db.backup_database()
    print(f"\n[成功] 数据库备份完成：\n{backup_path}")
    _pause()


def _list_backup_files_menu() -> None:
    """查看已有数据库备份。"""
    backups = db.list_backup_files()
    print(f"\n【数据库备份列表】共 {len(backups)} 个")
    if backups:
        _show_backup_table(backups)
    else:
        print("（暂无备份文件）")
    logger.log_operation("查看数据库备份列表", f"数量={len(backups)}")
    _pause()


def _restore_database_interactive() -> None:
    """从备份文件恢复数据库。"""
    backups = db.list_backup_files()
    print("\n【从备份恢复数据库】")
    if not backups:
        print("[提示] export 文件夹中暂无系统备份文件。")
        _pause()
        return
    _show_backup_table(backups)
    raw = _prompt_text(
        "请输入备份序号或 .db 文件完整路径: ",
        required=True,
    )
    raw = _strip_quotes(raw)

    if raw.isdigit():
        index = int(raw)
        if not 1 <= index <= len(backups):
            raise ValueError("备份序号超出范围。")
        backup_path = backups[index - 1]
    else:
        backup_path = Path(raw).expanduser()
        if not backup_path.is_absolute():
            backup_path = (config.PROJECT_ROOT / backup_path).resolve()

    print(f"\n即将使用以下备份覆盖当前数据库：\n{backup_path}")
    print("系统会先自动备份当前数据库，再执行恢复。")
    if not _prompt_yes_no("确认恢复吗？", default=False):
        print("[提示] 已取消恢复。")
        _pause()
        return

    safety_backup = db.restore_database(backup_path)
    counts = db.get_record_counts()
    print("\n[成功] 数据库已恢复。")
    print(f"恢复前安全备份：{safety_backup}")
    print(
        "当前记录数："
        f"个人信息 {counts['personal_info']} 条，"
        f"简历模板 {counts['resume_templates']} 个，"
        f"投递记录 {counts['applications']} 条。"
    )
    _pause()


def backup_menu() -> None:
    """数据库备份与恢复子菜单。"""
    while True:
        print(config.MENU_SEPARATOR)
        print("数据库备份与恢复")
        print(config.SUB_SEPARATOR)
        print("1. 立即备份当前数据库")
        print("2. 查看备份文件列表")
        print("3. 从备份文件恢复数据库")
        print("0. 返回主菜单")
        choice = input("请选择功能: ").strip()

        if choice == "0":
            return
        if choice == "1":
            _run_action(_backup_database_now)
        elif choice == "2":
            _run_action(_list_backup_files_menu)
        elif choice == "3":
            _run_action(_restore_database_interactive)
        else:
            print("[提示] 无效选项，请输入 0-3。")


# ---------------------------------------------------------------------------
# 主菜单与程序入口
# ---------------------------------------------------------------------------
def _print_app_banner(created_database: bool) -> None:
    """显示启动欢迎信息和数据概况。"""
    try:
        counts = db.get_record_counts()
    except Exception as exc:
        logger.log_exception("读取启动数据概况", exc)
        counts = {"personal_info": 0, "resume_templates": 0, "applications": 0}

    print(config.MENU_SEPARATOR)
    print(f"  {config.APP_NAME}  v{config.APP_VERSION}")
    print("  应届生求职 · 本地离线 · 数据不上传互联网")
    print(config.SUB_SEPARATOR)
    print(f"  项目目录：{config.PROJECT_ROOT}")
    print(f"  数据库：  {db.get_database_path()}")
    print(
        "  当前数据："
        f"个人信息 {counts['personal_info']} 条 | "
        f"简历模板 {counts['resume_templates']} 个 | "
        f"投递记录 {counts['applications']} 条"
    )
    if created_database:
        print("  [提示] 未检测到数据库，已自动创建 resume.db。")
    print(f"  日志文件：{logger.get_log_file_path()}")
    print(config.MENU_SEPARATOR)


def main_menu() -> None:
    """主菜单循环。"""
    while True:
        print("\n主菜单")
        print(config.SUB_SEPARATOR)
        print("1. 个人基础信息管理")
        print("2. 简历模板管理 / Word 简历生成")
        print("3. 投递记录管理")
        print("4. 数据分析与报表导出")
        print("5. 数据库备份与恢复")
        print("0. 退出系统")
        choice = input("请选择功能: ").strip()

        if choice == "0":
            return
        try:
            if choice == "1":
                personal_info_menu()
            elif choice == "2":
                resume_template_menu()
            elif choice == "3":
                application_menu()
            elif choice == "4":
                analysis_menu()
            elif choice == "5":
                backup_menu()
            else:
                print("[提示] 无效选项，请输入 0-5。")
        except KeyboardInterrupt:
            print("\n[提示] 当前操作已取消，返回主菜单。")
            logger.log_warning("用户中断当前操作")
        except EOFError:
            raise
        except Exception as exc:
            logger.log_exception("执行主菜单功能", exc)
            print(f"\n[错误] {logger.friendly_error(exc)}")
            _pause()


def main() -> int:
    """程序入口：初始化目录、日志、数据库并进入主菜单。"""
    try:
        config.ensure_directories()
        logger.install_global_exception_handler()
        logger.log_info(
            "程序启动 | 版本=%s | 项目目录=%s",
            config.APP_VERSION,
            config.PROJECT_ROOT,
        )
        created_database = db.init_database()
    except Exception as exc:
        logger.log_exception("系统初始化", exc)
        print(f"\n[启动失败] {logger.friendly_error(exc)}")
        print("请检查项目目录写入权限，并查看 logs 文件夹中的日志。")
        return 1

    _print_app_banner(created_database)
    exit_code = 0
    try:
        main_menu()
    except KeyboardInterrupt:
        print("\n[提示] 收到退出信号，正在安全退出。")
        logger.log_warning("用户通过 Ctrl+C 退出程序")
    except EOFError:
        print("\n[提示] 输入流已结束，程序退出。")
    except Exception as exc:
        exit_code = 1
        logger.log_exception("主程序运行", exc)
        print(f"\n[异常] {logger.friendly_error(exc)}")
        print("详细错误已写入 logs 文件夹。")
    finally:
        logger.log_info("程序退出")
        print(f"\n感谢使用 {config.APP_NAME}，祝你求职顺利！")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())