"""求职投递管理模块。

职责：
    * 新增、查询、修改、删除投递记录；
    * 按公司、岗位、状态、投递日期区间筛选；
    * 更新投递状态并追加面试问题、薪资沟通等备注；
    * 从 CSV / Excel 批量导入投递记录，自动跳过重复与非法行；
    * 将投递记录导出为 UTF-8 文本备份，便于离线查看。

所有字段的最终合法性仍由 db.py 统一校验，本模块负责业务编排与文件读写。
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

import config
import db
import logger


class TrackerError(RuntimeError):
    """投递记录处理失败。"""


def add_application(
    company_name: str,
    position: str,
    apply_date: str | None = None,
    status: str | None = None,
    notes: str = "",
    resume_template_id: int | None = None,
) -> int:
    """新增一条投递记录。

    Args:
        company_name: 公司名称。
        position: 岗位名称。
        apply_date: 投递日期 YYYY-MM-DD，默认今天。
        status: 投递状态，默认“待投递”。
        notes: 备注，可记录面试问题、薪资沟通等信息。
        resume_template_id: 可选关联的简历模板 ID。

    Returns:
        新记录 ID。

    Raises:
        DuplicateRecordError: 同公司、同岗位、同日期记录已存在。
    """
    payload = {
        "company_name": company_name,
        "position": position,
        "apply_date": apply_date or date.today().strftime(config.DATE_FORMAT),
        "status": status or config.DEFAULT_STATUS,
        "notes": notes,
        "resume_template_id": resume_template_id,
    }
    record_id = db.create_application(payload)
    logger.log_operation(
        "新增投递记录",
        f"ID={record_id} | {company_name} / {position}",
    )
    return record_id


def query_applications(
    filters: Mapping[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """按条件查询投递记录，具体筛选字段见 db.list_applications。"""
    records = db.list_applications(filters)
    logger.log_operation("查询投递记录", f"结果数量={len(records)}")
    return records


def update_application(
    application_id: int,
    data: Mapping[str, Any],
) -> bool:
    """修改投递记录，未提供的字段保持原值。"""
    result = db.update_application(application_id, data)
    logger.log_operation("修改投递记录", f"ID={application_id}")
    return result


def update_status(
    application_id: int,
    new_status: str,
    note_append: str = "",
) -> bool:
    """更新投递状态，并可选追加一条带时间标记的备注。

    适合记录“一面问了哪些问题”“HR 沟通薪资”等动态信息。
    """
    record = db.get_application(application_id)
    if record is None:
        raise LookupError(f"ID 为 {application_id} 的投递记录不存在。")

    notes = (record.get("notes") or "").strip()
    extra = (note_append or "").strip()
    if extra:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        new_note = f"[{timestamp}] {extra}"
        notes = f"{notes}\n{new_note}".strip() if notes else new_note

    result = db.update_application(
        application_id,
        {"status": new_status, "notes": notes},
    )
    logger.log_operation(
        "更新投递状态",
        f"ID={application_id} | 状态={new_status} | 追加备注={bool(extra)}",
    )
    return result


def delete_application(application_id: int) -> bool:
    """删除指定投递记录。"""
    result = db.delete_application(application_id)
    logger.log_operation("删除投递记录", f"ID={application_id}")
    return result


# ---------------------------------------------------------------------------
# 批量导入
# ---------------------------------------------------------------------------
def _normalize_column_key(value: Any) -> str:
    """把表头转换为适合匹配的键，兼容 BOM、大小写和空格差异。"""
    text = str(value).replace("\ufeff", "").strip().lower()
    return re.sub(r"\s+", "", text)


def _build_column_mapping(columns: Any) -> dict[str, str]:
    """根据中文/英文别名定位导入文件中的实际列名。"""
    actual_columns = {_normalize_column_key(column): column for column in columns}
    mapping: dict[str, str] = {}
    for field, aliases in config.IMPORT_COLUMN_ALIASES.items():
        for alias in aliases:
            key = _normalize_column_key(alias)
            if key in actual_columns:
                mapping[field] = actual_columns[key]
                break

    missing = [
        field
        for field in ("company_name", "position", "apply_date")
        if field not in mapping
    ]
    if missing:
        display_names = {
            "company_name": "公司名称",
            "position": "岗位",
            "apply_date": "投递日期",
        }
        raise TrackerError(
            "导入文件缺少必需列："
            + "、".join(display_names[field] for field in missing)
        )
    return mapping


def _cell_text(value: Any) -> str:
    """把单元格转换为去掉首尾空白的文本。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).replace("\r\n", "\n").strip()


def _coerce_date_text(value: Any) -> str:
    """把 Excel 日期、斜杠日期等尽量规范为 YYYY-MM-DD。"""
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.strftime("%Y-%m-%d")

    text = _cell_text(value)
    if not text:
        return ""

    # Excel 序列日期（以 1899-12-30 为起点）
    if re.fullmatch(r"\d{5}(?:\.0+)?", text):
        try:
            return (
                datetime(1899, 12, 30) + pd.to_timedelta(float(text), unit="D")
            ).strftime("%Y-%m-%d")
        except (ValueError, TypeError):
            return text

    match = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", text)
    if match:
        try:
            return date(
                int(match.group(1)), int(match.group(2)), int(match.group(3))
            ).strftime("%Y-%m-%d")
        except ValueError:
            return text
    return text


def _coerce_template_id(value: Any) -> str | None:
    """处理 Excel 中常见的 1.0 形式模板 ID。"""
    text = _cell_text(value)
    if not text:
        return None
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _read_import_table(file_path: Path) -> pd.DataFrame:
    """读取 CSV / Excel 导入文件，并在读取失败时给出中文提示。"""
    suffix = file_path.suffix.lower()
    if suffix not in config.SUPPORTED_IMPORT_EXTENSIONS:
        raise TrackerError("仅支持 .csv、.xlsx、.xls 格式的导入文件。")

    try:
        if suffix == ".csv":
            last_error: Exception | None = None
            for encoding in ("utf-8-sig", "gbk", "utf-8"):
                try:
                    return pd.read_csv(
                        file_path,
                        dtype=str,
                        keep_default_na=False,
                        encoding=encoding,
                    )
                except UnicodeDecodeError as exc:
                    last_error = exc
                    continue
            raise TrackerError(
                f"CSV 编码无法识别，请另存为 UTF-8 后重试：{last_error}"
            )

        if suffix == ".xlsx":
            return pd.read_excel(
                file_path,
                dtype=str,
                keep_default_na=False,
                engine="openpyxl",
            )

        # .xls 依赖 xlrd；未安装时 pandas 会抛出 ImportError
        return pd.read_excel(file_path, dtype=str, keep_default_na=False)
    except FileNotFoundError as exc:
        logger.log_exception("读取批量导入文件", exc)
        raise TrackerError(f"导入文件不存在：{file_path}") from exc
    except pd.errors.EmptyDataError as exc:
        logger.log_exception("读取批量导入文件", exc)
        raise TrackerError("导入文件为空，没有可读取的数据。") from exc
    except ImportError as exc:
        logger.log_exception("读取批量导入文件", exc)
        raise TrackerError(
            "读取 .xls 文件需要安装 xlrd，建议将文件另存为 .xlsx 后导入。"
        ) from exc
    except (OSError, ValueError, PermissionError) as exc:
        logger.log_exception("读取批量导入文件", exc)
        raise TrackerError(f"读取导入文件失败：{exc}") from exc


def batch_import_applications(file_path: str | Path) -> dict[str, Any]:
    """从 CSV/Excel 文件批量导入投递记录。

    文件至少需要包含“公司名称、岗位、投递日期”三列；
    “投递状态、备注、简历模板ID”为可选列。重复行和非法行会被跳过。

    Args:
        file_path: 导入文件路径。

    Returns:
        包含总数、成功数、重复数、失败数、空行数和错误明细的字典。
    """
    path = Path(file_path).expanduser().resolve()
    if not path.exists() or not path.is_file():
        raise TrackerError(f"导入文件不存在：{path}")

    dataframe = _read_import_table(path)
    if dataframe.empty:
        raise TrackerError("导入文件没有数据行。")

    mapping = _build_column_mapping(dataframe.columns)
    records: list[dict[str, Any]] = []
    blank_count = 0

    for row_number, (_, row) in enumerate(dataframe.iterrows(), start=2):
        values = {
            field: _cell_text(row[column])
            for field, column in mapping.items()
        }
        if not any(values.values()):
            blank_count += 1
            continue

        record = {
            "company_name": values.get("company_name", ""),
            "position": values.get("position", ""),
            "apply_date": _coerce_date_text(
                row[mapping["apply_date"]] if "apply_date" in mapping else ""
            ),
            "status": values.get("status", "") or config.DEFAULT_STATUS,
            "notes": values.get("notes", ""),
            "resume_template_id": _coerce_template_id(
                row[mapping["resume_template_id"]]
                if "resume_template_id" in mapping
                else ""
            ),
            "_source_row": row_number,
        }
        records.append(record)

    if not records:
        raise TrackerError("导入文件没有有效数据行。")

    result = db.bulk_create_applications(records)
    detailed_errors: list[dict[str, Any]] = []
    for error in result["errors"]:
        index = int(error["row"]) - 1
        source_row = (
            records[index].get("_source_row", index + 2)
            if 0 <= index < len(records)
            else index + 2
        )
        detailed_errors.append(
            {
                "source_row": source_row,
                "type": error["type"],
                "message": error["message"],
            }
        )

    summary = {
        "file": str(path),
        "total": result["total"],
        "success_count": result["success_count"],
        "duplicate_count": result["duplicate_count"],
        "failed_count": result["failed_count"],
        "blank_count": blank_count,
        "errors": detailed_errors,
    }
    logger.log_operation(
        "批量导入投递记录",
        (
            f"文件={path.name} | 有效={summary['total']} | 成功={summary['success_count']} | "
            f"重复={summary['duplicate_count']} | 失败={summary['failed_count']} | "
            f"空行={blank_count}"
        ),
    )
    return summary


# ---------------------------------------------------------------------------
# 文本备份导出
# ---------------------------------------------------------------------------
def _next_text_report_path(output_dir: Path) -> Path:
    """生成不覆盖旧文件的 txt 报告路径。"""
    stamp = datetime.now().strftime(config.FILE_TIMESTAMP_FORMAT)
    candidate = output_dir / f"{config.TXT_REPORT_PREFIX}{stamp}.txt"
    serial = 1
    while candidate.exists():
        candidate = output_dir / f"{config.TXT_REPORT_PREFIX}{stamp}_{serial}.txt"
        serial += 1
    return candidate


def export_applications_txt(
    filters: Mapping[str, Any] | None = None,
    output_dir: str | Path | None = None,
) -> Path | None:
    """把投递记录导出为文本备份。

    Args:
        filters: 与查询一致的筛选条件；None 表示导出全部记录。
        output_dir: 输出目录，默认 export 文件夹。

    Returns:
        导出的 txt 路径；没有数据时返回 None。
    """
    records = db.list_applications(filters)
    if not records:
        logger.log_warning("导出 txt 时没有投递记录")
        return None

    target_directory = (
        Path(output_dir).resolve() if output_dir else config.EXPORT_DIR
    )
    try:
        target_directory.mkdir(parents=True, exist_ok=True)
        output_path = _next_text_report_path(target_directory)
        generated_at = datetime.now().strftime(config.DATETIME_FORMAT)

        lines: list[str] = [
            "=" * 60,
            "简历投递记录备份",
            f"生成时间：{generated_at}",
            f"记录数量：{len(records)}",
            "=" * 60,
            "",
        ]
        for index, record in enumerate(records, start=1):
            lines.extend(
                [
                    f"[{index}] 公司名称：{record.get('company_name', '')}",
                    f"    岗位名称：{record.get('position', '')}",
                    f"    投递日期：{record.get('apply_date', '')}",
                    f"    投递状态：{record.get('status', '')}",
                    f"    关联模板：{record.get('template_name') or '未关联'}",
                ]
            )
            notes = (record.get("notes") or "").strip()
            if notes:
                lines.append("    备注：")
                for note_line in notes.splitlines():
                    lines.append(f"      {note_line}")
            else:
                lines.append("    备注：无")
            lines.append("-" * 60)

        with output_path.open("w", encoding="utf-8-sig", newline="\n") as file:
            file.write("\n".join(lines))
    except (OSError, PermissionError, UnicodeError) as exc:
        logger.log_exception("导出投递记录 txt", exc)
        raise TrackerError(f"导出 txt 备份失败：{exc}") from exc

    logger.log_operation("导出投递记录 txt", f"输出={output_path}")
    return output_path