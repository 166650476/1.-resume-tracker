"""SQLite 数据库访问层。

职责：
    * 自动检测并创建 ``resume.db``；
    * 初始化个人基础信息、简历模板、投递记录三张业务表；
    * 对写入字段做统一校验，阻止手机号、邮箱、日期等脏数据入库；
    * 封装完整 CRUD、筛选查询、批量写入、备份与恢复；
    * 所有写操作使用事务，异常时自动回滚。

本模块不负责命令行展示，只返回字典、列表或记录 ID，便于 UI 层复用。
"""

from __future__ import annotations

import re
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterator, Mapping

import config
import logger


# ---------------------------------------------------------------------------
# 自定义异常
# ---------------------------------------------------------------------------
class DatabaseError(RuntimeError):
    """数据库读写失败。"""


class ValidationError(ValueError):
    """用户输入未通过字段校验。"""


class DuplicateRecordError(ValidationError):
    """检测到重复记录。"""


class RecordNotFoundError(LookupError):
    """指定记录不存在。"""


# 供自动化测试临时切换数据库文件；正常运行时始终使用桌面项目根目录的 resume.db
_DATABASE_PATH_OVERRIDE: Path | None = None


def set_database_path(path: str | Path | None) -> None:
    """设置数据库路径覆盖值（主要供自测脚本使用）。"""
    global _DATABASE_PATH_OVERRIDE
    _DATABASE_PATH_OVERRIDE = Path(path).resolve() if path else None


def get_database_path() -> Path:
    """返回当前生效的数据库文件路径。"""
    return _DATABASE_PATH_OVERRIDE or config.DATABASE_PATH


# ---------------------------------------------------------------------------
# 数据表结构
# ---------------------------------------------------------------------------
_SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS personal_info (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    name                TEXT NOT NULL,
    phone               TEXT NOT NULL,
    email               TEXT NOT NULL,
    gender              TEXT NOT NULL DEFAULT '',
    birth_date          TEXT NOT NULL DEFAULT '',
    city                TEXT NOT NULL DEFAULT '',
    job_intention       TEXT NOT NULL DEFAULT '',
    education           TEXT NOT NULL DEFAULT '',
    school              TEXT NOT NULL DEFAULT '',
    major               TEXT NOT NULL DEFAULT '',
    graduation_year     TEXT NOT NULL DEFAULT '',
    created_at          TEXT NOT NULL,
    updated_at          TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS resume_templates (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    template_name           TEXT NOT NULL,
    personal_info_id        INTEGER NOT NULL,
    target_position         TEXT NOT NULL,
    education_experience    TEXT NOT NULL DEFAULT '',
    internship_experience   TEXT NOT NULL DEFAULT '',
    project_experience      TEXT NOT NULL DEFAULT '',
    skills                  TEXT NOT NULL DEFAULT '',
    awards                  TEXT NOT NULL DEFAULT '',
    self_evaluation         TEXT NOT NULL DEFAULT '',
    font_name               TEXT NOT NULL DEFAULT '微软雅黑',
    font_size               REAL NOT NULL DEFAULT 10.5,
    line_spacing            REAL NOT NULL DEFAULT 1.25,
    paragraph_indent        REAL NOT NULL DEFAULT 0.0,
    created_at              TEXT NOT NULL,
    updated_at              TEXT NOT NULL,
    FOREIGN KEY (personal_info_id) REFERENCES personal_info(id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS applications (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    company_name        TEXT NOT NULL,
    position            TEXT NOT NULL,
    apply_date          TEXT NOT NULL,
    status              TEXT NOT NULL DEFAULT '待投递',
    notes               TEXT NOT NULL DEFAULT '',
    resume_template_id  INTEGER,
    created_at          TEXT NOT NULL,
    updated_at          TEXT NOT NULL,
    FOREIGN KEY (resume_template_id) REFERENCES resume_templates(id) ON DELETE SET NULL
);

CREATE UNIQUE INDEX IF NOT EXISTS idx_personal_phone_unique
    ON personal_info(phone);
CREATE UNIQUE INDEX IF NOT EXISTS idx_template_name_unique
    ON resume_templates(template_name);
CREATE UNIQUE INDEX IF NOT EXISTS idx_application_unique
    ON applications(company_name, position, apply_date);
CREATE INDEX IF NOT EXISTS idx_application_status
    ON applications(status);
CREATE INDEX IF NOT EXISTS idx_application_date
    ON applications(apply_date);
CREATE INDEX IF NOT EXISTS idx_application_company
    ON applications(company_name);
CREATE INDEX IF NOT EXISTS idx_template_personal
    ON resume_templates(personal_info_id);
"""

PERSONAL_FIELDS = (
    "name",
    "phone",
    "email",
    "gender",
    "birth_date",
    "city",
    "job_intention",
    "education",
    "school",
    "major",
    "graduation_year",
)

TEMPLATE_FIELDS = (
    "template_name",
    "personal_info_id",
    "target_position",
    "education_experience",
    "internship_experience",
    "project_experience",
    "skills",
    "awards",
    "self_evaluation",
    "font_name",
    "font_size",
    "line_spacing",
    "paragraph_indent",
)

APPLICATION_FIELDS = (
    "company_name",
    "position",
    "apply_date",
    "status",
    "notes",
    "resume_template_id",
)


# ---------------------------------------------------------------------------
# 连接与事务
# ---------------------------------------------------------------------------
@contextmanager
def get_connection() -> Iterator[sqlite3.Connection]:
    """创建数据库连接并保证事务提交/回滚。

    每次操作使用独立连接，避免命令行程序长期占用数据库文件。
    """
    database_path = get_database_path()
    try:
        database_path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(
            str(database_path),
            timeout=config.DATABASE_TIMEOUT_SECONDS,
        )
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        connection.execute(
            f"PRAGMA busy_timeout = {config.DATABASE_TIMEOUT_SECONDS * 1000}"
        )
    except (OSError, sqlite3.Error) as exc:
        logger.log_exception("打开数据库", exc)
        raise DatabaseError(f"无法打开数据库：{exc}") from exc

    try:
        yield connection
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def init_database() -> bool:
    """检测并初始化数据库。

    如果 resume.db 不存在会自动创建；如果表或索引缺失也会自动补齐。

    Returns:
        True 表示本次创建了新的数据库文件，False 表示使用已有文件。

    Raises:
        DatabaseError: 数据库初始化失败。
    """
    database_path = get_database_path()
    existed = database_path.exists()
    try:
        with get_connection() as connection:
            connection.executescript(_SCHEMA_SQL)
            connection.execute("PRAGMA user_version = 1")
    except sqlite3.Error as exc:
        logger.log_exception("初始化数据库", exc)
        raise DatabaseError(f"数据库初始化失败：{exc}") from exc

    if existed:
        logger.log_info("数据库已存在，完成结构检查：%s", database_path)
    else:
        logger.log_info("数据库不存在，已自动创建：%s", database_path)
    return not existed


# ---------------------------------------------------------------------------
# 通用校验函数
# ---------------------------------------------------------------------------
def _clean_text(
    value: Any,
    field_name: str,
    max_length: int,
    required: bool = False,
    multiline: bool = False,
) -> str:
    """清理并校验文本字段。"""
    if value is None:
        text = ""
    elif isinstance(value, str):
        text = value
    else:
        text = str(value)

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    if not multiline:
        text = " ".join(text.split())
    else:
        text = "\n".join(line.rstrip() for line in text.strip().split("\n"))

    if "\x00" in text:
        raise ValidationError(f"{field_name}包含非法控制字符。")
    if required and not text:
        raise ValidationError(f"{field_name}不能为空。")
    if len(text) > max_length:
        raise ValidationError(f"{field_name}长度不能超过 {max_length} 个字符。")
    return text


def _to_positive_int(
    value: Any,
    field_name: str,
    required: bool = True,
) -> int | None:
    """把用户输入转换为正整数 ID。"""
    if value is None or (isinstance(value, str) and not value.strip()):
        if required:
            raise ValidationError(f"{field_name}不能为空。")
        return None
    if isinstance(value, bool):
        raise ValidationError(f"{field_name}必须是正整数。")
    try:
        number = int(str(value).strip())
    except (TypeError, ValueError) as exc:
        raise ValidationError(f"{field_name}必须是正整数。") from exc
    if number <= 0:
        raise ValidationError(f"{field_name}必须是正整数。")
    return number


def _to_float(
    value: Any,
    field_name: str,
    default: float,
    minimum: float,
    maximum: float,
) -> float:
    """转换并校验浮点配置值。"""
    if value is None or (isinstance(value, str) and not value.strip()):
        return default
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValidationError(f"{field_name}必须是数字。") from exc
    if not minimum <= number <= maximum:
        raise ValidationError(
            f"{field_name}必须在 {minimum:g} 到 {maximum:g} 之间。"
        )
    return round(number, 2)


def _validate_date(
    value: Any,
    field_name: str,
    required: bool = False,
    reject_future: bool = False,
) -> str:
    """校验 YYYY-MM-DD 日期文本。"""
    text = _clean_text(value, field_name, 10, required=required)
    if not text:
        return ""
    try:
        parsed = datetime.strptime(text, config.DATE_FORMAT).date()
    except ValueError as exc:
        raise ValidationError(f"{field_name}格式应为 YYYY-MM-DD，例如 2026-07-01。") from exc
    if reject_future and parsed > date.today():
        raise ValidationError(f"{field_name}不能晚于今天。")
    return parsed.strftime(config.DATE_FORMAT)


def _validate_phone(value: Any) -> str:
    """校验并规范化中国大陆手机号。"""
    text = _clean_text(value, "手机号", config.PHONE_MAX_LEN, required=True)
    text = re.sub(r"[\s-]", "", text)
    if text.startswith("+86"):
        text = text[3:]
    if not re.fullmatch(config.PHONE_PATTERN, text):
        raise ValidationError("手机号格式不正确，请输入 11 位中国大陆手机号。")
    return text


def _validate_email(value: Any) -> str:
    """校验邮箱格式。"""
    text = _clean_text(value, "邮箱", config.EMAIL_MAX_LEN, required=True).lower()
    if not re.fullmatch(config.EMAIL_PATTERN, text):
        raise ValidationError("邮箱格式不正确，例如：name@example.com。")
    return text


def _validate_status(value: Any) -> str:
    """校验投递状态是否在系统内置选项中。"""
    status = _clean_text(value, "投递状态", 20, required=True)
    if not status:
        status = config.DEFAULT_STATUS
    if status not in config.STATUS_OPTIONS:
        raise ValidationError(
            "投递状态无效，可选值：" + "、".join(config.STATUS_OPTIONS)
        )
    return status


def _ensure_personal_exists(connection: sqlite3.Connection, personal_info_id: int) -> None:
    """确认个人基础信息存在。"""
    row = connection.execute(
        "SELECT 1 FROM personal_info WHERE id = ?", (personal_info_id,)
    ).fetchone()
    if row is None:
        raise RecordNotFoundError(f"ID 为 {personal_info_id} 的个人信息不存在。")


def _ensure_template_exists(
    connection: sqlite3.Connection,
    template_id: int | None,
) -> None:
    """确认简历模板存在；None 表示不关联模板，直接通过。"""
    if template_id is None:
        return
    row = connection.execute(
        "SELECT 1 FROM resume_templates WHERE id = ?", (template_id,)
    ).fetchone()
    if row is None:
        raise RecordNotFoundError(f"ID 为 {template_id} 的简历模板不存在。")


# ---------------------------------------------------------------------------
# 字段规范化
# ---------------------------------------------------------------------------
def _normalize_personal(
    data: Mapping[str, Any],
    existing: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """校验并规范化个人基础信息。"""
    if not isinstance(data, Mapping):
        raise ValidationError("个人信息数据格式不正确。")

    base: dict[str, Any] = {field: "" for field in PERSONAL_FIELDS}
    if existing:
        base.update({field: existing.get(field, "") for field in PERSONAL_FIELDS})
    base.update({key: data[key] for key in data if key in PERSONAL_FIELDS})

    base["name"] = _clean_text(
        base["name"], "姓名", config.NAME_MAX_LEN, required=True
    )
    base["phone"] = _validate_phone(base["phone"])
    base["email"] = _validate_email(base["email"])

    gender = _clean_text(base["gender"], "性别", 10)
    if gender and gender not in config.GENDER_OPTIONS:
        raise ValidationError("性别只能是：" + "、".join(config.GENDER_OPTIONS))
    base["gender"] = gender

    birth_date = _validate_date(
        base["birth_date"], "出生日期", required=False, reject_future=True
    )
    if birth_date and int(birth_date[:4]) < config.BIRTH_YEAR_MIN:
        raise ValidationError(f"出生年份不能早于 {config.BIRTH_YEAR_MIN} 年。")
    base["birth_date"] = birth_date

    base["city"] = _clean_text(base["city"], "所在城市", config.CITY_MAX_LEN)
    base["job_intention"] = _clean_text(
        base["job_intention"], "求职意向", config.JOB_INTENTION_MAX_LEN
    )

    education = _clean_text(base["education"], "学历", 20)
    if education and education not in config.EDUCATION_OPTIONS:
        raise ValidationError("学历只能是：" + "、".join(config.EDUCATION_OPTIONS))
    base["education"] = education

    base["school"] = _clean_text(base["school"], "学校", config.SCHOOL_MAX_LEN)
    base["major"] = _clean_text(base["major"], "专业", config.MAJOR_MAX_LEN)

    graduation_year = _clean_text(base["graduation_year"], "毕业年份", 4)
    if graduation_year:
        if not graduation_year.isdigit():
            raise ValidationError("毕业年份必须是 4 位数字，例如 2026。")
        year_number = int(graduation_year)
        if not config.GRADUATION_YEAR_MIN <= year_number <= config.GRADUATION_YEAR_MAX:
            raise ValidationError(
                f"毕业年份必须在 {config.GRADUATION_YEAR_MIN} 到 "
                f"{config.GRADUATION_YEAR_MAX} 之间。"
            )
    base["graduation_year"] = graduation_year
    return base


def _normalize_template(
    data: Mapping[str, Any],
    existing: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """校验并规范化简历模板。"""
    if not isinstance(data, Mapping):
        raise ValidationError("简历模板数据格式不正确。")

    base: dict[str, Any] = {
        "template_name": "",
        "personal_info_id": None,
        "target_position": "",
        "education_experience": "",
        "internship_experience": "",
        "project_experience": "",
        "skills": "",
        "awards": "",
        "self_evaluation": "",
        "font_name": config.DEFAULT_FONT_NAME,
        "font_size": config.DEFAULT_FONT_SIZE,
        "line_spacing": config.DEFAULT_LINE_SPACING,
        "paragraph_indent": config.DEFAULT_PARAGRAPH_INDENT_CM,
    }
    if existing:
        base.update({field: existing.get(field, "") for field in TEMPLATE_FIELDS})
    base.update({key: data[key] for key in data if key in TEMPLATE_FIELDS})

    base["template_name"] = _clean_text(
        base["template_name"],
        "模板名称",
        config.TEMPLATE_NAME_MAX_LEN,
        required=True,
    )
    base["personal_info_id"] = _to_positive_int(
        base["personal_info_id"], "个人信息ID", required=True
    )
    base["target_position"] = _clean_text(
        base["target_position"],
        "目标岗位",
        config.TARGET_POSITION_MAX_LEN,
        required=True,
    )

    for field, label in config.RESUME_SECTION_LABELS:
        base[field] = _clean_text(
            base[field],
            label,
            config.SECTION_CONTENT_MAX_LEN,
            multiline=True,
        )

    base["font_name"] = _clean_text(
        base["font_name"],
        "字体名称",
        config.FONT_NAME_MAX_LEN,
        required=True,
    )
    base["font_size"] = _to_float(
        base["font_size"],
        "字号",
        config.DEFAULT_FONT_SIZE,
        config.FONT_SIZE_MIN,
        config.FONT_SIZE_MAX,
    )
    base["line_spacing"] = _to_float(
        base["line_spacing"],
        "行距",
        config.DEFAULT_LINE_SPACING,
        config.LINE_SPACING_MIN,
        config.LINE_SPACING_MAX,
    )
    base["paragraph_indent"] = _to_float(
        base["paragraph_indent"],
        "首行缩进",
        config.DEFAULT_PARAGRAPH_INDENT_CM,
        config.PARAGRAPH_INDENT_MIN,
        config.PARAGRAPH_INDENT_MAX,
    )
    return base


def _normalize_application(
    data: Mapping[str, Any],
    existing: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """校验并规范化投递记录。"""
    if not isinstance(data, Mapping):
        raise ValidationError("投递记录数据格式不正确。")

    base: dict[str, Any] = {
        "company_name": "",
        "position": "",
        "apply_date": date.today().strftime(config.DATE_FORMAT),
        "status": config.DEFAULT_STATUS,
        "notes": "",
        "resume_template_id": None,
    }
    if existing:
        base.update({field: existing.get(field) for field in APPLICATION_FIELDS})
    base.update({key: data[key] for key in data if key in APPLICATION_FIELDS})

    base["company_name"] = _clean_text(
        base["company_name"],
        "公司名称",
        config.COMPANY_NAME_MAX_LEN,
        required=True,
    )
    base["position"] = _clean_text(
        base["position"],
        "岗位名称",
        config.POSITION_NAME_MAX_LEN,
        required=True,
    )
    base["apply_date"] = _validate_date(
        base["apply_date"], "投递日期", required=True
    )
    base["status"] = _validate_status(base["status"])
    base["notes"] = _clean_text(
        base["notes"],
        "备注",
        config.APPLICATION_NOTES_MAX_LEN,
        multiline=True,
    )
    base["resume_template_id"] = _to_positive_int(
        base["resume_template_id"], "关联简历模板ID", required=False
    )
    return base


def _now_text() -> str:
    """返回当前时间文本，统一写入 created_at / updated_at。"""
    return datetime.now().strftime(config.DATETIME_FORMAT)


# ---------------------------------------------------------------------------
# 个人基础信息 CRUD
# ---------------------------------------------------------------------------
def create_personal_info(data: Mapping[str, Any]) -> int:
    """新增个人基础信息。

    Args:
        data: 包含姓名、手机号、邮箱等字段的字典。

    Returns:
        新记录 ID。

    Raises:
        ValidationError: 字段校验失败。
        DuplicateRecordError: 手机号已存在。
        DatabaseError: 数据库写入失败。
    """
    normalized = _normalize_personal(data)
    now = _now_text()
    try:
        with get_connection() as connection:
            duplicate = connection.execute(
                "SELECT id FROM personal_info WHERE phone = ?",
                (normalized["phone"],),
            ).fetchone()
            if duplicate:
                raise DuplicateRecordError(
                    f"手机号 {normalized['phone']} 已存在，不能重复录入。"
                )
            cursor = connection.execute(
                """
                INSERT INTO personal_info
                    (name, phone, email, gender, birth_date, city, job_intention,
                     education, school, major, graduation_year, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    normalized["name"],
                    normalized["phone"],
                    normalized["email"],
                    normalized["gender"],
                    normalized["birth_date"],
                    normalized["city"],
                    normalized["job_intention"],
                    normalized["education"],
                    normalized["school"],
                    normalized["major"],
                    normalized["graduation_year"],
                    now,
                    now,
                ),
            )
            record_id = int(cursor.lastrowid)
    except (ValidationError, DuplicateRecordError):
        raise
    except sqlite3.IntegrityError as exc:
        logger.log_exception("新增个人基础信息", exc)
        raise DuplicateRecordError("手机号已存在，不能重复录入。") from exc
    except sqlite3.Error as exc:
        logger.log_exception("新增个人基础信息", exc)
        raise DatabaseError(f"新增个人信息失败：{exc}") from exc

    logger.log_db_change("新增", "personal_info", record_id, normalized["name"])
    return record_id


def get_personal_info(personal_info_id: int) -> dict[str, Any] | None:
    """按 ID 查询单条个人基础信息。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT * FROM personal_info WHERE id = ?", (personal_info_id,)
            ).fetchone()
        return dict(row) if row else None
    except sqlite3.Error as exc:
        logger.log_exception("查询个人基础信息", exc)
        raise DatabaseError(f"查询个人信息失败：{exc}") from exc


def list_personal_infos() -> list[dict[str, Any]]:
    """查询全部个人基础信息，最近更新的记录排在前面。"""
    try:
        with get_connection() as connection:
            rows = connection.execute(
                "SELECT * FROM personal_info ORDER BY updated_at DESC, id DESC"
            ).fetchall()
        return [dict(row) for row in rows]
    except sqlite3.Error as exc:
        logger.log_exception("查询个人基础信息列表", exc)
        raise DatabaseError(f"读取个人信息列表失败：{exc}") from exc


def update_personal_info(
    personal_info_id: int,
    data: Mapping[str, Any],
) -> bool:
    """更新个人基础信息，未提供的字段保持原值。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT * FROM personal_info WHERE id = ?", (personal_info_id,)
            ).fetchone()
            if row is None:
                raise RecordNotFoundError(
                    f"ID 为 {personal_info_id} 的个人信息不存在。"
                )
            normalized = _normalize_personal(data, dict(row))
            duplicate = connection.execute(
                "SELECT id FROM personal_info WHERE phone = ? AND id <> ?",
                (normalized["phone"], personal_info_id),
            ).fetchone()
            if duplicate:
                raise DuplicateRecordError(
                    f"手机号 {normalized['phone']} 已被其他个人信息使用。"
                )
            connection.execute(
                """
                UPDATE personal_info
                   SET name = ?, phone = ?, email = ?, gender = ?, birth_date = ?,
                       city = ?, job_intention = ?, education = ?, school = ?,
                       major = ?, graduation_year = ?, updated_at = ?
                 WHERE id = ?
                """,
                (
                    normalized["name"],
                    normalized["phone"],
                    normalized["email"],
                    normalized["gender"],
                    normalized["birth_date"],
                    normalized["city"],
                    normalized["job_intention"],
                    normalized["education"],
                    normalized["school"],
                    normalized["major"],
                    normalized["graduation_year"],
                    _now_text(),
                    personal_info_id,
                ),
            )
    except (ValidationError, DuplicateRecordError, RecordNotFoundError):
        raise
    except sqlite3.IntegrityError as exc:
        logger.log_exception("更新个人基础信息", exc)
        raise DuplicateRecordError("手机号已存在，更新失败。") from exc
    except sqlite3.Error as exc:
        logger.log_exception("更新个人基础信息", exc)
        raise DatabaseError(f"更新个人信息失败：{exc}") from exc

    logger.log_db_change("修改", "personal_info", personal_info_id)
    return True


def delete_personal_info(personal_info_id: int) -> bool:
    """删除个人基础信息。

    若该人员仍有关联的简历模板，则拒绝删除，避免连带丢失模板数据。
    """
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT name FROM personal_info WHERE id = ?", (personal_info_id,)
            ).fetchone()
            if row is None:
                raise RecordNotFoundError(
                    f"ID 为 {personal_info_id} 的个人信息不存在。"
                )
            reference_count = connection.execute(
                "SELECT COUNT(*) AS total FROM resume_templates WHERE personal_info_id = ?",
                (personal_info_id,),
            ).fetchone()["total"]
            if reference_count:
                raise ValidationError(
                    f"该个人信息仍被 {reference_count} 个简历模板引用，"
                    "请先删除或修改相关模板。"
                )
            connection.execute(
                "DELETE FROM personal_info WHERE id = ?", (personal_info_id,)
            )
    except (ValidationError, RecordNotFoundError):
        raise
    except sqlite3.Error as exc:
        logger.log_exception("删除个人基础信息", exc)
        raise DatabaseError(f"删除个人信息失败：{exc}") from exc

    logger.log_db_change("删除", "personal_info", personal_info_id)
    return True


# ---------------------------------------------------------------------------
# 简历模板 CRUD
# ---------------------------------------------------------------------------
def create_resume_template(data: Mapping[str, Any]) -> int:
    """新增简历模板。

    Returns:
        新模板 ID。

    Raises:
        ValidationError: 字段校验失败。
        DuplicateRecordError: 模板名称重复。
        RecordNotFoundError: 关联的个人信息不存在。
        DatabaseError: 数据库写入失败。
    """
    normalized = _normalize_template(data)
    now = _now_text()
    try:
        with get_connection() as connection:
            _ensure_personal_exists(connection, normalized["personal_info_id"])
            duplicate = connection.execute(
                "SELECT id FROM resume_templates WHERE template_name = ?",
                (normalized["template_name"],),
            ).fetchone()
            if duplicate:
                raise DuplicateRecordError(
                    f"模板名称“{normalized['template_name']}”已存在，请换一个名称。"
                )
            cursor = connection.execute(
                """
                INSERT INTO resume_templates
                    (template_name, personal_info_id, target_position,
                     education_experience, internship_experience, project_experience,
                     skills, awards, self_evaluation, font_name, font_size,
                     line_spacing, paragraph_indent, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    normalized["template_name"],
                    normalized["personal_info_id"],
                    normalized["target_position"],
                    normalized["education_experience"],
                    normalized["internship_experience"],
                    normalized["project_experience"],
                    normalized["skills"],
                    normalized["awards"],
                    normalized["self_evaluation"],
                    normalized["font_name"],
                    normalized["font_size"],
                    normalized["line_spacing"],
                    normalized["paragraph_indent"],
                    now,
                    now,
                ),
            )
            record_id = int(cursor.lastrowid)
    except (ValidationError, DuplicateRecordError, RecordNotFoundError):
        raise
    except sqlite3.IntegrityError as exc:
        logger.log_exception("新增简历模板", exc)
        raise DuplicateRecordError("模板名称重复，新增失败。") from exc
    except sqlite3.Error as exc:
        logger.log_exception("新增简历模板", exc)
        raise DatabaseError(f"新增简历模板失败：{exc}") from exc

    logger.log_db_change(
        "新增", "resume_templates", record_id, normalized["template_name"]
    )
    return record_id


def get_resume_template(template_id: int) -> dict[str, Any] | None:
    """按 ID 查询简历模板。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT * FROM resume_templates WHERE id = ?", (template_id,)
            ).fetchone()
        return dict(row) if row else None
    except sqlite3.Error as exc:
        logger.log_exception("查询简历模板", exc)
        raise DatabaseError(f"查询简历模板失败：{exc}") from exc


def list_resume_templates() -> list[dict[str, Any]]:
    """查询全部简历模板，并附带对应人员姓名。"""
    try:
        with get_connection() as connection:
            rows = connection.execute(
                """
                SELECT t.*, p.name AS personal_name
                  FROM resume_templates AS t
                  JOIN personal_info AS p ON p.id = t.personal_info_id
                 ORDER BY t.updated_at DESC, t.id DESC
                """
            ).fetchall()
        return [dict(row) for row in rows]
    except sqlite3.Error as exc:
        logger.log_exception("查询简历模板列表", exc)
        raise DatabaseError(f"读取简历模板列表失败：{exc}") from exc


def template_name_exists(name: str, exclude_id: int | None = None) -> bool:
    """检查模板名称是否已经存在。"""
    try:
        with get_connection() as connection:
            if exclude_id is None:
                row = connection.execute(
                    "SELECT 1 FROM resume_templates WHERE template_name = ?", (name,)
                ).fetchone()
            else:
                row = connection.execute(
                    "SELECT 1 FROM resume_templates WHERE template_name = ? AND id <> ?",
                    (name, exclude_id),
                ).fetchone()
        return row is not None
    except sqlite3.Error as exc:
        logger.log_exception("检查模板名称", exc)
        raise DatabaseError(f"检查模板名称失败：{exc}") from exc


def update_resume_template(template_id: int, data: Mapping[str, Any]) -> bool:
    """更新简历模板，未提供的字段保持原值。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT * FROM resume_templates WHERE id = ?", (template_id,)
            ).fetchone()
            if row is None:
                raise RecordNotFoundError(f"ID 为 {template_id} 的简历模板不存在。")
            normalized = _normalize_template(data, dict(row))
            _ensure_personal_exists(connection, normalized["personal_info_id"])
            duplicate = connection.execute(
                "SELECT id FROM resume_templates WHERE template_name = ? AND id <> ?",
                (normalized["template_name"], template_id),
            ).fetchone()
            if duplicate:
                raise DuplicateRecordError(
                    f"模板名称“{normalized['template_name']}”已存在，请换一个名称。"
                )
            connection.execute(
                """
                UPDATE resume_templates
                   SET template_name = ?, personal_info_id = ?, target_position = ?,
                       education_experience = ?, internship_experience = ?,
                       project_experience = ?, skills = ?, awards = ?,
                       self_evaluation = ?, font_name = ?, font_size = ?,
                       line_spacing = ?, paragraph_indent = ?, updated_at = ?
                 WHERE id = ?
                """,
                (
                    normalized["template_name"],
                    normalized["personal_info_id"],
                    normalized["target_position"],
                    normalized["education_experience"],
                    normalized["internship_experience"],
                    normalized["project_experience"],
                    normalized["skills"],
                    normalized["awards"],
                    normalized["self_evaluation"],
                    normalized["font_name"],
                    normalized["font_size"],
                    normalized["line_spacing"],
                    normalized["paragraph_indent"],
                    _now_text(),
                    template_id,
                ),
            )
    except (ValidationError, DuplicateRecordError, RecordNotFoundError):
        raise
    except sqlite3.IntegrityError as exc:
        logger.log_exception("更新简历模板", exc)
        raise DuplicateRecordError("模板名称重复，更新失败。") from exc
    except sqlite3.Error as exc:
        logger.log_exception("更新简历模板", exc)
        raise DatabaseError(f"更新简历模板失败：{exc}") from exc

    logger.log_db_change("修改", "resume_templates", template_id)
    return True


def delete_resume_template(template_id: int) -> bool:
    """删除简历模板；关联的投递记录会保留，但模板关联字段自动置空。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT template_name FROM resume_templates WHERE id = ?",
                (template_id,),
            ).fetchone()
            if row is None:
                raise RecordNotFoundError(f"ID 为 {template_id} 的简历模板不存在。")
            connection.execute(
                "DELETE FROM resume_templates WHERE id = ?", (template_id,)
            )
    except RecordNotFoundError:
        raise
    except sqlite3.Error as exc:
        logger.log_exception("删除简历模板", exc)
        raise DatabaseError(f"删除简历模板失败：{exc}") from exc

    logger.log_db_change("删除", "resume_templates", template_id, row["template_name"])
    return True


# ---------------------------------------------------------------------------
# 投递记录 CRUD
# ---------------------------------------------------------------------------
def create_application(data: Mapping[str, Any]) -> int:
    """新增投递记录，并检测同公司、同岗位、同日期重复。"""
    normalized = _normalize_application(data)
    now = _now_text()
    try:
        with get_connection() as connection:
            _ensure_template_exists(connection, normalized["resume_template_id"])
            duplicate = connection.execute(
                """
                SELECT id FROM applications
                 WHERE company_name = ? AND position = ? AND apply_date = ?
                """,
                (
                    normalized["company_name"],
                    normalized["position"],
                    normalized["apply_date"],
                ),
            ).fetchone()
            if duplicate:
                raise DuplicateRecordError(
                    "检测到重复投递记录：同一公司、同一岗位、同一投递日期已存在。"
                )
            cursor = connection.execute(
                """
                INSERT INTO applications
                    (company_name, position, apply_date, status, notes,
                     resume_template_id, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    normalized["company_name"],
                    normalized["position"],
                    normalized["apply_date"],
                    normalized["status"],
                    normalized["notes"],
                    normalized["resume_template_id"],
                    now,
                    now,
                ),
            )
            record_id = int(cursor.lastrowid)
    except (ValidationError, RecordNotFoundError):
        raise
    except sqlite3.IntegrityError as exc:
        logger.log_exception("新增投递记录", exc)
        raise DuplicateRecordError(
            "检测到重复投递记录：同一公司、同一岗位、同一投递日期已存在。"
        ) from exc
    except sqlite3.Error as exc:
        logger.log_exception("新增投递记录", exc)
        raise DatabaseError(f"新增投递记录失败：{exc}") from exc

    logger.log_db_change(
        "新增",
        "applications",
        record_id,
        f"{normalized['company_name']} / {normalized['position']}",
    )
    return record_id


def get_application(application_id: int) -> dict[str, Any] | None:
    """按 ID 查询单条投递记录（附带模板名称）。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                """
                SELECT a.*, t.template_name
                  FROM applications AS a
                  LEFT JOIN resume_templates AS t ON t.id = a.resume_template_id
                 WHERE a.id = ?
                """,
                (application_id,),
            ).fetchone()
        return dict(row) if row else None
    except sqlite3.Error as exc:
        logger.log_exception("查询投递记录", exc)
        raise DatabaseError(f"查询投递记录失败：{exc}") from exc


def list_applications(
    filters: Mapping[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """按条件查询投递记录。

    支持的筛选字段：
        company_name: 公司名称模糊匹配；
        position: 岗位名称模糊匹配；
        status: 投递状态精确匹配；
        apply_date: 投递日期精确匹配；
        start_date / end_date: 投递日期闭区间；
        resume_template_id: 关联模板 ID。
    """
    filters = filters or {}
    conditions: list[str] = []
    parameters: list[Any] = []

    company_name = str(filters.get("company_name", "") or "").strip()
    if company_name:
        conditions.append("a.company_name LIKE ?")
        parameters.append(f"%{company_name}%")

    position = str(filters.get("position", "") or "").strip()
    if position:
        conditions.append("a.position LIKE ?")
        parameters.append(f"%{position}%")

    status = str(filters.get("status", "") or "").strip()
    if status:
        status = _validate_status(status)
        conditions.append("a.status = ?")
        parameters.append(status)

    apply_date = _validate_date(filters.get("apply_date", ""), "投递日期")
    if apply_date:
        conditions.append("a.apply_date = ?")
        parameters.append(apply_date)

    start_date = _validate_date(filters.get("start_date", ""), "开始日期")
    if start_date:
        conditions.append("a.apply_date >= ?")
        parameters.append(start_date)

    end_date = _validate_date(filters.get("end_date", ""), "结束日期")
    if end_date:
        conditions.append("a.apply_date <= ?")
        parameters.append(end_date)

    if start_date and end_date and start_date > end_date:
        raise ValidationError("开始日期不能晚于结束日期。")

    template_id = _to_positive_int(
        filters.get("resume_template_id"), "关联简历模板ID", required=False
    )
    if template_id is not None:
        conditions.append("a.resume_template_id = ?")
        parameters.append(template_id)

    where_sql = (" WHERE " + " AND ".join(conditions)) if conditions else ""
    sql = f"""
        SELECT a.*, t.template_name
          FROM applications AS a
          LEFT JOIN resume_templates AS t ON t.id = a.resume_template_id
        {where_sql}
         ORDER BY a.apply_date DESC, a.id DESC
    """
    try:
        with get_connection() as connection:
            rows = connection.execute(sql, parameters).fetchall()
        return [dict(row) for row in rows]
    except sqlite3.Error as exc:
        logger.log_exception("筛选投递记录", exc)
        raise DatabaseError(f"筛选投递记录失败：{exc}") from exc


def update_application(application_id: int, data: Mapping[str, Any]) -> bool:
    """更新投递记录，未提供的字段保持原值。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT * FROM applications WHERE id = ?", (application_id,)
            ).fetchone()
            if row is None:
                raise RecordNotFoundError(f"ID 为 {application_id} 的投递记录不存在。")
            normalized = _normalize_application(data, dict(row))
            _ensure_template_exists(connection, normalized["resume_template_id"])
            duplicate = connection.execute(
                """
                SELECT id FROM applications
                 WHERE company_name = ? AND position = ? AND apply_date = ?
                   AND id <> ?
                """,
                (
                    normalized["company_name"],
                    normalized["position"],
                    normalized["apply_date"],
                    application_id,
                ),
            ).fetchone()
            if duplicate:
                raise DuplicateRecordError(
                    "修改后将与现有记录重复：同一公司、同一岗位、同一投递日期。"
                )
            connection.execute(
                """
                UPDATE applications
                   SET company_name = ?, position = ?, apply_date = ?,
                       status = ?, notes = ?, resume_template_id = ?, updated_at = ?
                 WHERE id = ?
                """,
                (
                    normalized["company_name"],
                    normalized["position"],
                    normalized["apply_date"],
                    normalized["status"],
                    normalized["notes"],
                    normalized["resume_template_id"],
                    _now_text(),
                    application_id,
                ),
            )
    except (ValidationError, RecordNotFoundError):
        raise
    except sqlite3.IntegrityError as exc:
        logger.log_exception("更新投递记录", exc)
        raise DuplicateRecordError("投递记录重复，更新失败。") from exc
    except sqlite3.Error as exc:
        logger.log_exception("更新投递记录", exc)
        raise DatabaseError(f"更新投递记录失败：{exc}") from exc

    logger.log_db_change("修改", "applications", application_id)
    return True


def delete_application(application_id: int) -> bool:
    """按 ID 删除投递记录。"""
    try:
        with get_connection() as connection:
            row = connection.execute(
                "SELECT company_name, position FROM applications WHERE id = ?",
                (application_id,),
            ).fetchone()
            if row is None:
                raise RecordNotFoundError(f"ID 为 {application_id} 的投递记录不存在。")
            connection.execute(
                "DELETE FROM applications WHERE id = ?", (application_id,)
            )
    except RecordNotFoundError:
        raise
    except sqlite3.Error as exc:
        logger.log_exception("删除投递记录", exc)
        raise DatabaseError(f"删除投递记录失败：{exc}") from exc

    logger.log_db_change(
        "删除",
        "applications",
        application_id,
        f"{row['company_name']} / {row['position']}",
    )
    return True


def bulk_create_applications(
    records: list[Mapping[str, Any]],
) -> dict[str, Any]:
    """批量新增投递记录。

    单条记录校验失败或重复时跳过该条，其余合法记录仍在同一事务中写入。
    这样既支持批量导入，也不会因一行脏数据导致全部失败。

    Returns:
        {
            "total": 提交总数,
            "success_count": 成功数,
            "duplicate_count": 重复数,
            "failed_count": 非重复失败数,
            "created_ids": 新记录 ID 列表,
            "errors": [{"row": 序号, "type": 类型, "message": 中文说明}, ...]
        }
    """
    created_ids: list[int] = []
    errors: list[dict[str, Any]] = []
    duplicate_count = 0
    now = _now_text()

    try:
        with get_connection() as connection:
            for index, raw_record in enumerate(records, start=1):
                try:
                    normalized = _normalize_application(raw_record)
                    _ensure_template_exists(
                        connection, normalized["resume_template_id"]
                    )
                    duplicate = connection.execute(
                        """
                        SELECT id FROM applications
                         WHERE company_name = ? AND position = ? AND apply_date = ?
                        """,
                        (
                            normalized["company_name"],
                            normalized["position"],
                            normalized["apply_date"],
                        ),
                    ).fetchone()
                    if duplicate:
                        duplicate_count += 1
                        errors.append(
                            {
                                "row": index,
                                "type": "duplicate",
                                "message": "同一公司、同一岗位、同一日期已存在，已跳过。",
                            }
                        )
                        continue
                    cursor = connection.execute(
                        """
                        INSERT INTO applications
                            (company_name, position, apply_date, status, notes,
                             resume_template_id, created_at, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            normalized["company_name"],
                            normalized["position"],
                            normalized["apply_date"],
                            normalized["status"],
                            normalized["notes"],
                            normalized["resume_template_id"],
                            now,
                            now,
                        ),
                    )
                    created_ids.append(int(cursor.lastrowid))
                except DuplicateRecordError as exc:
                    duplicate_count += 1
                    errors.append(
                        {"row": index, "type": "duplicate", "message": str(exc)}
                    )
                except (ValidationError, RecordNotFoundError) as exc:
                    errors.append(
                        {"row": index, "type": "validation", "message": str(exc)}
                    )
                except sqlite3.IntegrityError as exc:
                    duplicate_count += 1
                    errors.append(
                        {
                            "row": index,
                            "type": "duplicate",
                            "message": f"数据库唯一约束拦截：{exc}",
                        }
                    )
    except sqlite3.Error as exc:
        logger.log_exception("批量新增投递记录", exc)
        raise DatabaseError(f"批量新增投递记录失败：{exc}") from exc

    failed_count = len(errors) - duplicate_count
    logger.log_info(
        "批量写入投递记录 | 总数=%s | 成功=%s | 重复=%s | 失败=%s",
        len(records),
        len(created_ids),
        duplicate_count,
        failed_count,
    )
    return {
        "total": len(records),
        "success_count": len(created_ids),
        "duplicate_count": duplicate_count,
        "failed_count": failed_count,
        "created_ids": created_ids,
        "errors": errors,
    }


# ---------------------------------------------------------------------------
# 数据库备份与恢复
# ---------------------------------------------------------------------------
def _next_available_path(directory: Path, prefix: str, suffix: str = ".db") -> Path:
    """生成不会覆盖已有文件的备份路径。"""
    stamp = datetime.now().strftime(config.FILE_TIMESTAMP_FORMAT)
    candidate = directory / f"{prefix}{stamp}{suffix}"
    serial = 1
    while candidate.exists():
        candidate = directory / f"{prefix}{stamp}_{serial}{suffix}"
        serial += 1
    return candidate


def backup_database(destination_dir: str | Path | None = None) -> Path:
    """使用 SQLite 在线备份 API 备份整个数据库。

    相比直接复制文件，在线备份可以避免数据库正在写入时得到损坏副本。

    Args:
        destination_dir: 备份保存目录，默认 export 文件夹。

    Returns:
        备份文件路径。

    Raises:
        DatabaseError: 备份失败。
    """
    source_path = get_database_path()
    if not source_path.exists():
        init_database()

    target_directory = (
        Path(destination_dir).resolve() if destination_dir else config.EXPORT_DIR
    )
    target_directory.mkdir(parents=True, exist_ok=True)
    backup_path = _next_available_path(
        target_directory, config.BACKUP_FILE_PREFIX
    )

    source_connection: sqlite3.Connection | None = None
    destination_connection: sqlite3.Connection | None = None
    try:
        source_connection = sqlite3.connect(str(source_path))
        destination_connection = sqlite3.connect(str(backup_path))
        source_connection.backup(destination_connection)
        destination_connection.commit()
        destination_connection.close()
        destination_connection = None
        source_connection.close()
        source_connection = None

        check_connection = sqlite3.connect(str(backup_path))
        try:
            result = check_connection.execute("PRAGMA quick_check").fetchone()
        finally:
            check_connection.close()
        if not result or result[0] != "ok":
            raise DatabaseError("备份文件完整性检查未通过。")
    except (sqlite3.Error, OSError, DatabaseError) as exc:
        logger.log_exception("备份数据库", exc)
        try:
            if backup_path.exists():
                backup_path.unlink()
        except OSError:
            pass
        if isinstance(exc, DatabaseError):
            raise
        raise DatabaseError(f"数据库备份失败：{exc}") from exc
    finally:
        if destination_connection is not None:
            destination_connection.close()
        if source_connection is not None:
            source_connection.close()

    logger.log_operation("数据库备份", str(backup_path))
    return backup_path


def list_backup_files(
    backup_dir: str | Path | None = None,
) -> list[Path]:
    """列出指定目录中的数据库备份文件，按时间从新到旧排序。"""
    directory = Path(backup_dir).resolve() if backup_dir else config.EXPORT_DIR
    try:
        if not directory.exists():
            return []
        files = [
            path
            for path in directory.glob(f"{config.BACKUP_FILE_PREFIX}*.db")
            if path.is_file()
        ]
        return sorted(files, key=lambda item: item.stat().st_mtime, reverse=True)
    except OSError as exc:
        logger.log_exception("读取备份文件列表", exc)
        raise DatabaseError(f"读取备份列表失败：{exc}") from exc


def restore_database(backup_path: str | Path) -> Path:
    """从备份文件恢复数据库。

    恢复前会自动把当前数据库再次备份到 export 文件夹，形成安全副本。

    Args:
        backup_path: 已存在的 .db 备份文件路径。

    Returns:
        恢复前自动生成的安全备份路径。

    Raises:
        ValidationError: 备份文件无效或缺少必要数据表。
        DatabaseError: 恢复过程失败。
    """
    source_path = Path(backup_path).expanduser().resolve()
    current_path = get_database_path().resolve()

    if not source_path.exists() or not source_path.is_file():
        raise ValidationError(f"备份文件不存在：{source_path}")
    if source_path == current_path:
        raise ValidationError("不能使用当前正在使用的数据库文件进行恢复。")

    required_tables = {"personal_info", "resume_templates", "applications"}
    uri = f"file:{source_path.as_posix()}?mode=ro"
    try:
        check_connection = sqlite3.connect(uri, uri=True)
        try:
            quick_check = check_connection.execute("PRAGMA quick_check").fetchone()
            table_rows = check_connection.execute(
                "SELECT name FROM sqlite_master WHERE type = 'table'"
            ).fetchall()
        finally:
            check_connection.close()
    except sqlite3.Error as exc:
        logger.log_exception("检查待恢复备份", exc)
        raise ValidationError(
            f"备份文件无法读取或不是有效的 SQLite 数据库：{exc}"
        ) from exc

    if not quick_check or quick_check[0] != "ok":
        raise ValidationError("备份文件完整性检查未通过，已取消恢复。")
    table_names = {row[0] for row in table_rows}
    missing_tables = required_tables - table_names
    if missing_tables:
        raise ValidationError(
            "备份文件缺少必要数据表：" + "、".join(sorted(missing_tables))
        )

    # 先保护现场，再覆盖当前数据库
    safety_backup = backup_database()

    source_connection: sqlite3.Connection | None = None
    destination_connection: sqlite3.Connection | None = None
    try:
        source_connection = sqlite3.connect(uri, uri=True)
        destination_connection = sqlite3.connect(str(current_path))
        source_connection.backup(destination_connection)
        destination_connection.commit()
        destination_connection.close()
        destination_connection = None
        source_connection.close()
        source_connection = None

        init_database()
    except (sqlite3.Error, OSError) as exc:
        logger.log_exception("恢复数据库", exc)
        raise DatabaseError(
            f"数据库恢复失败：{exc}。恢复前安全备份位于：{safety_backup}"
        ) from exc
    finally:
        if destination_connection is not None:
            destination_connection.close()
        if source_connection is not None:
            source_connection.close()

    logger.log_operation(
        "数据库恢复",
        f"来源={source_path} | 恢复前安全备份={safety_backup}",
    )
    return safety_backup


def get_record_counts() -> dict[str, int]:
    """返回三张业务表当前记录数量，用于启动界面和健康检查。"""
    try:
        with get_connection() as connection:
            personal_count = connection.execute(
                "SELECT COUNT(*) AS total FROM personal_info"
            ).fetchone()["total"]
            template_count = connection.execute(
                "SELECT COUNT(*) AS total FROM resume_templates"
            ).fetchone()["total"]
            application_count = connection.execute(
                "SELECT COUNT(*) AS total FROM applications"
            ).fetchone()["total"]
    except sqlite3.Error as exc:
        logger.log_exception("统计数据库记录数", exc)
        raise DatabaseError(f"统计数据库记录数失败：{exc}") from exc
    return {
        "personal_info": int(personal_count),
        "resume_templates": int(template_count),
        "applications": int(application_count),
    }
