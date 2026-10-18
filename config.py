"""项目全局配置模块。

本模块集中管理项目路径、文件名、投递状态、字段长度限制、默认排版参数等常量。
其他模块只从这里读取配置，避免路径和魔法数字散落在代码各处。

说明：
    项目默认部署在 Windows 桌面的 resume-system 文件夹中。这里通过
    ``Path(__file__).resolve().parent`` 获取项目根目录，因此只要整个
    resume-system 文件夹位于桌面，所有数据、日志和输出都会保存在其中；
    即使用户整体移动项目目录，项目仍然可以正常工作。
"""

from __future__ import annotations

from pathlib import Path

# ---------------------------------------------------------------------------
# 项目基本信息
# ---------------------------------------------------------------------------
APP_NAME = "简历批量生成与投递管理系统"
APP_VERSION = "1.0.0"
APP_AUTHOR = "应届生求职助手"
APP_LOGGER_NAME = "resume_system"

# 项目根目录：即桌面上的 resume-system 文件夹
PROJECT_ROOT = Path(__file__).resolve().parent
DESKTOP_DIR = PROJECT_ROOT.parent
PROJECT_FOLDER_NAME = "resume-system"

# ---------------------------------------------------------------------------
# 可修改的目录名称
# ---------------------------------------------------------------------------
# 如需修改输出目录名称，只需修改下面三个字符串；程序启动时会自动创建目录。
RESUME_OUTPUT_DIR_NAME = "resume_output"
LOG_DIR_NAME = "logs"
EXPORT_DIR_NAME = "export"

RESUME_OUTPUT_DIR = PROJECT_ROOT / RESUME_OUTPUT_DIR_NAME
LOG_DIR = PROJECT_ROOT / LOG_DIR_NAME
EXPORT_DIR = PROJECT_ROOT / EXPORT_DIR_NAME

# ---------------------------------------------------------------------------
# 数据库配置
# ---------------------------------------------------------------------------
DATABASE_FILENAME = "resume.db"
DATABASE_PATH = PROJECT_ROOT / DATABASE_FILENAME
DATABASE_TIMEOUT_SECONDS = 15

# ---------------------------------------------------------------------------
# 文件命名与日期格式
# ---------------------------------------------------------------------------
DATE_FORMAT = "%Y-%m-%d"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"
FILE_TIMESTAMP_FORMAT = "%Y%m%d_%H%M%S"
BACKUP_FILE_PREFIX = "resume_backup_"
TXT_REPORT_PREFIX = "applications_backup_"
EXCEL_REPORT_PREFIX = "applications_report_"
STATUS_CHART_PREFIX = "status_distribution_"
MONTHLY_CHART_PREFIX = "monthly_applications_"
LOG_FILE_PREFIX = "resume_system_"
LOG_FILE_SUFFIX = ".log"
LOG_MAX_BYTES = 2 * 1024 * 1024
LOG_BACKUP_COUNT = 5

# ---------------------------------------------------------------------------
# 投递状态与阶段
# ---------------------------------------------------------------------------
# 这 8 个状态是系统内置的可选项，新增/修改投递记录时只能从中选择。
STATUS_OPTIONS = (
    "待投递",
    "已投递",
    "笔试",
    "面试",
    "复试",
    "收到offer",
    "拒绝",
    "落选",
)
DEFAULT_STATUS = "待投递"
STATUS_PENDING = "待投递"
STATUS_OFFER = "收到offer"
REJECTED_STATUSES = ("拒绝", "落选")
# 处于以下状态意味着候选人已经进入（或通过）笔试阶段
WRITTEN_REACHED_STATUSES = ("笔试", "面试", "复试", "收到offer")
# 处于以下状态意味着候选人已经进入（或通过）面试阶段
INTERVIEW_REACHED_STATUSES = ("面试", "复试", "收到offer")

# 状态对应的图表颜色，便于饼图稳定显示
STATUS_COLORS = {
    "待投递": "#B0BEC5",
    "已投递": "#42A5F5",
    "笔试": "#26A69A",
    "面试": "#FFA726",
    "复试": "#AB47BC",
    "收到offer": "#66BB6A",
    "拒绝": "#EF5350",
    "落选": "#8D6E63",
}

# ---------------------------------------------------------------------------
# 个人基础信息选项与长度限制
# ---------------------------------------------------------------------------
GENDER_OPTIONS = ("男", "女", "保密", "其他")
EDUCATION_OPTIONS = ("高中", "大专", "本科", "硕士", "博士", "其他")

NAME_MAX_LEN = 50
PHONE_MAX_LEN = 20
EMAIL_MAX_LEN = 120
CITY_MAX_LEN = 50
JOB_INTENTION_MAX_LEN = 100
SCHOOL_MAX_LEN = 100
MAJOR_MAX_LEN = 100
GRADUATION_YEAR_MIN = 1980
GRADUATION_YEAR_MAX = 2100
BIRTH_YEAR_MIN = 1940

# 中国大陆手机号简单校验：1 开头，第二位 3-9，共 11 位
PHONE_PATTERN = r"^1[3-9]\d{9}$"
# 邮箱简单校验，不追求 RFC 全量规则，满足常见求职邮箱即可
EMAIL_PATTERN = r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$"

# ---------------------------------------------------------------------------
# 简历模板配置
# ---------------------------------------------------------------------------
TEMPLATE_NAME_MAX_LEN = 80
TARGET_POSITION_MAX_LEN = 100
SECTION_CONTENT_MAX_LEN = 5000
TEMPLATE_COPY_SUFFIX = "（副本）"
DEFAULT_FONT_NAME = "微软雅黑"
DEFAULT_FONT_SIZE = 10.5
DEFAULT_LINE_SPACING = 1.25
DEFAULT_PARAGRAPH_INDENT_CM = 0.0
FONT_NAME_MAX_LEN = 50
FONT_SIZE_MIN = 8.0
FONT_SIZE_MAX = 24.0
LINE_SPACING_MIN = 1.0
LINE_SPACING_MAX = 3.0
PARAGRAPH_INDENT_MIN = 0.0
PARAGRAPH_INDENT_MAX = 5.0

# 简历各内容区块在预览与 Word 中的顺序
RESUME_SECTION_LABELS = (
    ("education_experience", "教育背景"),
    ("internship_experience", "实习经历"),
    ("project_experience", "项目经历"),
    ("skills", "专业技能"),
    ("awards", "荣誉奖项"),
    ("self_evaluation", "自我评价"),
)

# ---------------------------------------------------------------------------
# 投递记录配置
# ---------------------------------------------------------------------------
COMPANY_NAME_MAX_LEN = 100
POSITION_NAME_MAX_LEN = 100
APPLICATION_NOTES_MAX_LEN = 3000
SUPPORTED_IMPORT_EXTENSIONS = (".csv", ".xlsx", ".xls")

# 批量导入支持的中英文列名别名
IMPORT_COLUMN_ALIASES = {
    "company_name": ("公司名称", "公司", "company_name", "company"),
    "position": ("岗位", "岗位名称", "职位", "position", "job"),
    "apply_date": ("投递日期", "投递时间", "日期", "apply_date", "date"),
    "status": ("投递状态", "状态", "status"),
    "notes": ("备注", "notes", "note"),
    "resume_template_id": ("简历模板ID", "模板ID", "resume_template_id", "template_id"),
}

# ---------------------------------------------------------------------------
# 分析图表配置
# ---------------------------------------------------------------------------
CHART_FONT_CANDIDATES = ("Microsoft YaHei", "SimHei", "SimSun", "Arial Unicode MS")
CHART_DPI = 160
CHART_FIGURE_SIZE = (8.0, 5.5)

# ---------------------------------------------------------------------------
# 命令行显示配置
# ---------------------------------------------------------------------------
MENU_SEPARATOR = "=" * 72
SUB_SEPARATOR = "-" * 72


def ensure_directories() -> dict[str, Path]:
    """创建项目运行所需的全部目录。

    Returns:
        以目录用途为键、Path 对象为值的字典。

    Raises:
        OSError: 当目录因权限不足或磁盘异常而创建失败时抛出。
    """
    directories = {
        "resume_output": RESUME_OUTPUT_DIR,
        "logs": LOG_DIR,
        "export": EXPORT_DIR,
    }
    for directory in directories.values():
        directory.mkdir(parents=True, exist_ok=True)
    return directories
