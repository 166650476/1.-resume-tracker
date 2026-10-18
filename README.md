# 简历批量生成与投递管理系统

面向应届生求职的本地命令行工具。使用 SQLite 保存个人资料、简历模板和投递记录，使用 python-docx 一键生成规范 Word 简历，使用 pandas/matplotlib 输出投递统计图表与 Excel 报表。

> 所有数据都保存在本机项目目录，不会上传互联网。

## 1. 项目位置

```text
C:\Users\<你的用户名>\Desktop\resume-system\
```

项目根目录内主要文件：

```text
resume-system/
├─ main.py                 # 中文分层菜单主程序
├─ config.py               # 路径、状态、校验规则等全局配置
├─ logger.py               # 日志与全局异常处理
├─ db.py                   # SQLite 建表、CRUD、事务、备份恢复
├─ resume_generator.py     # Word 简历生成、预览、模板复制、批量生成
├─ tracker.py              # 投递记录管理、批量导入、txt 导出
├─ analysis.py             # 统计、转化率、饼图、柱状图、Excel 导出
├─ selftest.py             # 全流程自动化自测脚本
├─ requirements.txt        # Python 依赖
├─ resume.db               # 首次启动自动创建的正式数据库
├─ resume_output/          # Word 简历输出目录
├─ logs/                   # 运行日志目录
└─ export/                 # Excel、图表、txt 与数据库备份
```

## 2. 环境要求

- Windows 10/11
- Python 3.10 及以上（本项目已在 Python 3.14.7 下完成自测）
- 依赖：python-docx、pandas、matplotlib、openpyxl；导入 `.xls` 时还需要 xlrd

## 3. 安装与启动

在 PowerShell 中执行：

```powershell
cd "$env:USERPROFILE\Desktop\resume-system"
python -m pip install -r requirements.txt
python main.py
```

程序启动后按数字选择功能，输入 `0` 返回上一层或退出。

## 4. 首次使用建议流程

1. 进入“个人基础信息管理”，新增姓名、手机号、邮箱、学校、专业等资料。
2. 进入“简历模板管理”，新建模板，逐项填写教育背景、实习经历、项目经历、技能等。
3. 使用“预览模板文本”检查内容；以 `-` 开头的行会生成 Word 项目符号。
4. 使用“复制模板”快速派生针对不同岗位的简历，再“修改模板内容和排版”。
5. 使用“生成单份 Word 简历”或“批量生成全部 Word 简历”，文件输出到 `resume_output`。
6. 进入“投递记录管理”新增投递，使用“更新状态 / 追加备注”记录面试问题和薪资沟通。
7. 进入“数据分析与报表”查看转化率、生成图表和 Excel 报表。
8. 使用“数据库备份与恢复”定期备份 `resume.db`。

## 5. 投递状态

内置状态：

```text
待投递、已投递、笔试、面试、复试、收到offer、拒绝、落选
```

统计口径：

- 笔试/面试/Offer 数量按“当前状态已达到该阶段”计算；
- 拒绝/落选只计入投递总量和已投递量，不计入具体到达阶段；
- 转化率的分母为 0 时显示 0%，不会报错。

## 6. 批量导入格式

支持 `.csv`、`.xlsx`、`.xls`，推荐 UTF-8 编码的 CSV 或 XLSX。

必需列：

| 列名 | 说明 |
| --- | --- |
| 公司名称 | 必填 |
| 岗位 | 必填 |
| 投递日期 | 必填，格式 `YYYY-MM-DD` |

可选列：

| 列名 | 说明 |
| --- | --- |
| 投递状态 | 留空默认为“待投递”，必须是内置状态之一 |
| 备注 | 可记录面试问题、薪资沟通 |
| 简历模板ID | 可留空；填写时必须是已有模板 ID |

也支持英文列名：`company_name`、`position`、`apply_date`、`status`、`notes`、`resume_template_id`。

示例：

```csv
公司名称,岗位,投递日期,投递状态,备注,简历模板ID
示例科技有限公司,Python开发工程师,2026-07-01,已投递,校招官网投递,1
```

批量导入会逐行校验：重复记录和非法行会跳过，并在结果中显示行号与原因。

## 7. 数据备份与恢复

- 备份文件保存在 `export/resume_backup_年月日_时分秒.db`；
- 备份使用 SQLite 在线备份 API，数据库正在写入时也能获得一致副本；
- 恢复前会自动再备份当前数据库，生成安全副本；
- 恢复前会校验备份完整性，并检查三张业务表是否存在。

## 8. 自动化自测

```powershell
cd "$env:USERPROFILE\Desktop\resume-system"
python selftest.py
```

自测使用 `export/_selftest_resume.db` 临时数据库，不会修改正式 `resume.db`。自测覆盖：个人信息录入、重复/格式校验、模板新建/复制/修改/预览、单份和批量 Word、投递新增/更新/重复检测、批量导入、统计、图表、Excel、txt、数据库备份与恢复。测试完成后会自动清理临时数据库及产物。

## 9. 日志与故障排查

- 日志文件：`logs/resume_system_YYYYMMDD.log`；
- 普通输入错误会直接给出中文提示，不会退出程序；
- 未处理异常会写入日志，并显示可读错误信息；
- 如果图表中文显示为方框，请确认 Windows 已安装“微软雅黑”或“黑体”；
- 如果 `.xlsx` 导出失败，请检查文件是否正在 Excel/WPS 中打开；
- 如果数据库文件被占用，请关闭其他正在访问 `resume.db` 的程序后重试。

## 10. 隐私说明

程序只访问本机文件系统和本地 SQLite 数据库，不包含网络请求、不调用云端接口、不上传简历或投递信息。
