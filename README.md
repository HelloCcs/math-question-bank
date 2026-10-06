# 高中数学题库与智能组卷

面向高中数学教师的本地题库工具：导入 Word、PDF 或图片试题，识别并整理题目，维护题库，手动组卷，并导出 Word、PDF 或 Excel。应用默认使用本地 SQLite；AI 分类和 OCR 均可选。

## 功能

- 导入 `.docx`、`.pdf`、`.png`、`.jpg` 和 `.jpeg`。
- 解析题干、选项、答案与解析，并保留题目中的图片资源。
- 按题型、难度和高中数学知识分类；可在导入确认和题库管理中人工修正。
- 可选调用兼容 OpenAI Chat Completions 格式的 AI 服务，对导入题目重新分类。
- 搜索、筛选、维护题库，并按需选择题目、调整顺序和保存组卷历史。
- 导出题目、试卷和答案解析；部分导出格式依赖本机 Office 软件。

题目解析、分类和排版可能受原始文档格式影响。发布前请在自己的试卷上核对导入结果和导出文档，不要将自动识别结果直接视为权威答案。

## 环境要求

- Python 3.11 或更高版本。
- Windows 10/11 是当前主要使用环境；基础 Web 应用基于 Flask。
- Microsoft Word 可用于 Word 预览和 PDF 转换。没有 Word 时，部分预览或 PDF 导出功能不可用。
- Tesseract OCR 及简体中文语言数据用于扫描 PDF/图片的 OCR；纯 Word 导入不依赖 OCR。
- AI 功能需要用户自行提供兼容服务的 API Key；不配置 AI 时仍可使用本地规则分类。

## 安装与启动

在项目目录中运行以下命令：

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
python run.py
```

然后打开 <http://127.0.0.1:5008>。首次启动会在 `data/` 中创建本地数据库和运行数据目录。Windows 用户也可以先运行 `install_windows.bat` 安装依赖，再运行 `start.bat`。

如果 PowerShell 阻止虚拟环境激活脚本，可直接运行：

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run.py
```

## AI 配置

可以在应用的 AI 设置页面填写 API Key，也可以自行编辑被 Git 忽略的 `.env`。`.env.example` 只包含空配置，不含可用凭据。

```dotenv
AI_ENABLED=false
AI_BASE_URL=
AI_API_KEY=
AI_MODEL=
AI_TIMEOUT=20
```

启用 AI 时填写服务商提供的 API Key；若服务商使用 OpenAI 兼容接口，再填写其 Base URL 和模型名。密钥保存在本机 `.env`，不要提交到 GitHub、截图或日志。发给 AI 服务的题目内容会离开本机并交由该服务处理，请先确认试卷内容的授权与隐私要求。

## 测试

```powershell
python -m pip install -r requirements-dev.txt
python -m pytest
```

测试使用临时数据库和项目内的脱敏测试夹具，不应依赖个人题库或 `.env`。

## 项目结构

```text
question_bank/       Flask 应用、业务模块、模板、样式和课程目录数据
tests/               自动化测试
run.py               本地开发启动入口
wsgi.py              WSGI 入口
requirements.txt     运行依赖
```

主要模块：

- `app.py`：页面路由、导入工作流和功能编排。
- `parser.py`：Word、PDF、图片预览内容的题目解析。
- `taxonomy.py`、`curriculum.py`：本地题型、难度和知识分类。
- `ai_classifier.py`：可选 AI 配置、请求、结果校验和分类。
- `storage.py`：SQLite 数据库、持久化及旧结构迁移。
- `papers.py`：组卷数据和导出模型。
- `exporter.py`、`word_preview.py`、`ocr.py`、`assets.py`：导出、预览、OCR 与题目媒体资源处理。

## 数据与隐私

- 题库数据库、上传的原卷、预览文件、待确认导入和导出文件均保存在本机 `data/` 下，并由 `.gitignore` 排除。
- 项目不附带真实客户试卷、题库数据库或 API Key。
- 首次使用前检查 Git 暂存列表，确认没有加入 `.env`、试卷、数据库、日志或个人导出文件。
- 用户上传的试卷可能受版权保护。使用者需自行确认拥有处理和分享这些材料的权利。

## 已知限制

- Word 文档的复杂公式、文本框、分栏、嵌入对象和特殊排版可能无法完全还原。
- 扫描件识别质量依赖图片清晰度、Tesseract 安装和语言包。
- AI 分类依赖外部服务可用性、模型能力和用户配置；结果应由教师复核。
- Word/PDF 最终版式可能因 Office 版本、字体和操作系统而不同。

## 许可证

本项目以 MIT License 发布，版权持有人为 HelloCcs，详见仓库根目录 `LICENSE`。

第三方依赖仍受各自许可证约束。特别是 PyMuPDF 提供 GNU AGPL 或商业许可选项；使用者应根据自己的安装、分发和部署方式核对其许可义务。项目的 MIT 许可不替代第三方依赖许可证。

## 维护与反馈

请通过 GitHub Issues 提交可复现的问题。报告问题时请移除 API Key、学生或学校个人信息，以及未经授权的试卷内容。
