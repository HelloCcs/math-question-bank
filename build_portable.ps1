$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot
$python = Join-Path $projectRoot ".venv-dev\Scripts\python.exe"
if (-not (Test-Path $python)) { throw "未找到 .venv-dev，请先创建开发环境。" }
$portable = Join-Path $projectRoot "portable"
& $python -m PyInstaller --noconfirm --clean --onedir --name MathQuestionBank `
  --distpath $portable --workpath (Join-Path $projectRoot "build\portable") --specpath (Join-Path $projectRoot "build") `
  "--add-data=$projectRoot\question_bank\templates:question_bank/templates" `
  "--add-data=$projectRoot\question_bank\static:question_bank/static" `
  "--add-data=$projectRoot\question_bank\data:question_bank/data" `
  --hidden-import win32com.client --hidden-import pythoncom `
  --exclude-module torch --exclude-module torchvision --exclude-module pandas --exclude-module scipy `
  --exclude-module pytest --exclude-module IPython --exclude-module zmq --exclude-module tkinter run.py
Write-Host "程序构建完成：$portable\MathQuestionBank"
Write-Host "请将 tesseract.exe、DLL 和 tessdata/chi_sim.traineddata 放入 portable\tesseract。"
