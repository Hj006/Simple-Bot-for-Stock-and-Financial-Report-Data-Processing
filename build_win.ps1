# build_win.ps1 —— Windows 打包脚本（PowerShell 5.1 兼容）
$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

# 0) 选择 Python：优先 .venv
$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) { $python = "python" }

# 1) 确保 PyInstaller 可用
try {
  & $python -m PyInstaller --version > $null 2>&1
} catch {
  & $python -m pip install --upgrade pip
  & $python -m pip install pyinstaller
}

# 2) 清理旧产物
Remove-Item -Recurse -Force ".\build",".\dist" -ErrorAction SilentlyContinue

# 3) 基本参数
$Name     = "StockSearch"
$IconPath = "assets\app.ico"

# 3.1 获取 streamlit 的 static 目录（确保前端资源被打包）
$staticDir = & $python -c "import pathlib, streamlit; p=(pathlib.Path(streamlit.__file__).parent/'static'); print(p.resolve() if p.exists() else '')"
$staticDataArgs = @()
if ($staticDir -and (Test-Path $staticDir)) {
  $staticDataArgs = @("--add-data", "$staticDir;streamlit\static")
}

# 4) 组装 PyInstaller 参数
$pyiArgs = @(
  "-y","--clean",
  "--name",$Name,
  "--noconfirm",

  # 关键：把 streamlit 的所有资源+元数据打入（避免 3000 端口 & PackageNotFoundError）
  "--collect-all","streamlit",
  "--copy-metadata","streamlit",

  # 常见依赖收集
  "--collect-data","pandas",
  "--collect-submodules","numpy",
  "--collect-submodules","pandas"
) + $staticDataArgs + @(
  # 你的项目资源（新增本地包）
  "--add-data","utils;utils",
  "--add-data","processor;processor",
  "--add-data","streamlit_app.py;.",
  "--add-data","data;data",
  "--add-data","config.py;.",

  # 可选模块排除（未使用 langchain 时可避免警告）
  "--exclude-module","streamlit.external.langchain",
  "--exclude-module","langchain",

  # pandas 隐式模块兜底
  "--hidden-import","pandas._libs.tslibs.timestamps",
  "--hidden-import","pandas._libs.tslibs.np_datetime",
  "--hidden-import","pandas._libs.tslibs.nattype",

  # 若未使用 pyarrow，保持注释（可显著减小体积）
  # "--hidden-import","pyarrow.fs",
  # "--hidden-import","pyarrow._cuda",

  "run_app.py"
)

# 图标（若存在）
if (Test-Path $IconPath) { $pyiArgs = @("--icon",$IconPath) + $pyiArgs }

# 如需隐藏控制台发布给客户，可启用下一行：
# $pyiArgs = @("--noconsole") + $pyiArgs

# 5) 执行打包（用 python -m，避免 PATH 问题）
& $python -m PyInstaller @pyiArgs

# 6) 生成启动脚本与说明（UTF-8 + 切换代码页）
New-Item -ItemType Directory -Force -Path ".\dist" > $null

$bat = @"
@echo off
chcp 65001 >nul
REM 启动本地应用（无需联网）
REM 如需改端口，取消下一行注释并改为 8510 等：
REM set APP_PORT=8501
start "" "%~dp0$Name\$Name.exe"
"@
Set-Content -Path ".\dist\启动应用.bat" -Value $bat -Encoding UTF8

$readme = @"
【StockSearch 本地版（Windows）】

使用步骤：
1. 解压压缩包
2. 双击 “启动应用.bat”
3. 浏览器会自动打开 http://127.0.0.1:8501（如未自动打开请手动输入）
4. 在页面中选择要扫描的本地文件夹并搜索

备注：
- 首次运行若有防火墙提示，请选择“允许”。程序仅监听本机 127.0.0.1。
- 若 8501 端口被占用，可在 “启动应用.bat” 里设置 APP_PORT=新端口。
"@
Set-Content -Path ".\dist\README_使用说明.txt" -Value $readme -Encoding UTF8

Write-Host "打包完成：请查看 dist\ 目录" -ForegroundColor Green
