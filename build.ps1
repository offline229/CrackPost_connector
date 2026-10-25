# CrackPost_connector 一键打包：生成 dist\download_gui.exe 并同步运行所需文件
# 用法：双击 build.bat，或在项目根目录执行  powershell -File build.ps1

$ErrorActionPreference = "Stop"
$Root = $PSScriptRoot
Set-Location $Root

function Write-Step($msg) {
    Write-Host ""
    Write-Host "==> $msg" -ForegroundColor Cyan
}

Write-Step "项目目录: $Root"

$VenvPy = Join-Path $Root "venv\Scripts\python.exe"
if (-not (Test-Path $VenvPy)) {
    Write-Step "未找到 venv，正在创建虚拟环境…"
    python -m venv (Join-Path $Root "venv")
    if (-not (Test-Path $VenvPy)) {
        throw "无法创建 venv，请确认已安装 Python 并已加入 PATH"
    }
}

$IconIco = Join-Path $Root "icon\CrackPost_v1.ico"
if (-not (Test-Path $IconIco)) {
    throw "缺少图标: $IconIco"
}

Write-Step "安装/更新打包依赖"
& $VenvPy -m pip install -q --upgrade pip
& $VenvPy -m pip install -q -r (Join-Path $Root "requirements-build.txt")

$Dist = Join-Path $Root "dist"
$Work = Join-Path $Root "build\pyinstaller"
$SpecDir = Join-Path $Root "build"
$OutExe = Join-Path $Dist "download_gui.exe"
New-Item -ItemType Directory -Force -Path $Dist, $Work, $SpecDir | Out-Null

function Stop-LockedExe {
    param([string]$ExePath)
    if (-not (Test-Path $ExePath)) { return }

    $full = (Resolve-Path $ExePath).Path
    $procs = Get-CimInstance Win32_Process -Filter "Name = 'download_gui.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.ExecutablePath -and ($_.ExecutablePath -ieq $full) }

    if ($procs) {
        Write-Step "检测到 dist\download_gui.exe 正在运行，正在结束以便覆盖打包"
        foreach ($p in $procs) {
            Write-Host "  结束进程 PID $($p.ProcessId)"
            Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Seconds 1
    }

    # 仍被占用时先改名，再删（杀软偶发锁文件）
    for ($i = 0; $i -lt 8; $i++) {
        try {
            Remove-Item -LiteralPath $ExePath -Force -ErrorAction Stop
            return
        } catch {
            $bak = "$ExePath.old_$PID"
            try {
                Move-Item -LiteralPath $ExePath -Destination $bak -Force -ErrorAction Stop
                Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue
                return
            } catch {
                Start-Sleep -Milliseconds 400
            }
        }
    }

    throw @"
无法覆盖 $ExePath
请先关掉正在运行的「地缝可视化 / download_gui」窗口（任务管理器里也可结束 download_gui.exe），再重新打包。
"@
}

Stop-LockedExe -ExePath $OutExe

Write-Step "PyInstaller 编译 download_gui.exe（单文件、无控制台）"
$PyInstallerArgs = @(
    "-F", "-w",
    "-i", $IconIco,
    "--name", "download_gui",
    "--distpath", $Dist,
    "--specpath", $SpecDir,
    "--workpath", $Work,
    "--add-data", ("{0}{1}icon" -f (Join-Path $Root "icon"), [IO.Path]::PathSeparator),
    "--hidden-import", "babel.numbers",
    "--hidden-import", "imapclient",
    "--hidden-import", "PIL.Image",
    "--hidden-import", "PIL.ImageTk",
    "--hidden-import", "tkcalendar",
    "--hidden-import", "window_icon",
    (Join-Path $Root "download_gui.py")
)
& $VenvPy -m PyInstaller @PyInstallerArgs
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller 失败，退出码 $LASTEXITCODE"
}

Write-Step "同步 dist 运行资源（与 exe 同级）"
Copy-Item -Force (Join-Path $Root "visualization_private.html") (Join-Path $Dist "visualization_private.html")

foreach ($dir in @("font", "icon")) {
    $src = Join-Path $Root $dir
    $dst = Join-Path $Dist $dir
    if (Test-Path $src) {
        if (Test-Path $dst) { Remove-Item -Recurse -Force $dst }
        Copy-Item -Recurse -Force $src $dst
    }
}

$GuideSrc = Join-Path $Root "guide"
if (Test-Path $GuideSrc) {
    $GuideDst = Join-Path $Dist "guide"
    if (Test-Path $GuideDst) { Remove-Item -Recurse -Force $GuideDst }
    Copy-Item -Recurse -Force $GuideSrc $GuideDst
}

Write-Host ""
Write-Host "打包完成。" -ForegroundColor Green
Write-Host "  程序: $Dist\download_gui.exe"
Write-Host "  可视化: $Dist\visualization_private.html"
Write-Host "  将整个 dist 文件夹（或其中上述文件 + 自建的 email 数据）分发给用户即可。"
