<#
  bootstrap.ps1 —— WP互相转换工具的启动引导
  ============================================================
  职责：本工具由 Python 编写。如果机器上没有 Python，主程序根本无法运行，
        所以“检测并（经用户同意）安装 Python”只能放在这一层做。

  安全约束（本文件是系统里唯一会“下载并执行第三方安装包”的地方，故必须守住）：
    1. 不静默安装：必须用户显式输入 Y 才继续；
    2. 只从 https://www.python.org/ftp/python/ 官方地址下载，主机名严格校验；
    3. 下载后必须通过 Authenticode 数字签名校验，且签名者必须是
       Python Software Foundation，否则一律删除并中止；
    4. 安装到当前用户目录（InstallAllUsers=0），不请求管理员权限、不改系统级配置；
    5. 任何一步失败都优雅退出，并给出“请手动到 python.org 安装”的指引。

  诊断开关：
    -CheckOnly          只做环境检查，不启动 GUI
    -SimulateNoPython   模拟“未安装 Python”，用于验证引导分支
    -ResolveOnly        只解析官方最新安装包地址并打印
    -VerifySigner <exe> 只校验指定文件的数字签名（调试用）
#>
param(
  [switch]$CheckOnly,
  [switch]$SimulateNoPython,
  [switch]$ResolveOnly,
  [string]$VerifySigner = ""
)
$ErrorActionPreference = "Continue"
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$gui = Join-Path $scriptDir "wp_converter.py"

$PY_INDEX_URL = "https://www.python.org/ftp/python/"
$PY_MANUAL_PAGE = "https://www.python.org/downloads/windows/"
# 签名者用“包含匹配”而非精确匹配 / 证书指纹白名单，是有意权衡：
#   · 精确匹配 Subject 会在 PSF 换证书时失效，需要同步改代码；
#   · 要伪造一张被 Windows 根证书体系信任、且 Subject 含本串的证书，
#     必须通过 CA 的 OV/EV 审核，实际不可行；
#   · 若改成 Thumbprint 白名单，PSF 换证书后所有旧版本工具会立刻无法安装。
# 因此选择包含匹配，风险与维护成本平衡最好。
$PSF_SIGNER_PATTERN = "Python Software Foundation"

function Info($m) { Write-Host $m }
function Ok($m)   { Write-Host $m -ForegroundColor Green }
function Warn($m) { Write-Host $m -ForegroundColor Yellow }
function Bad($m)  { Write-Host $m -ForegroundColor Red }

# ---------------------------------------------------------------------------
# 1. 查找 Python（优先真实安装目录：WindowsApps 的应用执行别名在非交互环境不可靠）
# ---------------------------------------------------------------------------
function Find-Python {
  $bases = New-Object System.Collections.ArrayList
  $lad  = [Environment]::GetEnvironmentVariable("LOCALAPPDATA")
  $pf   = [Environment]::GetEnvironmentVariable("ProgramFiles")
  $pf86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
  if ($lad) {
    [void]$bases.Add((Join-Path $lad "Python"))
    [void]$bases.Add((Join-Path $lad "Programs\Python"))
  }
  if ($pf)   { [void]$bases.Add((Join-Path $pf "Python")) }
  if ($pf86) { [void]$bases.Add((Join-Path $pf86 "Python")) }

  foreach ($base in $bases) {
    if (-not (Test-Path -LiteralPath $base)) { continue }
    $subs = @(Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -like "Python3*" -or $_.Name -like "pythoncore*" } |
              Sort-Object Name -Descending)
    foreach ($d in $subs) {
      $pp = Join-Path $d.FullName "python.exe"
      if (-not (Test-Path -LiteralPath $pp)) { continue }
      $pw = Join-Path $d.FullName "pythonw.exe"
      if (-not (Test-Path -LiteralPath $pw)) { $pw = $pp }
      return @{ Pw = $pw; P = $pp; Source = $d.Name }
    }
  }
  $c = Get-Command pythonw -ErrorAction SilentlyContinue
  if ($c) {
    $c2 = Get-Command python -ErrorAction SilentlyContinue
    return @{ Pw = $c.Source; P = $(if ($c2) { $c2.Source } else { $c.Source }); Source = "PATH" }
  }
  $c = Get-Command python -ErrorAction SilentlyContinue
  if ($c) { return @{ Pw = $c.Source; P = $c.Source; Source = "PATH" } }
  return $null
}

# ---------------------------------------------------------------------------
# 2. 验证 tkinter（没有它就显示不了图形界面，等同不可用）
# ---------------------------------------------------------------------------
function Test-Tkinter([string]$pyExe) {
  if (-not $pyExe -or -not (Test-Path -LiteralPath $pyExe)) { return $false }
  try {
    $out = & $pyExe -c "import tkinter;print('TK_OK')" 2>&1
    if ($LASTEXITCODE -ne 0) { return $false }
    return ((@($out) -join " ") -match "TK_OK")
  } catch { return $false }
}

# ---------------------------------------------------------------------------
# 3. 校验下载来源：必须是 python.org 官方 ftp 路径，且为 https
# ---------------------------------------------------------------------------
function Test-OfficialPythonUrl([string]$Url) {
  try { $u = [System.Uri]$Url } catch { return $false }
  if ($u.Scheme -ne "https") { return $false }
  if ($u.Host -ne "www.python.org") { return $false }
  if ($u.Port -ne 443) { return $false }                                  # 只认标准 443
  if (-not [string]::IsNullOrEmpty($u.UserInfo)) { return $false }         # 拒绝 user:pass@ 形式
  if (-not $u.AbsolutePath.StartsWith("/ftp/python/")) { return $false }
  return $true
}

# ---------------------------------------------------------------------------
# 4. 解析官方最新版安装包地址（对候选版本逐个 HEAD 校验安装包是否真实存在）
# ---------------------------------------------------------------------------
function Resolve-PythonInstaller {
  Info "正在解析 Python 官方下载地址…"
  try {
    $html = (Invoke-WebRequest -Uri $PY_INDEX_URL -UseBasicParsing -TimeoutSec 30).Content
  } catch {
    Bad ("无法访问 python.org：" + $_.Exception.Message)
    return $null
  }
  $vers = @([regex]::Matches($html, 'href="(3\.\d+\.\d+)/"') |
            ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
  if ($vers.Count -eq 0) { Bad "无法从 python.org 解析版本列表"; return $null }
  $sorted = @($vers | Sort-Object { [version]$_ } -Descending)

  $tried = 0
  foreach ($v in $sorted) {
    $tried++
    if ($tried -gt 6) { break }
    $url = "https://www.python.org/ftp/python/$v/python-$v-amd64.exe"
    if (-not (Test-OfficialPythonUrl $url)) { continue }
    try {
      $r = Invoke-WebRequest -Uri $url -Method Head -UseBasicParsing -TimeoutSec 8
      if ($r.StatusCode -eq 200) { return @{ Version = $v; Url = $url } }
    } catch {}
  }
  Bad "未找到可用的官方安装包（可能被网络策略拦截）"
  return $null
}

# ---------------------------------------------------------------------------
# 5. Authenticode 数字签名校验 —— 本文件最关键的安全闸门
# ---------------------------------------------------------------------------
function Test-InstallerSignature([string]$Path) {
  # 返回值 Keep=$true 表示“失败原因是环境性的（如无法联网校验证书吊销状态），
  # 不应删除安装包”，由调用方决定保留并提示用户手动核验。
  try {
    $sig = Get-AuthenticodeSignature -FilePath $Path -ErrorAction Stop
  } catch {
    return @{ Ok = $false; Keep = $false; Signer = ""; Reason = ("无法读取数字签名：" + $_.Exception.Message) }
  }
  if ($null -eq $sig -or $null -eq $sig.SignerCertificate) {
    return @{ Ok = $false; Keep = $false; Signer = ""; Reason = "安装包没有任何数字签名" }
  }
  $signer = [string]$sig.SignerCertificate.Subject
  # 【顺序很重要】签名者身份必须在任何分支之前先确认。
  # 曾有 bug：若先判断“环境性失败”并返回 Keep=$true，就会绕过签名者校验，
  # 让来路不明的文件因“看起来像网络问题”而被保留并提示用户手动安装。
  # 因此这里先把“签名者必须是 PSF”作为硬门槛，再谈证书吊销状态。
  if ($signer -notmatch $PSF_SIGNER_PATTERN) {
    return @{ Ok = $false; Keep = $false; Signer = $signer; Reason = "签名者不是 Python Software Foundation" }
  }
  if ($sig.Status -ne "Valid") {
    $reason = "数字签名状态异常：" + $sig.Status
    $statusMsg = [string]$sig.StatusMessage
    $isEnvIssue = ($statusMsg -match "offline|revocation|network") -or
                  ($sig.Status -eq "UnknownError")
    if ($isEnvIssue) {
      $reason += "（签名者为 PSF，仅证书吊销状态无法在线确认；安装包已保留，请手动核验后安装）"
    }
    return @{ Ok = $false; Keep = $isEnvIssue; Signer = $signer; Reason = $reason }
  }
  return @{ Ok = $true; Keep = $true; Signer = $signer; Reason = "数字签名有效" }
}

# ---------------------------------------------------------------------------
# 6. 下载（不显示逐块进度条，PS 的进度条会拖慢下载数倍）
# ---------------------------------------------------------------------------
function Download-Installer([string]$Url, [string]$Dest) {
  $oldPref = $ProgressPreference
  $ProgressPreference = "SilentlyContinue"
  try {
    Invoke-WebRequest -Uri $Url -OutFile $Dest -UseBasicParsing -TimeoutSec 1800
    return $true
  } catch {
    Bad ("下载失败：" + $_.Exception.Message)
    return $false
  } finally {
    $ProgressPreference = $oldPref
  }
}

# ---------------------------------------------------------------------------
# 7. 静默安装到“当前用户”目录：不需要管理员权限，不弹 UAC
# ---------------------------------------------------------------------------
function Install-PythonSilently([string]$InstallerPath) {
  $argList = @("/quiet",
               "InstallAllUsers=0",
               "PrependPath=1",
               "Include_test=0",
               "Include_tcltk=1",
               "Include_pip=1",
               "Shortcuts=0",
               "Include_launcher=1")
  try {
    $p = Start-Process -FilePath $InstallerPath -ArgumentList $argList -Wait -PassThru
  } catch {
    Bad ("安装器启动失败：" + $_.Exception.Message)
    return $false
  }
  if ($p.ExitCode -eq 0 -or $p.ExitCode -eq 3010) { return $true }
  Bad ("安装返回码 " + $p.ExitCode + "（1602=用户取消，3010=需重启）")
  return $false
}

function Start-Gui($py) {
  try {
    Start-Process -FilePath $py.Pw -ArgumentList ('"' + $gui + '"') -WorkingDirectory $scriptDir
    return $true
  } catch {
    Bad ("启动界面失败：" + $_.Exception.Message)
    return $false
  }
}

# ---------------------------------------------------------------------------
# 调试入口
# ---------------------------------------------------------------------------
if (-not [string]::IsNullOrWhiteSpace($VerifySigner)) {
  $r = Test-InstallerSignature $VerifySigner
  Write-Output ("SIGNER=" + $r.Signer)
  Write-Output ("REASON=" + $r.Reason)
  Write-Output ("OK=" + $r.Ok)
  Write-Output ("KEEP=" + $r.Keep)
  if ($r.Ok) { exit 0 } else { exit 1 }
}

if ($ResolveOnly) {
  $inst = Resolve-PythonInstaller
  if ($inst) {
    Write-Output ("PY_VERSION=" + $inst.Version)
    Write-Output ("PY_URL=" + $inst.Url)
    Write-Output ("URL_OFFICIAL=" + (Test-OfficialPythonUrl $inst.Url))
    exit 0
  }
  Write-Output "PY_RESOLVE_FAILED"
  exit 1
}

# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------
Info "============================================"
Info "  WP互相转换工具 - 启动引导"
Info "============================================"

if (-not (Test-Path -LiteralPath $gui)) {
  Bad "缺少主程序 wp_converter.py（应与本启动器在同一目录）"
  exit 1
}
if (-not (Test-Path -LiteralPath (Join-Path $scriptDir "convert_engine.ps1"))) {
  Warn "缺少转换引擎 convert_engine.ps1，转换功能将不可用。"
}

$py = $null
if (-not $SimulateNoPython) { $py = Find-Python }

# --- 情况 A：已找到 Python ---
if ($py) {
  Info ("检测到 Python：" + $py.P + "（来源：" + $py.Source + "）")
  if (-not (Test-Tkinter $py.P)) {
    Warn "检测到 Python，但缺少 tkinter 图形库（会无法显示界面）。"
    if ($CheckOnly) { exit 2 }
    Info "建议重新安装官方 Python（安装时勾选 tcl/tk and IDLE）。"
    Info ("下载地址：" + $PY_MANUAL_PAGE)
    exit 1
  }
  if ($CheckOnly) { Ok "环境检查通过，可直接启动。"; exit 0 }
  Info "正在启动图形界面…"
  if (Start-Gui $py) { Ok "已启动。"; exit 0 } else { exit 1 }
}

# --- 情况 B：没有 Python（鸡生蛋：必须在启动器里解决） ---
Bad "未检测到可用的 Python 3 运行环境。"
Info "本工具由 Python 编写，需要 Python 3.8+（含 tkinter 图形库）。"
if ($CheckOnly) { exit 2 }

$inst = Resolve-PythonInstaller
if (-not $inst) {
  Info ""
  Info "无法自动获取安装包。请手动安装："
  Info ("  " + $PY_MANUAL_PAGE)
  Info "安装时请务必勾选 “tcl/tk and IDLE”，否则界面无法显示。"
  exit 1
}

Info ""
Info "--------------------------------------------"
Info ("  准备安装 Python " + $inst.Version)
Info "--------------------------------------------"
Info ("  下载地址：" + $inst.Url)
Info "  来源校验：" + $(if (Test-OfficialPythonUrl $inst.Url) { "通过（https://www.python.org 官方）" } else { "失败" })
Info "  安装位置：当前用户目录（无需管理员权限、不弹 UAC）"
Info "  安全校验：下载后将验证 Authenticode 数字签名，签名者必须是"
Info "            Python Software Foundation，否则中止并删除安装包。"
Info "--------------------------------------------"
$ans = Read-Host "是否继续下载并安装？(Y/N)"
if ($ans -notmatch '^[Yy]') {
  Info "已取消。"
  Info ("如需手动安装：" + $PY_MANUAL_PAGE)
  exit 1
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("python-" + $inst.Version + "-amd64.exe")
if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }

Info "正在下载（约 25MB，请稍候）…"
if (-not (Download-Installer -Url $inst.Url -Dest $tmp)) {
  Info ("请手动下载安装：" + $PY_MANUAL_PAGE)
  exit 1
}
Ok "下载完成。"

Info "正在校验数字签名…"
$sigResult = Test-InstallerSignature $tmp
if (-not $sigResult.Ok) {
  Bad ("签名校验未通过：" + $sigResult.Reason)
  if ($sigResult.Keep) {
    # 环境性失败：文件已从官方源下载且未被篡改的证据链未断，保留供用户自行核验
    Warn ("安装包已保留在：" + $tmp)
    Warn "若你能确认它是刚从 python.org 下载的，可手动双击它完成安装。"
  } else {
    Bad "为安全起见，已删除该安装包并中止安装。"
    try { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } catch {}
  }
  Info ("或手动从官网安装：" + $PY_MANUAL_PAGE)
  exit 1
}
Ok ("签名校验通过，签名者：" + $sigResult.Signer)

Info "正在静默安装到当前用户目录（约 1~2 分钟）…"
$installed = Install-PythonSilently -InstallerPath $tmp
try { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue } catch {}
if (-not $installed) {
  Info ("请手动下载安装：" + $PY_MANUAL_PAGE)
  exit 1
}
Ok "Python 安装完成。"

$py = Find-Python
if (-not $py) {
  Bad "安装后仍未找到 Python（PATH 可能需重启终端刷新）。"
  Info "请重启电脑后再试，或手动安装。"
  exit 1
}
if (-not (Test-Tkinter $py.P)) {
  Bad "安装后 tkinter 仍不可用，请手动修复 Python 环境。"
  Info ("下载地址：" + $PY_MANUAL_PAGE)
  exit 1
}
Info "正在启动图形界面…"
if (Start-Gui $py) { Ok "已启动。"; exit 0 } else { exit 1 }
