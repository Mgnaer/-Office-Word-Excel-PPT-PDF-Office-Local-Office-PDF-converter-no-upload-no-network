param(
  [string]$InputFile,
  [string]$OutputFile,
  [string]$Mode,
  [string]$JobsJson,
  [string]$ResultJson,
  [string]$SessionJson,
  [string]$Backend = "office",
  [string]$Soffice = ""
)
$ErrorActionPreference = "Stop"

# 支持的转换模式：
#   word2pdf   Word   -> PDF   (Word.Application)
#   excel2pdf  Excel  -> PDF   (Excel.Application)
#   ppt2pdf    PPT    -> PDF   (PowerPoint.Application)
#   pdf2word   PDF    -> Word  (Word.Application)

function Set-PdfWarningSuppressed {
  try {
    $officeRoot = "HKCU:\Software\Microsoft\Office"
    $vers = @()
    if (Test-Path -LiteralPath $officeRoot) {
      $vers = @(Get-ChildItem -LiteralPath $officeRoot -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | ForEach-Object { $_.PSChildName })
    }
    if ($vers.Count -eq 0) { $vers = @("16.0") }
    foreach ($v in $vers) {
      $rp = $officeRoot + "\" + $v + "\Word\Options"
      if (-not (Test-Path -LiteralPath $rp)) { New-Item -Path $rp -Force -ErrorAction SilentlyContinue | Out-Null }
      if (Test-Path -LiteralPath $rp) {
        $cur = $null
        try { $cur = (Get-ItemProperty -Path $rp -Name "DisableConvertPdfWarning" -ErrorAction SilentlyContinue).DisableConvertPdfWarning } catch {}
        if ($cur -ne 1) { Set-ItemProperty -Path $rp -Name "DisableConvertPdfWarning" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue }
      }
    }
  } catch {}
}

$script:wpSessionFile = $null
# 三类 Office 进程分别记录自己创建的 PID，供 Python 精确清理
$script:wpPids = @{ word = $null; excel = $null; ppt = $null }

function Init-SessionFile {
  if ([string]::IsNullOrWhiteSpace($script:wpSessionFile)) { return }
  try {
    $dir = [System.IO.Path]::GetDirectoryName($script:wpSessionFile)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($script:wpSessionFile, "", (New-Object System.Text.UTF8Encoding($false)))
  } catch {}
}

function Register-Pid([int]$PidValue) {
  if ([string]::IsNullOrWhiteSpace($script:wpSessionFile)) { return }
  try {
    [System.IO.File]::AppendAllText($script:wpSessionFile, ([string]$PidValue + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
  } catch {}
}

function Write-KeepMark {
  if ([string]::IsNullOrWhiteSpace($script:wpSessionFile)) { return }
  try {
    [System.IO.File]::AppendAllText($script:wpSessionFile, ("KEEP" + [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
  } catch {}
}

function Write-JobResult {
  param([string]$ResultFile, $Utf8, [string]$Nl, [int]$Index, [string]$In, [string]$Out, [string]$ConvMode, [bool]$Ok, [string]$Err)
  $obj = [ordered]@{ index = $Index; input = $In; output = $Out; mode = $ConvMode; ok = $Ok; error = $Err }
  [System.IO.File]::AppendAllText($ResultFile, ((($obj | ConvertTo-Json -Compress -Depth 4)) + $Nl), $Utf8)
}

# ---------- 通用的 Office 进程探测 ----------
function Get-AppPids([string]$Image) {
  @(Get-Process -Name $Image -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
}

function Find-NewPid([string]$Image, [int[]]$Before) {
  for ($t = 0; $t -lt 25; $t++) {
    $after = @(Get-AppPids $Image)
    $fresh = @($after | Where-Object { $Before -notcontains $_ })
    if ($fresh.Count -gt 0) { return [int]$fresh[0] }
    Start-Sleep -Milliseconds 100
  }
  return $null
}

# 退出后确认进程真的消失，仍未退出则强杀（只处理本脚本自己创建的那个 PID）
function Stop-TrackedPid($PidValue) {
  if ($null -eq $PidValue) { return }
  for ($k = 0; $k -lt 30; $k++) {
    if (-not (Get-Process -Id $PidValue -ErrorAction SilentlyContinue)) { break }
    Start-Sleep -Milliseconds 200
  }
  if (Get-Process -Id $PidValue -ErrorAction SilentlyContinue) {
    try { Stop-Process -Id $PidValue -Force -ErrorAction SilentlyContinue } catch {}
  }
}

# 判断实例里是否还残留用户文档（本工具打开的文档都会在转换后关闭）
function Test-AppHasUserDocs {
  param($App, [string]$Kind)
  try {
    if ($Kind -eq "word")  { return ([int]$App.Documents.Count -gt 0) }
    if ($Kind -eq "excel") { return ([int]$App.Workbooks.Count -gt 0) }
    return ([int]$App.Presentations.Count -gt 0)
  } catch { return $false }
}

# 发起退出请求但不等待进程消失；返回 $true 表示“含用户文档，必须保留”
function Close-AppInstance {
  param($App, [string]$Kind)
  if ($null -eq $App) { return $false }
  if (Test-AppHasUserDocs $App $Kind) {
    if ($Kind -eq "word" -or $Kind -eq "excel") { try { $App.Visible = $true } catch {} }
    $label = @{ word = "Word"; excel = "Excel"; ppt = "PowerPoint" }[$Kind]
    $unit  = @{ word = "文档"; excel = "工作簿"; ppt = "演示文稿" }[$Kind]
    $n = 0
    if ($Kind -eq "word") { try { $n = [int]$App.Documents.Count } catch {} }
    elseif ($Kind -eq "excel") { try { $n = [int]$App.Workbooks.Count } catch {} }
    else { try { $n = [int]$App.Presentations.Count } catch {} }
    Write-Output ("NOTICE 检测到用户" + $unit + "被合并进本工具的 " + $label + " 实例（" + $n + " 个），已保留未关闭，请手动保存后关闭。")
    Write-KeepMark
    return $true
  }
  try {
    if ($Kind -eq "word") {
      $sc = 0
      try { $App.Quit([ref]$sc) } catch { try { $App.Quit() } catch {} }
    } else {
      try { $App.Quit() } catch {}
    }
  } catch {}
  try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {}
  return $false
}

function Assert-Output([string]$OutFull) {
  if (-not (Test-Path -LiteralPath $OutFull)) { throw "未生成输出文件: $OutFull" }
  if ((Get-Item -LiteralPath $OutFull).Length -le 0) { throw "输出文件为空: $OutFull" }
}

function Prepare-Paths([string]$InputPath, [string]$OutputPath) {
  $inFull  = [System.IO.Path]::GetFullPath($InputPath)
  $outFull = [System.IO.Path]::GetFullPath($OutputPath)
  if (-not (Test-Path -LiteralPath $inFull)) { throw "输入文件不存在: $inFull" }
  $outDir = [System.IO.Path]::GetDirectoryName($outFull)
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
  if (Test-Path -LiteralPath $outFull) { Remove-Item -LiteralPath $outFull -Force }
  return @{ In = $inFull; Out = $outFull }
}

# ================= Word =================
function New-WordApp {
  $before = @(Get-AppPids "WINWORD")
  $app = New-Object -ComObject Word.Application
  try { $app.Visible = $false } catch {}
  try { $app.DisplayAlerts = 0 } catch {}
  try { $app.Options.ConfirmConversions = $false } catch {}
  $script:wpPids.word = Find-NewPid "WINWORD" $before
  if ($null -ne $script:wpPids.word) { Register-Pid $script:wpPids.word }
  return $app
}

function Release-WordApp {
  param($App)
  if ($null -eq $App) { return }
  $kept = Close-AppInstance -App $App -Kind "word"
  $pidValue = $script:wpPids.word
  $script:wpPids.word = $null
  if (-not $kept) { Stop-TrackedPid $pidValue }
}

function Convert-OneWord {
  param($App, [string]$InputPath, [string]$OutputPath, [string]$ConvMode)
  if ($null -eq $App) { throw "Word 实例不可用" }
  $p = Prepare-Paths $InputPath $OutputPath
  $doc = $null
  try {
    if ($ConvMode -eq "word2pdf") {
      $doc = $App.Documents.Open($p.In, $false, $true)
      $doc.ExportAsFixedFormat($p.Out, 17)
    }
    elseif ($ConvMode -eq "pdf2word") {
      $doc = $App.Documents.Open($p.In, $false, $true, $false)
      $doc.SaveAs2($p.Out, 16)
    }
    else { throw "Word 后端不支持该模式: $ConvMode" }
  }
  finally {
    if ($null -ne $doc) {
      try { $doc.Close(0) } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch {}
    }
  }
  Assert-Output $p.Out
}

# ================= Excel =================
function New-ExcelApp {
  $before = @(Get-AppPids "EXCEL")
  $app = New-Object -ComObject Excel.Application
  try { $app.Visible = $false } catch {}
  try { $app.DisplayAlerts = $false } catch {}
  try { $app.ScreenUpdating = $false } catch {}
  try { $app.EnableEvents = $false } catch {}
  try { $app.AskToUpdateLinks = $false } catch {}
  $script:wpPids.excel = Find-NewPid "EXCEL" $before
  if ($null -ne $script:wpPids.excel) { Register-Pid $script:wpPids.excel }
  return $app
}

function Release-ExcelApp {
  param($App)
  if ($null -eq $App) { return }
  $kept = Close-AppInstance -App $App -Kind "excel"
  $pidValue = $script:wpPids.excel
  $script:wpPids.excel = $null
  if (-not $kept) { Stop-TrackedPid $pidValue }
}

function Convert-OneExcel {
  param($App, [string]$InputPath, [string]$OutputPath)
  if ($null -eq $App) { throw "Excel 实例不可用" }
  $p = Prepare-Paths $InputPath $OutputPath
  $wb = $null
  try {
    # Open(FileName, UpdateLinks=0, ReadOnly=$true) —— 避免外部链接提示
    $wb = $App.Workbooks.Open($p.In, 0, $true)
    # ExportAsFixedFormat(Type=0 即 xlTypePDF, Filename)
    $wb.ExportAsFixedFormat(0, $p.Out)
  }
  finally {
    if ($null -ne $wb) {
      try { $wb.Close($false) } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
    }
  }
  Assert-Output $p.Out
}

# ================= PowerPoint =================
function New-PptApp {
  $before = @(Get-AppPids "POWERPNT")
  $app = New-Object -ComObject PowerPoint.Application
  # 注意：PowerPoint 的 Application.Visible 不能设为 $false（会抛异常），
  # 隐藏窗口要靠 Presentations.Open 的 WithWindow=$false 参数实现。
  $script:wpPids.ppt = Find-NewPid "POWERPNT" $before
  if ($null -ne $script:wpPids.ppt) { Register-Pid $script:wpPids.ppt }
  return $app
}

function Release-PptApp {
  param($App)
  if ($null -eq $App) { return }
  $kept = Close-AppInstance -App $App -Kind "ppt"
  $pidValue = $script:wpPids.ppt
  $script:wpPids.ppt = $null
  if (-not $kept) { Stop-TrackedPid $pidValue }
}

function Convert-OnePpt {
  param($App, [string]$InputPath, [string]$OutputPath)
  if ($null -eq $App) { throw "PowerPoint 实例不可用" }
  $p = Prepare-Paths $InputPath $OutputPath
  $pres = $null
  try {
    # Open(FileName, ReadOnly=$true, Untitled=$false, WithWindow=$false)
    $pres = $App.Presentations.Open($p.In, $true, $false, $false)
    try {
      # ExportAsFixedFormat(Path, FixedFormatType=2 即 ppFixedFormatTypePDF)
      $pres.ExportAsFixedFormat($p.Out, 2)
    }
    catch {
      # 兜底：部分版本该重载不可用时，用 SaveAs(Path, ppSaveAsPDF=32)
      $pres.SaveAs($p.Out, 32)
    }
  }
  finally {
    if ($null -ne $pres) {
      try { $pres.Close() } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) } catch {}
    }
  }
  Assert-Output $p.Out
}

# ================= LibreOffice 备用后端 =================
function Find-Soffice {
  param([string]$Hint)
  if ($Hint -and (Test-Path -LiteralPath $Hint)) { return $Hint }
  $cands = New-Object System.Collections.ArrayList
  $pf = [Environment]::GetEnvironmentVariable("ProgramFiles")
  $pf86 = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
  $lad = [Environment]::GetEnvironmentVariable("LOCALAPPDATA")
  if ($pf) { [void]$cands.Add((Join-Path $pf "LibreOffice\program\soffice.exe")) }
  if ($pf86) { [void]$cands.Add((Join-Path $pf86 "LibreOffice\program\soffice.exe")) }
  if ($lad) { [void]$cands.Add((Join-Path $lad "Programs\LibreOffice\program\soffice.exe")) }
  foreach ($c in $cands) { if (Test-Path -LiteralPath $c) { return $c } }
  $cmd = Get-Command soffice -ErrorAction SilentlyContinue
  if ($cmd) { return $cmd.Source }
  return $null
}

function Convert-LO {
  param([string]$SofficePath, [string]$InputPath, [string]$OutputPath, [string]$ConvMode)
  $p = Prepare-Paths $InputPath $OutputPath
  $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("wp_lo_out_" + [guid]::NewGuid().ToString("N"))
  $profDir = Join-Path ([System.IO.Path]::GetTempPath()) ("wp_lo_prof_" + [guid]::NewGuid().ToString("N"))
  New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
  New-Item -ItemType Directory -Path $profDir -Force | Out-Null
  try {
    # LibreOffice 对 Word / Excel / PPT 都是统一 --convert-to pdf
    if ($ConvMode -eq "pdf2ppt" -or $ConvMode -eq "pdf2excel") {
      throw "PDF → PPT / PDF → Excel 需要 Microsoft Office：Excel 与 PowerPoint 的 API 不能直接打开 PDF"
    }
    $target = "pdf"
    if ($ConvMode -eq "pdf2word") { $target = "docx" }
    $profUrl = "file:///" + ($profDir -replace '\\', '/')
    $loArgs = @("-env:UserInstallation=" + $profUrl, "--headless", "--norestore", "--nolockcheck", "--nodefault", "--nofirststartwizard")
    if ($ConvMode -eq "pdf2word") { $loArgs += "--infilter=writer_pdf_import" }
    $loArgs += @("--convert-to", $target, "--outdir", $tmpDir, $p.In)
    $outText = & $SofficePath @loArgs 2>&1
    $produced = @(Get-ChildItem -LiteralPath $tmpDir -File -ErrorAction SilentlyContinue)
    if ($produced.Count -lt 1) {
      $msg = ($outText | Out-String).Trim()
      $detail = ""
      if ($msg) { $detail = "：" + $msg }
      throw ("LibreOffice 未生成输出文件" + $detail)
    }
    Move-Item -LiteralPath $produced[0].FullName -Destination $p.Out -Force
  }
  finally {
    Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $profDir -Recurse -Force -ErrorAction SilentlyContinue
  }
  Assert-Output $p.Out
}

# ================= 批量 / 单个 =================

# ================= PDF -> PPT / Excel（内容级重建） =================
# 说明：Excel 与 PowerPoint 的 COM API 都不支持打开 PDF，只有 Word 有 PDF 重排能力。
# 因此这里采用两阶段：先用 Word 把 PDF 重排成中间 DOCX，再解析 DOCX 的段落/
# 表格/图片，最后用 PowerPoint / Excel 重建目标文档。
# 定位是“内容级重建”（文字与图片按页/按表重组），不是版式级还原。

function Read-DocxContent([string]$DocxPath) {
  if (-not (Test-Path -LiteralPath $DocxPath)) { throw "中间 DOCX 不存在: $DocxPath" }
  Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
  $W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  $blocks = New-Object System.Collections.ArrayList
  $images = New-Object System.Collections.ArrayList
  $zip = $null
  try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($DocxPath)
    $docEntry = $zip.Entries | Where-Object { $_.FullName -eq "word/document.xml" } | Select-Object -First 1
    if ($null -eq $docEntry) { throw "中间 DOCX 缺少 word/document.xml" }
    $sr = New-Object System.IO.StreamReader($docEntry.Open())
    $xmlText = $sr.ReadToEnd()
    $sr.Dispose()
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($xmlText)

    # 按文档顺序遍历正文子节点，产出文字/表格/分页交错的 blocks
    $bodyNode = $xmlDoc.SelectSingleNode("//*[local-name()='body']")
    if ($null -ne $bodyNode) {
      foreach ($node in $bodyNode.ChildNodes) {
        $ln = [string]$node.LocalName
        if ($ln -eq "p") {
          $text = ""
          foreach ($t in $node.SelectNodes(".//*[local-name()='t']")) { $text = $text + $t.InnerText }
          $isBreak = $false
          foreach ($br in $node.SelectNodes(".//*[local-name()='br']")) {
            if ($br.GetAttribute("type", $W) -eq "page") { $isBreak = $true }
          }
          $style = ""
          $ps = $node.SelectSingleNode("./*[local-name()='pPr']/*[local-name()='pStyle']")
          if ($null -ne $ps) { $style = $ps.GetAttribute("val", $W) }
          $text = $text.Trim()
          if ($text.Length -gt 0) { [void]$blocks.Add(@{ kind = "text"; text = $text; style = $style }) }
          if ($isBreak) { [void]$blocks.Add(@{ kind = "pagebreak" }) }
        }
        elseif ($ln -eq "tbl") {
          $rows = New-Object System.Collections.ArrayList
          foreach ($tr in $node.SelectNodes("./*[local-name()='tr']")) {
            $cells = New-Object System.Collections.ArrayList
            foreach ($tc in $tr.SelectNodes("./*[local-name()='tc']")) {
              $cellText = ""
              foreach ($t in $tc.SelectNodes(".//*[local-name()='t']")) { $cellText = $cellText + $t.InnerText }
              [void]$cells.Add($cellText.Trim())
            }
            if ($cells.Count -gt 0) { [void]$rows.Add($cells) }
          }
          if ($rows.Count -gt 0) { [void]$blocks.Add(@{ kind = "table"; rows = $rows }) }
        }
      }
    }

    foreach ($entry in @($zip.Entries | Where-Object { $_.FullName -like "word/media/*" })) {
      $ext = [System.IO.Path]::GetExtension($entry.FullName).ToLower()
      if (@(".png", ".jpg", ".jpeg", ".bmp", ".gif") -contains $ext) {
        $tmp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("wp_img_" + [guid]::NewGuid().ToString("N") + $ext))
        [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $tmp, $true)
        [void]$images.Add($tmp)
      }
    }
  }
  catch {
    # 失败时把已经解压出来的图片清掉，避免在 %TEMP% 里堆积
    foreach ($img in $images) { Remove-Item -LiteralPath $img -Force -ErrorAction SilentlyContinue }
    throw
  }
  finally {
    if ($null -ne $zip) { $zip.Dispose() }
  }
  return @{ blocks = $blocks; images = $images }
}

# Word 的 PDF 重排常把视觉相邻的多行合并成一个长段落，
# 放进幻灯片会变成“一行超长文本”。这里按中文句末标点切成多行以便阅读。
function Split-LongLine([string]$Text, [int]$MaxLen) {
  $lines = New-Object System.Collections.ArrayList
  $t = $Text.Trim()
  if ($t.Length -eq 0) { return $lines }
  if ($t.Length -le $MaxLen) { [void]$lines.Add($t); return $lines }

  $parts = [regex]::Split($t, '(?<=[。！？；])')
  $cur = ""
  foreach ($seg in $parts) {
    $s = $seg.Trim()
    if ($s.Length -eq 0) { continue }
    if (($cur.Length + $s.Length) -gt $MaxLen -and $cur.Length -gt 0) {
      [void]$lines.Add($cur)
      $cur = $s
    }
    else { $cur = $cur + $s }
  }
  if ($cur.Trim().Length -gt 0) { [void]$lines.Add($cur) }

  # 单句本身仍超长则硬切
  $final = New-Object System.Collections.ArrayList
  foreach ($line in $lines) {
    if ($line.Length -le ($MaxLen * 2)) { [void]$final.Add($line) }
    else {
      $i = 0
      while ($i -lt $line.Length) {
        $len = [Math]::Min($MaxLen, $line.Length - $i)
        [void]$final.Add($line.Substring($i, $len))
        $i = $i + $MaxLen
      }
    }
  }
  return $final
}

function Build-PptFromContent {
  param($App, $Content, [string]$OutPath)
  if ($null -eq $App) { throw "PowerPoint 实例不可用" }
  $outFull = [System.IO.Path]::GetFullPath($OutPath)
  $outDir = [System.IO.Path]::GetDirectoryName($outFull)
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
  if (Test-Path -LiteralPath $outFull) { Remove-Item -LiteralPath $outFull -Force }

  $pres = $null
  try {
    $pres = $App.Presentations.Add($false)
    $pageW = 720.0
    $pageH = 540.0

    $slides = New-Object System.Collections.ArrayList
    $cur = New-Object System.Collections.ArrayList
    foreach ($b in $Content.blocks) {
      if ($b.kind -eq "pagebreak") {
        if ($cur.Count -gt 0) { [void]$slides.Add($cur); $cur = New-Object System.Collections.ArrayList }
        continue
      }
      if ($b.kind -eq "text") {
        foreach ($line in (Split-LongLine -Text ([string]$b.text) -MaxLen 58)) {
          [void]$cur.Add([string]$line)
        }
        if ($cur.Count -ge 12) { [void]$slides.Add($cur); $cur = New-Object System.Collections.ArrayList }
        continue
      }
      if ($b.kind -eq "table") {
        # 表格另起一页展示：每行 cells 用 | 连接，超过 12 行则分多张
        if ($cur.Count -gt 0) { [void]$slides.Add($cur); $cur = New-Object System.Collections.ArrayList }
        $tblLines = New-Object System.Collections.ArrayList
        foreach ($r in $b.rows) {
          [void]$tblLines.Add((@($r) -join "  |  "))
          if ($tblLines.Count -ge 12) { [void]$slides.Add($tblLines); $tblLines = New-Object System.Collections.ArrayList }
        }
        if ($tblLines.Count -gt 0) { [void]$slides.Add($tblLines) }
      }
    }
    if ($cur.Count -gt 0) { [void]$slides.Add($cur) }
    if ($slides.Count -eq 0 -and $Content.images.Count -eq 0) {
      $empty = New-Object System.Collections.ArrayList
      [void]$empty.Add("（未从 PDF 中提取到文本或图片）")
      [void]$slides.Add($empty)
    }

    foreach ($lines in $slides) {
      $idx = [int]$pres.Slides.Count + 1
      $slide = $pres.Slides.Add($idx, 12)
      $box = $slide.Shapes.AddTextbox(1, [single]36, [single]30, [single]648, [single]470)
      $tr = $box.TextFrame.TextRange
      $tr.Text = (@($lines) -join [Environment]::NewLine)
      $n = @($lines).Count
      if ($n -le 3) { $tr.Font.Size = 20 }
      elseif ($n -le 8) { $tr.Font.Size = 16 }
      else { $tr.Font.Size = 13 }
      $box.TextFrame.WordWrap = $true
      try { $box.TextFrame2.AutoSize = 2 } catch {}   # 2 = msoAutoSizeTextToFitShape
    }

    # 图片按等比缩放单独成页
    foreach ($img in $Content.images) {
      try {
        $idx = [int]$pres.Slides.Count + 1
        $slide = $pres.Slides.Add($idx, 12)
        $pic = $slide.Shapes.AddPicture($img, 0, -1, [single]36, [single]30)
        $pw = [double]$pic.Width
        $ph = [double]$pic.Height
        if ($pw -gt 0 -and $ph -gt 0) {
          $scale = [Math]::Min((648.0 / $pw), (470.0 / $ph))
          if ($scale -lt 1.0) {
            $pic.Width = [single]($pw * $scale)
            $pic.Height = [single]($ph * $scale)
          }
          $pic.Left = [single]((($pageW - [double]$pic.Width) / 2.0))
          $pic.Top = [single]((($pageH - [double]$pic.Height) / 2.0))
        }
      } catch {}
    }

    $pres.SaveAs($outFull, 24)
  }
  finally {
    if ($null -ne $pres) {
      try { $pres.Close() } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) } catch {}
    }
  }
  Assert-Output $outFull
}

function Build-XlsxFromContent {
  param($App, $Content, [string]$OutPath)
  if ($null -eq $App) { throw "Excel 实例不可用" }
  $outFull = [System.IO.Path]::GetFullPath($OutPath)
  $outDir = [System.IO.Path]::GetDirectoryName($outFull)
  if ($outDir -and -not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
  if (Test-Path -LiteralPath $outFull) { Remove-Item -LiteralPath $outFull -Force }

  $wb = $null
  try {
    $wb = $App.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "PDF内容"
    # 先把内容整理成二维数组，最后一次性写入。
    # 逐格赋值每次都是一次 COM 调用（1~5ms），大 PDF 会累积到十几秒；
    # 用 Range.Value2 整块写入可把上千次调用压成 1 次。
    $allRows = New-Object System.Collections.ArrayList
    $maxCol = 1
    foreach ($b in $Content.blocks) {
      if ($b.kind -eq "table") {
        foreach ($r in $b.rows) {
          $line = @($r)
          [void]$allRows.Add($line)
          if ($line.Count -gt $maxCol) { $maxCol = $line.Count }
        }
        [void]$allRows.Add(@(""))
        continue
      }
      if ($b.kind -eq "pagebreak") { [void]$allRows.Add(@("")); continue }
      if ($b.kind -eq "text") { [void]$allRows.Add(@([string]$b.text)) }
    }

    $rows = $allRows.Count
    if ($rows -gt 0) {
      $matrix = New-Object 'object[,]' $rows, $maxCol
      for ($r = 0; $r -lt $rows; $r++) {
        $line = @($allRows[$r])
        $lc = $line.Count
        for ($c = 0; $c -lt $maxCol; $c++) {
          if ($c -lt $lc) { $matrix[$r, $c] = [string]$line[$c] }
          else { $matrix[$r, $c] = "" }
        }
      }
      $range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows, $maxCol))
      $range.Value2 = $matrix

      try {
        $range.Columns.AutoFit() | Out-Null
        for ($c = 1; $c -le $maxCol; $c++) {
          try { if ([double]$ws.Columns.Item($c).ColumnWidth -gt 60) { $ws.Columns.Item($c).ColumnWidth = 60 } } catch {}
        }
      } catch {}
    }

    $wb.SaveAs($outFull, 51)
  }
  finally {
    if ($null -ne $wb) {
      try { $wb.Close($false) } catch {}
      try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
    }
  }
  Assert-Output $outFull
}

function Convert-PdfToPpt {
  param($WordApp, $PptApp, [string]$InputPath, [string]$OutputPath)
  $tmpDocx = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("wp_reflow_" + [guid]::NewGuid().ToString("N") + ".docx"))
  $content = $null
  try {
    Convert-OneWord -App $WordApp -InputPath $InputPath -OutputPath $tmpDocx -ConvMode "pdf2word"
    $content = Read-DocxContent $tmpDocx
    Build-PptFromContent -App $PptApp -Content $content -OutPath $OutputPath
  }
  finally {
    Remove-Item -LiteralPath $tmpDocx -Force -ErrorAction SilentlyContinue
    if ($null -ne $content) {
      foreach ($img in $content.images) { Remove-Item -LiteralPath $img -Force -ErrorAction SilentlyContinue }
    }
  }
}

function Convert-PdfToExcel {
  param($WordApp, $ExcelApp, [string]$InputPath, [string]$OutputPath)
  $tmpDocx = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), ("wp_reflow_" + [guid]::NewGuid().ToString("N") + ".docx"))
  $content = $null
  try {
    Convert-OneWord -App $WordApp -InputPath $InputPath -OutputPath $tmpDocx -ConvMode "pdf2word"
    $content = Read-DocxContent $tmpDocx
    Build-XlsxFromContent -App $ExcelApp -Content $content -OutPath $OutputPath
  }
  finally {
    Remove-Item -LiteralPath $tmpDocx -Force -ErrorAction SilentlyContinue
    if ($null -ne $content) {
      foreach ($img in $content.images) { Remove-Item -LiteralPath $img -Force -ErrorAction SilentlyContinue }
    }
  }
}

function Invoke-OneJob {
  param($Job, $Apps)
  $m = [string]$Job.mode
  if ($m -eq "word2pdf" -or $m -eq "pdf2word") {
    if ($null -eq $Apps.word) { $Apps.word = New-WordApp }
    Convert-OneWord -App $Apps.word -InputPath $Job.input -OutputPath $Job.output -ConvMode $m
  }
  elseif ($m -eq "excel2pdf") {
    if ($null -eq $Apps.excel) { $Apps.excel = New-ExcelApp }
    Convert-OneExcel -App $Apps.excel -InputPath $Job.input -OutputPath $Job.output
  }
  elseif ($m -eq "ppt2pdf") {
    if ($null -eq $Apps.ppt) { $Apps.ppt = New-PptApp }
    Convert-OnePpt -App $Apps.ppt -InputPath $Job.input -OutputPath $Job.output
  }
  elseif ($m -eq "pdf2ppt") {
    if ($null -eq $Apps.word) { $Apps.word = New-WordApp }
    if ($null -eq $Apps.ppt) { $Apps.ppt = New-PptApp }
    Convert-PdfToPpt -WordApp $Apps.word -PptApp $Apps.ppt -InputPath $Job.input -OutputPath $Job.output
  }
  elseif ($m -eq "pdf2excel") {
    if ($null -eq $Apps.word) { $Apps.word = New-WordApp }
    if ($null -eq $Apps.excel) { $Apps.excel = New-ExcelApp }
    Convert-PdfToExcel -WordApp $Apps.word -ExcelApp $Apps.excel -InputPath $Job.input -OutputPath $Job.output
  }
  else { throw "未知转换模式: $m" }
}

function Reset-AppAfterFailure {
  param($Apps, [string]$Mode)
  if ($Mode -eq "pdf2ppt") {
    try { Release-PptApp $Apps.ppt } catch {}
    try { Release-WordApp $Apps.word } catch {}
    $Apps.ppt = $null
    $Apps.word = $null
    return
  }
  if ($Mode -eq "pdf2excel") {
    try { Release-ExcelApp $Apps.excel } catch {}
    try { Release-WordApp $Apps.word } catch {}
    $Apps.excel = $null
    $Apps.word = $null
    return
  }
  if ($Mode -eq "excel2pdf") {
    try { Release-ExcelApp $Apps.excel } catch {}
    $Apps.excel = $null
  } elseif ($Mode -eq "ppt2pdf") {
    try { Release-PptApp $Apps.ppt } catch {}
    $Apps.ppt = $null
  } else {
    try { Release-WordApp $Apps.word } catch {}
    $Apps.word = $null
  }
}

function Rebuild-App {
  param($Apps, [string]$Mode)
  if ($Mode -eq "pdf2ppt") { $Apps.word = New-WordApp; $Apps.ppt = New-PptApp; return }
  if ($Mode -eq "pdf2excel") { $Apps.word = New-WordApp; $Apps.excel = New-ExcelApp; return }
  if ($Mode -eq "excel2pdf") { $Apps.excel = New-ExcelApp }
  elseif ($Mode -eq "ppt2pdf") { $Apps.ppt = New-PptApp }
  else { $Apps.word = New-WordApp }
}

function Release-AllApps($Apps) {
  # 阶段一：三类实例各自发起退出。Quit 本身是异步的，先全部触发，让关闭过程并行
  $pending = @{}
  foreach ($kind in @("word", "excel", "ppt")) {
    $app = $Apps[$kind]
    if ($null -eq $app) { continue }
    $kept = $false
    try { $kept = Close-AppInstance -App $app -Kind $kind } catch {}
    if ((-not $kept) -and ($null -ne $script:wpPids[$kind])) { $pending[$kind] = $script:wpPids[$kind] }
    $script:wpPids[$kind] = $null
    $Apps[$kind] = $null
  }
  # 阶段二：统一并行等待，最多 4 秒（而不是每个实例各等 6 秒、串行累加）
  for ($k = 0; $k -lt 20; $k++) {
    $alive = 0
    foreach ($kind in @($pending.Keys)) {
      if (Get-Process -Id $pending[$kind] -ErrorAction SilentlyContinue) { $alive++ }
    }
    if ($alive -eq 0) { break }
    Start-Sleep -Milliseconds 200
  }
  # 阶段三：仍未退出的强杀（只处理本脚本自己创建、且不在 KEEP 名单里的 PID）
  foreach ($kind in @($pending.Keys)) {
    if (Get-Process -Id $pending[$kind] -ErrorAction SilentlyContinue) {
      try { Stop-Process -Id $pending[$kind] -Force -ErrorAction SilentlyContinue } catch {}
    }
  }
}

function Run-Batch {
  param([string]$JobsFile, [string]$ResultFile, [string]$Bk, [string]$SofficePath)
  if (-not (Test-Path -LiteralPath $JobsFile)) { throw "任务清单不存在: $JobsFile" }
  if ([string]::IsNullOrWhiteSpace($ResultFile)) { throw "缺少 -ResultJson" }

  $raw = Get-Content -LiteralPath $JobsFile -Raw -Encoding UTF8
  $parsed = ConvertFrom-Json -InputObject $raw
  $jobs = @($parsed)
  if ($jobs.Count -eq 1 -and ($jobs[0] -is [System.Array])) { $jobs = @($jobs[0]) }
  if ($jobs.Count -eq 0) { throw "任务清单为空" }

  if (Test-Path -LiteralPath $ResultFile) { Remove-Item -LiteralPath $ResultFile -Force }
  Init-SessionFile
  $utf8 = New-Object System.Text.UTF8Encoding($false)
  $nl = [Environment]::NewLine

  if ($Bk -eq "libreoffice") {
    # 已知限制：soffice 的 --convert-to 无法复用进程，每个文件都要冷启动一次
    #（约 3~5 秒）并创建独立 profile。批量场景明显慢于 Word/Excel/PPT 直连后端。
    for ($i = 0; $i -lt $jobs.Count; $i++) {
      $job = $jobs[$i]
      $ok = $false
      $err = ""
      try { Convert-LO -SofficePath $SofficePath -InputPath $job.input -OutputPath $job.output -ConvMode ([string]$job.mode); $ok = $true }
      catch { $err = [string]$_.Exception.Message }
      Write-JobResult -ResultFile $ResultFile -Utf8 $utf8 -Nl $nl -Index $i -In ([string]$job.input) -Out ([string]$job.output) -ConvMode ([string]$job.mode) -Ok $ok -Err $err
      Write-Output ("PROGRESS {0} {1} {2}" -f ($i + 1), $jobs.Count, $(if ($ok) { "1" } else { "0" }))
    }
    return
  }

  # pdf2ppt / pdf2excel 内部同样调用 Word 做 PDF 重排，因此也要抑制确认框
  $pdfModes = @("pdf2word", "pdf2ppt", "pdf2excel")
  $needSuppress = $false
  foreach ($j in $jobs) { if ($pdfModes -contains [string]$j.mode) { $needSuppress = $true; break } }
  if ($needSuppress) { Set-PdfWarningSuppressed }

  $apps = @{ word = $null; excel = $null; ppt = $null }
  try {
    for ($i = 0; $i -lt $jobs.Count; $i++) {
      $job = $jobs[$i]
      $ok = $false
      $err = ""
      $fatal = $false
      try {
        Invoke-OneJob -Job $job -Apps $apps
        $ok = $true
      }
      catch {
        $err = [string]$_.Exception.Message
        # 失败后对应应用的 COM 状态可能被污染：释放并重建该应用实例
        Reset-AppAfterFailure -Apps $apps -Mode ([string]$job.mode)
        Start-Sleep -Milliseconds 500
        try { Rebuild-App -Apps $apps -Mode ([string]$job.mode) }
        catch { $err = $err + "（Office 实例重启失败，批次终止）"; $fatal = $true }
      }
      Write-JobResult -ResultFile $ResultFile -Utf8 $utf8 -Nl $nl -Index $i -In ([string]$job.input) -Out ([string]$job.output) -ConvMode ([string]$job.mode) -Ok $ok -Err $err
      Write-Output ("PROGRESS {0} {1} {2}" -f ($i + 1), $jobs.Count, $(if ($ok) { "1" } else { "0" }))
      if ($fatal) { throw "Office 实例无法恢复，批次已终止" }
    }
  }
  finally {
    Release-AllApps $apps
  }
}

function Run-Single {
  param([string]$In, [string]$Out, [string]$ConvMode, [string]$Bk, [string]$SofficePath)
  if ([string]::IsNullOrWhiteSpace($In))  { throw "缺少 -InputFile" }
  if ([string]::IsNullOrWhiteSpace($Out)) { throw "缺少 -OutputFile" }
  if ([string]::IsNullOrWhiteSpace($ConvMode)) { throw "缺少 -Mode (word2pdf|excel2pdf|ppt2pdf|pdf2word)" }
  Init-SessionFile
  if ($Bk -eq "libreoffice") {
    Convert-LO -SofficePath $SofficePath -InputPath $In -OutputPath $Out -ConvMode $ConvMode
  }
  else {
    if (@("pdf2word", "pdf2ppt", "pdf2excel") -contains $ConvMode) { Set-PdfWarningSuppressed }
    $apps = @{ word = $null; excel = $null; ppt = $null }
    try {
      $job = [pscustomobject]@{ input = $In; output = $Out; mode = $ConvMode }
      Invoke-OneJob -Job $job -Apps $apps
    }
    finally {
      Release-AllApps $apps
    }
  }
  Write-Output ("OK:" + [System.IO.Path]::GetFullPath($Out))
}

try {
  $bk = $Backend.ToLower()
  $soPath = $null
  if ($bk -eq "libreoffice") {
    $soPath = Find-Soffice -Hint $Soffice
    if (-not $soPath) { throw "未找到 LibreOffice（soffice.exe），请先安装或改用 Office 后端" }
  }
  if (-not [string]::IsNullOrWhiteSpace($JobsJson)) {
    if ([string]::IsNullOrWhiteSpace($SessionJson)) { $SessionJson = $ResultJson + ".session" }
    $script:wpSessionFile = $SessionJson
    Run-Batch -JobsFile $JobsJson -ResultFile $ResultJson -Bk $bk -SofficePath $soPath
  }
  else {
    if ([string]::IsNullOrWhiteSpace($SessionJson)) { $SessionJson = $OutputFile + ".session" }
    $script:wpSessionFile = $SessionJson
    Run-Single -In $InputFile -Out $OutputFile -ConvMode $Mode -Bk $bk -SofficePath $soPath
  }
}
catch {
  Write-Error ([string]$_.Exception.Message)
  exit 1
}
