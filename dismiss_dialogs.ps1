param([int]$DurationSeconds = 120, [int]$ParentPid = 0)
# 守护进程：PDF->Word 期间监视 Office 的 NUIDialog。
# 若出现「将 PDF 转换为可编辑文档」确认框：勾选「不再显示此消息」并点确认。
# 附带父进程存活检查：Python 异常退出时本进程立即结束，不会变成 30 分钟的孤儿。
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public class WpWinEnum {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder t, int c);
}
"@

$attempts = @{}
$deadline = (Get-Date).AddSeconds($DurationSeconds)
$nextParentCheck = (Get-Date).AddSeconds(10)

while ((Get-Date) -lt $deadline) {
  # 父进程（Python）不在了 -> 立即退出
  if ($ParentPid -gt 0 -and (Get-Date) -ge $nextParentCheck) {
    $nextParentCheck = (Get-Date).AddSeconds(10)
    if (-not (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { break }
  }

  $dialogs = New-Object System.Collections.ArrayList
  [WpWinEnum]::EnumWindows({ param($h, $l)
    $sb = New-Object System.Text.StringBuilder 256
    [void][WpWinEnum]::GetClassName($h, $sb, 256)
    if ($sb.ToString() -eq "NUIDialog") { [void]$dialogs.Add($h) }
    return $true
  }, [IntPtr]::Zero) | Out-Null

  foreach ($h in $dialogs) {
    $key = [string]$h
    if (-not $attempts.ContainsKey($key)) { $attempts[$key] = 0 }
    if ($attempts[$key] -ge 3) { continue }
    $attempts[$key] = $attempts[$key] + 1

    try {
      $auto = [System.Windows.Automation.AutomationElement]::FromHandle($h)
      if ($null -eq $auto) { continue }

      # 1) 只处理与 PDF 相关的对话框（中英文均含 "PDF"）
      $texts = New-Object System.Collections.ArrayList
      try {
        $all = $auto.FindAll([System.Windows.Automation.TreeScope]::Descendants,
                             [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($el in $all) {
          try {
            if ($el.Current.ControlType.ProgrammaticName -eq "ControlType.Text") {
              [void]$texts.Add([string]$el.Current.Name)
            }
          } catch {}
        }
      } catch {}
      if (($texts -join " ") -notmatch "PDF") { continue }

      # 2) 勾选「不再显示此消息」（独立 try/catch）
      try {
        $cbCond = New-Object System.Windows.Automation.PropertyCondition(
          [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
          [System.Windows.Automation.ControlType]::CheckBox)
        $cb = $auto.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cbCond)
        if ($null -ne $cb) {
          $tp = $cb.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
          if ($tp.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) { $tp.Toggle() }
        }
      } catch {}

      # 3) 按钮：中文名 -> 英文名 -> 默认按钮 -> 首个非取消按钮
      $clicked = $false
      foreach ($name in @("确定", "OK", "是", "Yes")) {
        if ($clicked) { break }
        try {
          $c = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::NameProperty, $name)
          $btn = $auto.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $c)
          if ($null -ne $btn) {
            $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            $clicked = $true
          }
        } catch {}
      }
      if (-not $clicked) {
        try {
          $btnCond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button)
          $btns = $auto.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
          foreach ($b in $btns) {
            try {
              if ($b.Current.IsDefault) {
                $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                $clicked = $true
                break
              }
            } catch {}
          }
        } catch {}
      }
      if (-not $clicked) {
        try {
          $btnCond = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Button)
          $btns = $auto.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
          $cancelNames = @("取消", "Cancel", "否", "No")
          foreach ($b in $btns) {
            try {
              $n = [string]$b.Current.Name
              if ($cancelNames -contains $n) { continue }
              if ($b.Current.IsEnabled) {
                $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
                $clicked = $true
                break
              }
            } catch {}
          }
        } catch {}
      }
    } catch {}
  }
  Start-Sleep -Milliseconds 400
}
