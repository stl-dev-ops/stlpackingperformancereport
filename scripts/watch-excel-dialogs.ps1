param(
    [string]$OutputPath = "C:\dev\STLPackingPerformanceReport\artifacts\dialog-watch.log",
    [int]$TimeoutSeconds = 180
)

Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class Win32DialogWatch {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
}
"@

$shell = New-Object -ComObject WScript.Shell
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)

function Write-Log {
    param([string]$Message)
    $folder = Split-Path -Path $OutputPath -Parent
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
    }
    Add-Content -LiteralPath $OutputPath -Value ((Get-Date).ToString("s") + " | " + $Message)
}

function Get-VisibleWindows {
    $results = New-Object System.Collections.Generic.List[object]
    $callback = [Win32DialogWatch+EnumWindowsProc]{
        param($hWnd, $lParam)
        if (-not [Win32DialogWatch]::IsWindowVisible($hWnd)) {
            return $true
        }

        $titleBuilder = New-Object System.Text.StringBuilder 512
        [void][Win32DialogWatch]::GetWindowText($hWnd, $titleBuilder, $titleBuilder.Capacity)
        $classBuilder = New-Object System.Text.StringBuilder 256
        [void][Win32DialogWatch]::GetClassName($hWnd, $classBuilder, $classBuilder.Capacity)

        $title = $titleBuilder.ToString()
        if ([string]::IsNullOrWhiteSpace($title)) {
            return $true
        }

        $results.Add([pscustomobject]@{
            Handle = $hWnd
            Title = $title
            Class = $classBuilder.ToString()
        }) | Out-Null
        return $true
    }

    [void][Win32DialogWatch]::EnumWindows($callback, [IntPtr]::Zero)
    return $results
}

Write-Log "Dialog watchdog started"
while ((Get-Date) -lt $deadline) {
    $windows = Get-VisibleWindows
    foreach ($window in $windows) {
        $title = $window.Title
        if ($title -like '*Microsoft Visual Basic for Applications*' -and $title -like '*[break]*') {
            Write-Log ("VBE break window detected: " + $title)
            [void]$shell.AppActivate($title)
            Start-Sleep -Milliseconds 150
            $shell.SendKeys('%{F4}')
            Write-Log ("VBE break window close requested: " + $title)
            continue
        }

        if ($window.Class -eq '#32770' -and (
            $title -like '*Microsoft Visual Basic*' -or
            $title -like '*Compile error*' -or
            $title -like '*Microsoft Excel*'
        )) {
            Write-Log ("Dialog detected: " + $title)
            [void]$shell.AppActivate($title)
            Start-Sleep -Milliseconds 150
            $shell.SendKeys('{ENTER}')
            Write-Log ("Dialog dismissed: " + $title)
        }
    }
    Start-Sleep -Milliseconds 300
}
Write-Log "Dialog watchdog finished"