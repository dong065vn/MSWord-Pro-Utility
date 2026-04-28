$ErrorActionPreference = "Stop"

$root = "C:\du_an_test\PDF_to_Word_Fixer_v6update"
$autoIt3 = "C:\Program Files (x86)\AutoIt3\AutoIt3.exe"
$au3Check = "C:\Program Files (x86)\AutoIt3\Au3Check.exe"
$logsDir = Join-Path $root "Tests\Logs"
$focusedDir = Join-Path $logsDir "Focused"
$summaryPath = Join-Path $logsDir "RegressionSummary.txt"

New-Item -ItemType Directory -Force $logsDir | Out-Null
New-Item -ItemType Directory -Force $focusedDir | Out-Null

function Write-Summary {
    param([string]$Line)
    Add-Content -Path $summaryPath -Value $Line
}

function Run-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )
    Write-Host "Running: $Name"
    try {
        & $Action
        Write-Summary "PASS: $Name"
    } catch {
        Write-Summary "FAIL: $Name"
        Write-Summary "  $_"
        throw
    }
}

function Assert-LogSuccess {
    param(
        [string]$Path,
        [string]$RequiredLine
    )
    $deadline = (Get-Date).AddSeconds(45)
    $content = @()
    while ((Get-Date) -lt $deadline) {
        if (Test-Path $Path) {
            $content = Get-Content -Path $Path
            $hasRequired = [bool]($content | Where-Object { $_ -eq $RequiredLine })
            if ($hasRequired) {
                $fail = $content | Where-Object { $_ -like 'FAIL:*' } | Select-Object -First 1
                if ($fail) { throw "Log failed: $Path :: $fail" }
                return
            }
        }
        Start-Sleep -Milliseconds 500
    }
    if (-not (Test-Path $Path)) { throw "Missing log: $Path" }
    $fail = $content | Where-Object { $_ -like 'FAIL:*' } | Select-Object -First 1
    if ($fail) { throw "Log failed: $Path :: $fail" }
    throw "Log missing required success line: $RequiredLine"
}

function Invoke-AutoItScript {
    param(
        [string]$ScriptPath,
        [string[]]$Arguments = @(),
        [string]$StdOutPath = "",
        [string]$StdErrPath = ""
    )
    $startArgs = @{
        FilePath = $autoIt3
        ArgumentList = @($ScriptPath) + $Arguments
        Wait = $true
        PassThru = $true
        WindowStyle = 'Hidden'
    }
    if ($StdOutPath -ne "") {
        if (Test-Path $StdOutPath) { Remove-Item $StdOutPath -Force }
        $startArgs.RedirectStandardOutput = $StdOutPath
    }
    if ($StdErrPath -ne "") {
        if (Test-Path $StdErrPath) { Remove-Item $StdErrPath -Force }
        $startArgs.RedirectStandardError = $StdErrPath
    }
    $process = Start-Process @startArgs
    if ($process.ExitCode -ne 0) {
        throw "AutoIt script failed with exit code $($process.ExitCode): $ScriptPath"
    }
}

function Run-AutoItTestWithDialogClose {
    param(
        [string]$ScriptPath,
        [string[]]$WindowTitles,
        [string]$StdOutPath,
        [string]$StdErrPath
    )

    if (Test-Path $StdOutPath) { Remove-Item $StdOutPath -Force }
    if (Test-Path $StdErrPath) { Remove-Item $StdErrPath -Force }

    $job = Start-Job -ArgumentList $WindowTitles -ScriptBlock {
        param($Titles)
        $ws = New-Object -ComObject WScript.Shell
        while ($true) {
            Start-Sleep -Milliseconds 700
            foreach ($title in $Titles) {
                try {
                    if ($ws.AppActivate($title)) {
                        $ws.SendKeys('{ENTER}')
                    }
                } catch {
                }
            }
        }
    }

    try {
        $cmdLine = '"' + $autoIt3 + '" "' + $ScriptPath + '" > "' + $StdOutPath + '" 2> "' + $StdErrPath + '"'
        & cmd.exe /c $cmdLine
        if ($LASTEXITCODE -ne 0) {
            throw "AutoIt test failed with exit code $LASTEXITCODE"
        }
    } finally {
        if ($job) {
            Stop-Job $job | Out-Null
            Remove-Job $job | Out-Null
        }
    }
}

Set-Content -Path $summaryPath -Value @(
    "PDF to Word Fixer Pro - Regression Summary"
    "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    ""
)

Run-Step "Au3Check Main.au3" {
    & $au3Check (Join-Path $root "Main.au3") *> (Join-Path $logsDir "Au3Check.log")
    if ($LASTEXITCODE -ne 0) { throw "Au3Check failed with exit code $LASTEXITCODE" }
}

Run-Step "Tab smoke" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tools\Run_TabSmokeTest.au3")
    Assert-LogSuccess -Path (Join-Path $logsDir "TabSmokeTest.log") -RequiredLine "PASS: Da chuyen het 9 tab ma khong thay app bi dong."
}

Run-Step "Runtime smoke" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tools\Run_RuntimeSmokeTest.au3")
    Assert-LogSuccess -Path (Join-Path $logsDir "RuntimeSmokeTest.log") -RequiredLine "PASS: Hoan tat runtime smoke test qua 9 tab."
}

$cases = @(
    "pdf_fix",
    "format",
    "tools",
    "toc",
    "copy_style",
    "advanced",
    "quick_utils",
    "smart_fix",
    "ai_format"
)

foreach ($case in $cases) {
    Run-Step "Focused $case" {
        Invoke-AutoItScript -ScriptPath (Join-Path $root "Tools\Run_FocusedTabTest.au3") -Arguments @($case)
        Assert-LogSuccess -Path (Join-Path $focusedDir "$case.log") -RequiredLine "PASS: completed"
    }
}

Run-Step "Hotkey flow" {
    Run-AutoItTestWithDialogClose `
        -ScriptPath (Join-Path $root "Tests\Test_HotkeyFlow.au3") `
        -WindowTitles @("Test Complete", "Test Failed", "AutoIt Error") `
        -StdOutPath (Join-Path $logsDir "Test_HotkeyFlow.out.txt") `
        -StdErrPath (Join-Path $logsDir "Test_HotkeyFlow.err.txt")
}

Run-Step "SaveToNormalDotm" {
    Run-AutoItTestWithDialogClose `
        -ScriptPath (Join-Path $root "Tests\Test_SaveToNormalDotm.au3") `
        -WindowTitles @("SaveToNormalDotm Test Results", "AutoIt Error") `
        -StdOutPath (Join-Path $logsDir "Test_SaveToNormalDotm.out.txt") `
        -StdErrPath (Join-Path $logsDir "Test_SaveToNormalDotm.err.txt")
}

Run-Step "Advanced exports + SmartFix" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AdvancedExportsAndSmartFix.au3")
}

Run-Step "Process tracker + DOM helpers" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_ProcessTrackerDomHelpers.au3")
}

Run-Step "WordPerf helpers" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_WordPerfHelpers.au3") `
        -StdOutPath (Join-Path $logsDir "Test_WordPerfHelpers.out.txt") `
        -StdErrPath (Join-Path $logsDir "Test_WordPerfHelpers.err.txt")
}

Run-Step "Smart Fix Pro pipeline" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_SmartFixProPipeline.au3") `
        -StdOutPath (Join-Path $logsDir "Test_SmartFixProPipeline.out.txt") `
        -StdErrPath (Join-Path $logsDir "Test_SmartFixProPipeline.err.txt")
}

Run-Step "Advanced document ops" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AdvancedDocumentOps.au3")
}

Run-Step "AIFormat + Cleanup" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AIFormatAndCleanup.au3")
}

Run-Step "AIFormat markdown structures" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AIFormatMarkdownStructures.au3")
}

Run-Step "AIBeautify + Italic" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AIBeautifyAndItalic.au3")
}

Run-Step "AILaTeX + Emoji" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AILaTeXAndEmoji.au3")
}

Run-Step "AIPreview counts" {
    Invoke-AutoItScript -ScriptPath (Join-Path $root "Tests\Test_AIPreviewCounts.au3")
}

Write-Summary ""
Write-Summary "Artifacts:"
Write-Summary "  Au3Check.log"
Write-Summary "  TabSmokeTest.log"
Write-Summary "  RuntimeSmokeTest.log"
Write-Summary "  Focused\\*.log"
Write-Summary "  Test_HotkeyFlow.out.txt"
Write-Summary "  Test_SaveToNormalDotm.out.txt"
Write-Summary "  Test_AdvancedExportsAndSmartFix.out.txt"
Write-Summary "  Test_ProcessTrackerDomHelpers.out.txt"
Write-Summary "  Test_WordPerfHelpers.out.txt"
Write-Summary "  Test_SmartFixProPipeline.out.txt"
Write-Summary "  Test_AdvancedDocumentOps.out.txt"
Write-Summary "  Test_AIFormatAndCleanup.out.txt"
Write-Summary "  Test_AIFormatMarkdownStructures.out.txt"
Write-Summary "  Test_AIBeautifyAndItalic.out.txt"
Write-Summary "  Test_AILaTeXAndEmoji.out.txt"
Write-Summary "  Test_AIPreviewCounts.out.txt"

Write-Host "Done. Summary: $summaryPath"
