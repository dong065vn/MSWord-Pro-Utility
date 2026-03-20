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
    & $autoIt3 (Join-Path $root "Tools\Run_TabSmokeTest.au3")
    if ($LASTEXITCODE -ne 0) { throw "Tab smoke failed with exit code $LASTEXITCODE" }
}

Run-Step "Runtime smoke" {
    & $autoIt3 (Join-Path $root "Tools\Run_RuntimeSmokeTest.au3")
    if ($LASTEXITCODE -ne 0) { throw "Runtime smoke failed with exit code $LASTEXITCODE" }
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
        & $autoIt3 (Join-Path $root "Tools\Run_FocusedTabTest.au3") $case
        if ($LASTEXITCODE -ne 0) { throw "Focused case '$case' failed with exit code $LASTEXITCODE" }
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
    & $autoIt3 (Join-Path $root "Tests\Test_AdvancedExportsAndSmartFix.au3")
    if ($LASTEXITCODE -ne 0) { throw "Advanced exports + SmartFix failed with exit code $LASTEXITCODE" }
}

Run-Step "Advanced document ops" {
    & $autoIt3 (Join-Path $root "Tests\Test_AdvancedDocumentOps.au3")
    if ($LASTEXITCODE -ne 0) { throw "Advanced document ops failed with exit code $LASTEXITCODE" }
}

Run-Step "AIFormat + Cleanup" {
    & $autoIt3 (Join-Path $root "Tests\Test_AIFormatAndCleanup.au3")
    if ($LASTEXITCODE -ne 0) { throw "AIFormat + Cleanup failed with exit code $LASTEXITCODE" }
}

Run-Step "AIFormat markdown structures" {
    & $autoIt3 (Join-Path $root "Tests\Test_AIFormatMarkdownStructures.au3")
    if ($LASTEXITCODE -ne 0) { throw "AIFormat markdown structures failed with exit code $LASTEXITCODE" }
}

Run-Step "AIBeautify + Italic" {
    & $autoIt3 (Join-Path $root "Tests\Test_AIBeautifyAndItalic.au3")
    if ($LASTEXITCODE -ne 0) { throw "AIBeautify + Italic failed with exit code $LASTEXITCODE" }
}

Run-Step "AILaTeX + Emoji" {
    & $autoIt3 (Join-Path $root "Tests\Test_AILaTeXAndEmoji.au3")
    if ($LASTEXITCODE -ne 0) { throw "AILaTeX + Emoji failed with exit code $LASTEXITCODE" }
}

Run-Step "AIPreview counts" {
    & $autoIt3 (Join-Path $root "Tests\Test_AIPreviewCounts.au3")
    if ($LASTEXITCODE -ne 0) { throw "AIPreview counts failed with exit code $LASTEXITCODE" }
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
Write-Summary "  Test_AdvancedDocumentOps.out.txt"
Write-Summary "  Test_AIFormatAndCleanup.out.txt"
Write-Summary "  Test_AIFormatMarkdownStructures.out.txt"
Write-Summary "  Test_AIBeautifyAndItalic.out.txt"
Write-Summary "  Test_AILaTeXAndEmoji.out.txt"
Write-Summary "  Test_AIPreviewCounts.out.txt"

Write-Host "Done. Summary: $summaryPath"
