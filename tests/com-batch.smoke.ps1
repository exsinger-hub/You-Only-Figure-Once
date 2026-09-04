param(
    [ValidateRange(0, 200)]
    [int]$ExtraShapeCount = 0
)

$ErrorActionPreference = "Stop"
$bridgePath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\scripts\powerpoint-bridge.ps1"))
$powerShellPath = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

if (@(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
    throw "Refusing to run the COM smoke test while a user PowerPoint process is open."
}

if (-not ("YouOnlyFigureOnceSmoke.FocusWindow" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace YouOnlyFigureOnceSmoke {
    public static class FocusWindow {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }
}
"@
}

function Invoke-Bridge {
    param([string]$Action, [Collections.IDictionary]$Arguments)
    $payloadJson = @{ action = $Action; arguments = $Arguments } | ConvertTo-Json -Depth 20 -Compress
    $payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payloadJson))
    $output = & $powerShellPath -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $bridgePath -PayloadBase64 $payload
    if ($LASTEXITCODE -ne 0) { throw "Bridge action '$Action' failed." }
    return $output | ConvertFrom-Json
}

$temporaryRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testDirectory = [IO.Path]::GetFullPath((Join-Path $temporaryRoot ("you-only-figure-once-com-smoke-" + [guid]::NewGuid().ToString("N"))))
if (-not $testDirectory.StartsWith($temporaryRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Resolved smoke-test directory is outside the system temporary directory."
}
$null = New-Item -ItemType Directory -Path $testDirectory
$outputPath = Join-Path $testDirectory "batch-smoke.pptx"
$createdApplication = $false
$processId = 0
$foregroundBefore = [YouOnlyFigureOnceSmoke.FocusWindow]::GetForegroundWindow()

try {
    $null = Invoke-Bridge "new_presentation" @{ focus_policy = "preserve"; maximize = $false }
    $createdApplication = $true
    $status = Invoke-Bridge "status" @{ focus_policy = "preserve" }
    $processId = [int]$status.active_application_process_id
    if (-not $status.active_presentation -or $processId -lt 1) { throw "The temporary PowerPoint presentation was not available through COM." }

    $operations = [Collections.Generic.List[object]]::new()
    $operations.Add(@{ type = "add_slide"; position = 1; layout = "blank"; name = "Smoke Slide" })
    $operations.Add(@{ type = "add_shape"; slide_index = 1; name = "smoke_box"; shape_type_id = 1; left = 72; top = 72; width = 144; height = 72; fill_color = "#DCEAF7"; line_color = "#355C7D" })
    $operations.Add(@{ type = "add_textbox"; slide_index = 1; name = "smoke_label"; left = 84; top = 92; width = 120; height = 30; text = "Background batch"; font_size = 16; font_color = "#16324F" })
    $operations.Add(@{ type = "update_shape"; slide_index = 1; shape_name = "smoke_box"; left = 76 })
    for ($index = 0; $index -lt $ExtraShapeCount; $index += 1) {
        $column = $index % 14
        $row = [math]::Floor($index / 14)
        $operations.Add(@{
            type = "add_shape"
            slide_index = 1
            name = "benchmark_shape_$index"
            shape_type_id = 1
            left = 20 + (48 * $column)
            top = 180 + (28 * $row)
            width = 34
            height = 16
            fill_color = "#EEF3F8"
            line_color = "#7990A6"
        })
    }

    $batchTimer = [Diagnostics.Stopwatch]::StartNew()
    $batch = Invoke-Bridge "apply_sequence" @{
        focus_policy = "preserve"
        pacing_mode = "fast"
        step_delay_ms = 0
        return_detail = "compact"
        operations = @($operations)
    }
    $batchTimer.Stop()
    if (-not $batch.success) {
        $failureDetail = if ($null -ne $batch.error) { $batch.error | ConvertTo-Json -Depth 8 -Compress } else { "no bridge error detail" }
        throw "COM sequence returned a partial failure at operation $($batch.failed_index): $failureDetail"
    }
    $expectedOperationCount = 4 + $ExtraShapeCount
    if ([int]$batch.applied_count -ne $expectedOperationCount -or [int]$batch.object_operations_applied -ne $expectedOperationCount) { throw "Unexpected COM sequence receipt." }
    if (-not $batch.view_untouched) { throw "The preserve-focus batch reported a PowerPoint view change." }

    $inspection = Invoke-Bridge "inspect" @{ focus_policy = "preserve"; max_slides = 1; max_shapes_per_slide = 20; include_text = $true }
    $expectedShapeCount = 2 + $ExtraShapeCount
    if ([int]$inspection.slides[0].shape_count -ne $expectedShapeCount) { throw "Expected $expectedShapeCount native objects on the smoke-test slide." }
    $null = Invoke-Bridge "save" @{ focus_policy = "preserve"; output_path = $outputPath; format = "pptx"; overwrite = $true }
    if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { throw "The smoke-test presentation was not saved." }

    $foregroundAfter = [YouOnlyFigureOnceSmoke.FocusWindow]::GetForegroundWindow()
    if ($foregroundAfter -ne $foregroundBefore) { throw "PowerPoint changed the user's foreground window during preserve-mode drawing." }

    [pscustomobject]@{
        passed = $true
        applied_count = [int]$batch.applied_count
        object_count = [int]$inspection.slides[0].shape_count
        batch_elapsed_ms = [int]$batchTimer.ElapsedMilliseconds
        bridge_process_count = 1
        view_untouched = [bool]$batch.view_untouched
        foreground_preserved = $true
    } | ConvertTo-Json -Compress
}
finally {
    if ($createdApplication) {
        try { $null = Invoke-Bridge "close_presentation" @{ focus_policy = "preserve"; confirm = $true; save_changes = "discard" } } catch {}
        if ($processId -gt 0) {
            try { $null = Invoke-Bridge "quit_application" @{ focus_policy = "preserve"; confirm = $true; expected_process_id = $processId } } catch {}
        }
    }
    if ((Test-Path -LiteralPath $testDirectory) -and $testDirectory.StartsWith($temporaryRoot, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $testDirectory -Recurse -Force
    }
}
