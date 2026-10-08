param(
    [Parameter(Mandatory=$true)][string]$InputFile,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateRange(1,3600)][int]$TimeoutSeconds = 120,
    [switch]$Worker,
    [string]$StateDirectory
)
$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
$outPath = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($outPath) | Out-Null
$pdfPath = Join-Path $outPath ([IO.Path]::GetFileNameWithoutExtension($inputPath) + '.pdf')
if (Test-Path -LiteralPath $pdfPath) { throw 'PDF already exists. Choose a new QA directory.' }

if (-not $Worker) {
    # A COM call can block without reaching finally; bound the entire worker.
    $statePath = Join-Path $outPath ('.gongwen-render-' + [guid]::NewGuid().ToString('N'))
    [IO.Directory]::CreateDirectory($statePath) | Out-Null
    $existingWordIds = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $process = $null
    try {
        function Quote-Literal([string]$value) { return "'" + $value.Replace("'", "''") + "'" }
        $workerCommand = '$ProgressPreference = ''SilentlyContinue''; try { & ' + (Quote-Literal $PSCommandPath) + ' -Worker -InputFile ' + (Quote-Literal $inputPath) + ' -OutputDirectory ' + (Quote-Literal $statePath) + ' -StateDirectory ' + (Quote-Literal $statePath) + ' } catch { [Console]::Error.WriteLine($_.ToString()); exit 1 }'
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($workerCommand))
        $shellPath = (Get-Process -Id $PID).Path
        $process = Start-Process -FilePath $shellPath -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-OutputFormat', 'Text', '-EncodedCommand', $encoded) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $statePath 'stdout.txt') -RedirectStandardError (Join-Path $statePath 'stderr.txt')
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
            $stage = Get-Content -LiteralPath (Join-Path $statePath 'stage.txt') -Raw -ErrorAction SilentlyContinue
            if ($null -ne $stage) { $stage = $stage.Trim() }
            throw "Word rendering exceeded $TimeoutSeconds seconds at $stage; no final PDF was published."
        }
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) {
            throw ('Word rendering failed: ' + (Get-Content -LiteralPath (Join-Path $statePath 'stderr.txt') -Raw))
        }
        $result = Get-Content -LiteralPath (Join-Path $statePath 'result.json') -Raw | ConvertFrom-Json
        # File.Move does not replace an existing file on Windows.
        [IO.File]::Move($result.pdf, $pdfPath)
        $result.pdf = $pdfPath
        $result | ConvertTo-Json -Compress
    }
    finally {
        # Only stop the exact new Word instance recorded by this worker.
        $pidFile = Join-Path $statePath 'word-process.json'
        if (Test-Path -LiteralPath $pidFile) {
            $identity = Get-Content -LiteralPath $pidFile -Raw | ConvertFrom-Json
            $wordProcess = Get-Process -Id $identity.id -ErrorAction SilentlyContinue
            if ($null -ne $wordProcess -and $existingWordIds -notcontains $wordProcess.Id -and $wordProcess.ProcessName -eq 'WINWORD' -and $wordProcess.StartTime.ToUniversalTime().Ticks -eq [long]$identity.started) {
                Stop-Process -Id $wordProcess.Id -Force -ErrorAction SilentlyContinue
            }
        }
        if ($null -ne $process) { $process.Dispose() }
        $resolvedState = [IO.Path]::GetFullPath($statePath)
        if ($resolvedState.StartsWith($outPath.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedState -Recurse -Force
        }
    }
    return
}

$word = $null
$document = $null
$previousOptions = @{}
function Set-RenderStage([string]$value) { Set-Content -LiteralPath (Join-Path $StateDirectory 'stage.txt') -Value $value -Encoding UTF8 }
try {
    Set-RenderStage 'start Word'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    # Word exposes Hwnd on Window, not Application. A disposable blank document
    # identifies our COM instance before opening the supplied file.
    Set-RenderStage 'create tracking document'
    $document = $word.Documents.Add()
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class GongwenWordProcess { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    [uint32]$wordPid = 0
    [void][GongwenWordProcess]::GetWindowThreadProcessId([IntPtr]$word.ActiveWindow.Hwnd, [ref]$wordPid)
    $ownedProcess = Get-Process -Id $wordPid
    @{ id = $wordPid; started = $ownedProcess.StartTime.ToUniversalTime().Ticks.ToString() } | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $StateDirectory 'word-process.json') -Encoding UTF8
    $saveChanges = 0
    $document.Close([ref]$saveChanges)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null
    foreach ($option in @('UpdateLinksAtOpen', 'UpdateFieldsAtPrint', 'UpdateLinksAtPrint')) {
        $previousOptions[$option] = $word.Options.$option
        $word.Options.$option = $false
    }
    Set-RenderStage 'open input'
    $document = $word.Documents.Open($inputPath, $false, $true, $false)
    Set-RenderStage 'repaginate'
    $document.Repaginate()
    # PAGE, NUMPAGES and SECTIONPAGES only; preserve cached dates, links and text.
    Set-RenderStage 'update page fields'
    foreach ($story in $document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            foreach ($field in $range.Fields) {
                if ([int]$field.Type -in @(33, 26, 66)) { $null = $field.Update() }
            }
            $range = $range.NextStoryRange
        }
    }
    Set-RenderStage 'export PDF'
    $document.ExportAsFixedFormat($pdfPath, 17)
    [pscustomobject]@{ pdf = $pdfPath; pages = $document.ComputeStatistics(2); renderer = 'Microsoft Word'; sourceSaved = $false } | ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $StateDirectory 'result.json') -Encoding UTF8
}
finally {
    Set-RenderStage 'close Word'
    try {
        if ($null -ne $document) { $saveChanges = 0; $document.Close([ref]$saveChanges); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) }
    }
    finally {
        if ($null -ne $word) {
            foreach ($option in $previousOptions.Keys) { $word.Options.$option = $previousOptions[$option] }
            $saveChanges = 0
            $word.Quit([ref]$saveChanges)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
