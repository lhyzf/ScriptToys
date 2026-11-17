#Requires -Version 7

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run this test with pwsh -NoProfile -STA -File tests\ui_hot_reload.ps1'
}

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework

$root = Split-Path $PSScriptRoot -Parent
$mainPath = Join-Path $root 'main_ui.ps1'
$tokens = $null
$parseErrors = $null
$mainAst = [System.Management.Automation.Language.Parser]::ParseFile($mainPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($name in 'Initialize-UiHelperTypes', 'Import-UiFunctions', 'New-MainWindow') {
    $definition = $mainAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name
    }, $true)
    if (-not $definition) { throw "Missing function: $name" }
    . ([scriptblock]::Create($definition.Extent.Text))
}

function Replace-TestFunction {
    param([string]$Source, [string]$Name, [string]$Replacement)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Source, [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors | Out-String) }
    $definition = $ast.Find({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name
    }, $true)
    if (-not $definition) { throw "Missing function: $Name" }
    return $Source.Substring(0, $definition.Extent.StartOffset) + $Replacement + $Source.Substring($definition.Extent.EndOffset)
}

function Wait-TestCondition {
    param([scriptblock]$Condition, [string]$Description)
    $timeout = [DateTime]::UtcNow.AddSeconds(12)
    while ([DateTime]::UtcNow -lt $timeout) {
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
            [Action]{}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle
        )
        if (& $Condition) { return }
        Start-Sleep -Milliseconds 50
    }
    throw "Timed out: $Description"
}

$directory = Join-Path ([System.IO.Path]::GetTempPath()) ("ScriptRunnerTest-" + [guid]::NewGuid())
[void][System.IO.Directory]::CreateDirectory($directory)
$uiFunctionsPath = Join-Path $directory 'functions.ps1'
$jobPath = Join-Path $directory 'job.ps1'
$configPath = Join-Path $directory 'config.json'
$timer = $null
$watcher = $null
$SyncHash = @{
    Host = $host
    ScriptRoot = $directory
    UiFunctionDefinitions = @{}
    UiFunctionsLastProcessedUtcTicks = 0L
    JobStates = @{}
    JobRecords = @{}
    ScheduledUi = [System.Collections.ArrayList]::new()
    ManualUi = [System.Collections.ArrayList]::new()
    RetiredUi = [System.Collections.ArrayList]::new()
    TestMessages = [System.Collections.ArrayList]::new()
}

try {
    $source = [System.IO.File]::ReadAllText((Join-Path $root 'ui\functions.ps1'))
    $source = Replace-TestFunction $source 'Show-Message' 'function Show-Message { param($Message, $Icon) [void]$SyncHash.TestMessages.Add($Message) }'
    [System.IO.File]::WriteAllText($uiFunctionsPath, $source)
    [System.IO.File]::WriteAllText($jobPath, "while (`$true) { Write-Output 'tick'; Start-Sleep -Milliseconds 120 }")
    $config = @{
        ui = @{ useBuiltInOutput = $true }
        longRunning = @(@{ id = 'test'; title = 'test'; displayName = 'start'; stopDisplayName = 'stop'; scriptPath = $jobPath })
        oneTime = @()
    } | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($configPath, $config)

    Initialize-UiHelperTypes
    Import-UiFunctions -Path $uiFunctionsPath
    $SyncHash.Config = Read-ValidatedConfig -Path $configPath
    $SyncHash.ScheduledOnStartCallback = { param($JobName) Invoke-ScheduledOnStart -JobName $JobName }
    $SyncHash.ScheduledOnExitCallback = { param($JobName) Invoke-ScheduledOnExit -JobName $JobName }
    $SyncHash.Form = New-MainWindow
    if ($SyncHash.Form.ThemeMode -ne [System.Windows.ThemeMode]::System) {
        throw 'Main window is not following the system theme'
    }
    $watcher = [ScriptRunner.ConfigChangeSignal]::new($directory, 'functions.ps1')
    $SyncHash.UiFunctionsWatcher = $watcher

    $tick = $mainAst.Find({
        param($node)
        $node -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            $node.Expression.Extent.Text -eq '$uiFunctionsRefreshTimer' -and
            $node.Member.Value -eq 'add_Tick'
    }, $true)
    if (-not $tick) { throw 'Missing UI function watcher callback' }
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $timer.add_Tick($tick.Arguments[0].ScriptBlock.GetScriptBlock())
    $timer.Start()

    $definition = $SyncHash.Config.longRunning[0]
    if (-not (Start-ScriptJob -JobName $definition.JobName -ScriptPath $jobPath -JobDescription 'test' -Item $definition.Item -TaskKey $definition.TaskKey -TaskKind Scheduled -OnStart $SyncHash.ScheduledOnStartCallback -OnExit $SyncHash.ScheduledOnExitCallback)) {
        throw 'Failed to start test job'
    }
    $job = $SyncHash.JobStates[$definition.JobName]
    $pidBefore = $job.ProcessId
    $outputTab = $job.OutputTab
    Wait-TestCondition { $job.OutputTextBox.Text.Contains('tick') } 'initial job output'

    $updated = $source.Replace('$button.Height = 25', '$button.Height = 42').Replace('$TextBox.AppendText($Line + [Environment]::NewLine)', '$TextBox.AppendText("updated: " + $Line + [Environment]::NewLine)')
    if ($updated -ceq $source) { throw 'Test replacements did not match' }
    [System.IO.File]::WriteAllText($uiFunctionsPath, $updated)
    Wait-TestCondition { $SyncHash.ScheduledUi[0].StartButton.Height -eq 42 -and $job.OutputTextBox.Text.Contains('updated: tick') } 'automatic reload and active job output'
    if ($job.Process.HasExited -or $job.ProcessId -ne $pidBefore -or
        -not [object]::ReferenceEquals($job.OutputTab, $outputTab) -or
        $SyncHash.ScheduledUi[0].StartButton.IsEnabled -or -not $SyncHash.ScheduledUi[0].StopButton.IsEnabled) {
        throw 'Reload did not preserve the running job, output tab, or button state'
    }

    $previousControl = $SyncHash.ScheduledUi[0].StartButton
    [System.IO.File]::WriteAllText($uiFunctionsPath, $updated + "`nfunction Invalid {")
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 1 } 'syntax error notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].StartButton) -or
        (New-StandardButton 'check').Height -ne 42 -or $job.Process.HasExited) {
        throw 'Syntax error did not preserve the previous UI and job'
    }

    $broken = (Replace-TestFunction $updated 'New-StandardButton' "function New-StandardButton { throw 'refresh failure' }") + "`nfunction Test-NewHelper { 'new' }"
    [System.IO.File]::WriteAllText($uiFunctionsPath, $broken)
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 2 } 'refresh error notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].StartButton) -or
        (New-StandardButton 'check').Height -ne 42 -or $job.Process.HasExited -or
        (Get-Command Test-NewHelper -CommandType Function -ErrorAction SilentlyContinue)) {
        throw 'Refresh failure did not restore the previous UI, methods, and job'
    }

    $removed = Replace-TestFunction $updated 'Get-UseBuiltInOutput' ''
    [System.IO.File]::WriteAllText($uiFunctionsPath, $removed)
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 3 } 'removed function notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].StartButton) -or
        -not (Get-Command Get-UseBuiltInOutput -CommandType Function -ErrorAction SilentlyContinue)) {
        throw 'Removing a loaded function was not rejected'
    }

    [System.IO.File]::Delete($uiFunctionsPath)
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 4 } 'deleted file notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].StartButton) -or $job.Process.HasExited) {
        throw 'Deleting the method file did not preserve the previous UI and job'
    }

    [System.IO.File]::WriteAllText($uiFunctionsPath, $source)
    Wait-TestCondition { $SyncHash.ScheduledUi[0].StartButton.Height -eq 25 } 'recovery after failure'
    $SyncHash.ScheduledUi[0].StopButton.RaiseEvent(
        [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
    )
    if ($SyncHash.JobStates.ContainsKey($definition.JobName)) { throw 'Rebuilt stop button did not stop the job' }
    'PASS: automatic reload, active job and output, rollback, removal, recovery, rebuilt event handler'
} finally {
    if ($timer) { $timer.Stop() }
    if ($watcher) { $watcher.Dispose() }
    foreach ($job in @($SyncHash.JobStates.Values)) {
        if ($job.Timer) { $job.Timer.Stop() }
        if ($job.Process) {
            if (-not $job.Process.HasExited) { $job.Process.Kill($true) }
            $job.Process.WaitForExit()
            $job.Process.Dispose()
        }
    }
    if (Test-Path -LiteralPath $directory) {
        Remove-Item -LiteralPath $directory -Recurse -Force
    }
}
