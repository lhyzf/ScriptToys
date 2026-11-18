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
$uiDirectory = Join-Path $directory 'ui'
[void][System.IO.Directory]::CreateDirectory($uiDirectory)
$uiFunctionsPath = Join-Path $uiDirectory 'functions.ps1'
$scheduledTogglePath = Join-Path $uiDirectory 'scheduled_toggle.ps1'
$customControlPath = Join-Path $uiDirectory 'test_control.ps1'
$jobPath = Join-Path $directory 'job.ps1'
$configPath = Join-Path $directory 'config.json'
$timer = $null
$watcher = $null
$windowShown = $false
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
    $toggleSource = [System.IO.File]::ReadAllText((Join-Path $root 'ui\scheduled_toggle.ps1'))
    $source = Replace-TestFunction $source 'Show-Message' 'function Show-Message { param($Message, $Icon) [void]$SyncHash.TestMessages.Add($Message) }'
    [System.IO.File]::WriteAllText($uiFunctionsPath, $source)
    [System.IO.File]::WriteAllText($scheduledTogglePath, $toggleSource)
    [System.IO.File]::WriteAllText($jobPath, "while (`$true) { Write-Output 'tick'; Start-Sleep -Milliseconds 120 }")
    $config = @{
        ui = @{ useBuiltInOutput = $true }
        longRunning = @(@{ id = 'test'; title = 'test'; displayName = 'start'; stopDisplayName = 'stop'; scriptPath = $jobPath })
        oneTime = @(@{ id = 'manual'; title = 'manual'; displayName = 'manual'; scriptPath = $jobPath })
    } | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($configPath, $config)

    Initialize-UiHelperTypes
    Import-UiFunctions -Directory $uiDirectory
    $SyncHash.Config = Read-ValidatedConfig -Path $configPath
    $SyncHash.ScheduledOnStartCallback = { param($JobName) Invoke-ScheduledOnStart -JobName $JobName }
    $SyncHash.ScheduledOnExitCallback = { param($JobName) Invoke-ScheduledOnExit -JobName $JobName }
    $SyncHash.Form = New-MainWindow
    if ($SyncHash.Form.ThemeMode -ne [System.Windows.ThemeMode]::System) {
        throw 'Main window is not following the system theme'
    }
    if ($SyncHash.Form.Content.RowDefinitions.Count -ne 3 -or
        [System.Windows.Controls.Grid]::GetRow($SyncHash.PanelSplitter) -ne 1 -or
        [System.Windows.Controls.Grid]::GetRow($SyncHash.OutputTabs) -ne 2 -or
        $SyncHash.PanelSplitter.ResizeDirection -ne [System.Windows.Controls.GridResizeDirection]::Rows -or
        $SyncHash.PanelSplitter.ResizeBehavior -ne [System.Windows.Controls.GridResizeBehavior]::PreviousAndNext -or
        $SyncHash.PanelSplitter.Visibility -ne [System.Windows.Visibility]::Collapsed) {
        throw 'Panel splitter is not configured between the operations and output panels'
    }
    $SyncHash.Form.ShowInTaskbar = $false
    $SyncHash.Form.WindowStyle = [System.Windows.WindowStyle]::None
    $SyncHash.Form.Opacity = 0
    $SyncHash.Form.Left = -10000
    $SyncHash.Form.Top = -10000
    $SyncHash.Form.Show()
    $windowShown = $true
    Wait-TestCondition { $SyncHash.ScheduledUi[0].ToggleButton.IsLoaded } 'scheduled toggle visual initialization'

    $scheduledPanel = $SyncHash.ScheduledUi[0].ToggleButton.Parent
    $scheduledButtons = @($scheduledPanel.Children | Where-Object { $_ -is [System.Windows.Controls.Button] })
    $toggleButton = $SyncHash.ScheduledUi[0].ToggleButton
    $manualButton = $SyncHash.ManualUi[0].Button
    if ($scheduledButtons.Count -ne 1 -or
        $toggleButton.Content -ne 'start' -or
        $toggleButton.ReadLocalValue([System.Windows.FrameworkElement]::StyleProperty) -ne [System.Windows.DependencyProperty]::UnsetValue -or
        $toggleButton.ReadLocalValue([System.Windows.Controls.Control]::TemplateProperty) -ne [System.Windows.DependencyProperty]::UnsetValue -or
        $toggleButton.Template -ne $manualButton.Template) {
        throw 'Long-running task did not retain the active theme style and template'
    }
    $startBorder = $toggleButton.Resources['ScriptRunner.Toggle.BackgroundElement']
    if ($null -eq $startBorder -or
        $startBorder -isnot [System.Windows.Controls.Border] -or
        $toggleButton.Background.ToString() -ne '#FFDCEBDD' -or
        $startBorder.Background.ToString() -ne '#FFDCEBDD' -or
        $toggleButton.Resources['ScriptRunner.Toggle.HoverBackground'].ToString() -ne '#FFC6DFC9' -or
        $toggleButton.Resources['ScriptRunner.Toggle.PressedBackground'].ToString() -ne '#FFB0D1B5' -or
        $toggleButton.Resources['ScriptRunner.Toggle.StateListeners'].Count -ne 2) {
        throw 'Start button background, hover, or pressed colors are incorrect'
    }
    Set-ScheduledToggleBackground -Button $toggleButton -InteractionState Hover
    if ($startBorder.Background.ToString() -ne '#FFC6DFC9') {
        throw 'Start button hover background is incorrect'
    }
    Set-ScheduledToggleBackground -Button $toggleButton -InteractionState Pressed
    if ($startBorder.Background.ToString() -ne '#FFB0D1B5') {
        throw 'Start button pressed background is incorrect'
    }
    Set-ScheduledToggleBackground -Button $toggleButton
    $watcher = [ScriptRunner.ConfigChangeSignal]::new($uiDirectory, '*.ps1')
    $SyncHash.UiFunctionsWatcher = $watcher
    $uiFunctionsDirectory = $uiDirectory

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
    if ($SyncHash.PanelSplitter.Visibility -ne [System.Windows.Visibility]::Visible) {
        throw 'Panel splitter was not shown with the output panel'
    }
    $SyncHash.OperationsExpander.IsExpanded = $false
    if ($SyncHash.PanelSplitter.Visibility -ne [System.Windows.Visibility]::Collapsed) {
        throw 'Panel splitter should be hidden when the operations panel is collapsed'
    }
    $SyncHash.OperationsExpander.IsExpanded = $true
    if ($SyncHash.PanelSplitter.Visibility -ne [System.Windows.Visibility]::Visible) {
        throw 'Panel splitter was not restored when the operations panel expanded'
    }

    [System.IO.File]::WriteAllText($customControlPath, "function New-TestControl { 'loaded' }")
    Wait-TestCondition {
        (Get-Command New-TestControl -CommandType Function -ErrorAction SilentlyContinue) -and
        $SyncHash.UiFunctionsLastProcessedUtcTicks -ge $watcher.LastChangeUtcTicks
    } 'automatic discovery of a new control file'
    if ((New-TestControl) -ne 'loaded' -or $job.Process.HasExited -or
        $job.ProcessId -ne $pidBefore -or -not [object]::ReferenceEquals($job.OutputTab, $outputTab)) {
        throw 'Adding a custom control file did not hot-reload while preserving the active job'
    }

    $toggleUpdated = $toggleSource.Replace('#F0DDDD', '#F1DDDD')
    if ($toggleUpdated -ceq $toggleSource) { throw 'Scheduled toggle test replacement did not match' }
    [System.IO.File]::WriteAllText($scheduledTogglePath, $toggleUpdated)
    Wait-TestCondition {
        $SyncHash.ScheduledUi[0].ToggleButton.Background.ToString() -eq '#FFF1DDDD' -or
        $SyncHash.TestMessages.Count -gt 0 -or
        ($watcher.LastChangeUtcTicks -gt 0 -and
            $SyncHash.UiFunctionsLastProcessedUtcTicks -ge $watcher.LastChangeUtcTicks)
    } 'automatic reload of dedicated control file'
    if ($SyncHash.ScheduledUi[0].ToggleButton.Background.ToString() -ne '#FFF1DDDD') {
        throw 'Dedicated control file did not update the running-state button background'
    }
    if ($job.Process.HasExited -or $job.ProcessId -ne $pidBefore -or
        -not [object]::ReferenceEquals($job.OutputTab, $outputTab)) {
        throw 'Reloading the dedicated control file did not preserve the active job and output tab'
    }

    [System.IO.File]::WriteAllText($scheduledTogglePath, $toggleSource)
    Wait-TestCondition { $SyncHash.ScheduledUi[0].ToggleButton.Background.ToString() -eq '#FFF0DDDD' } 'restore dedicated control file'

    $updated = $source.Replace('$button.Height = 25', '$button.Height = 42').Replace('$TextBox.AppendText($Line + [Environment]::NewLine)', '$TextBox.AppendText("updated: " + $Line + [Environment]::NewLine)')
    if ($updated -ceq $source) { throw 'Test replacements did not match' }
    [System.IO.File]::WriteAllText($uiFunctionsPath, $updated)
    Wait-TestCondition { $SyncHash.ScheduledUi[0].ToggleButton.Height -eq 42 -and $job.OutputTextBox.Text.Contains('updated: tick') } 'automatic reload and active job output'
    if ($job.Process.HasExited -or $job.ProcessId -ne $pidBefore -or
        -not [object]::ReferenceEquals($job.OutputTab, $outputTab) -or
        $SyncHash.ScheduledUi[0].ToggleButton.Content -ne 'stop' -or
        $SyncHash.ScheduledUi[0].ToggleButton.Tag -ne 'Running' -or
        $SyncHash.ScheduledUi[0].ToggleButton.Background.ToString() -ne '#FFF0DDDD' -or
        $startBorder.Background.ToString() -ne '#FFF0DDDD' -or
        $SyncHash.ScheduledUi[0].ToggleButton.Resources['ScriptRunner.Toggle.HoverBackground'].ToString() -ne '#FFE4C5C5' -or
        $SyncHash.ScheduledUi[0].ToggleButton.Resources['ScriptRunner.Toggle.PressedBackground'].ToString() -ne '#FFD6AAAA') {
        throw 'Reload did not preserve the running job, output tab, or button state'
    }

    $previousControl = $SyncHash.ScheduledUi[0].ToggleButton
    [System.IO.File]::WriteAllText($uiFunctionsPath, $updated + "`nfunction Invalid {")
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 1 } 'syntax error notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].ToggleButton) -or
        (New-StandardButton 'check').Height -ne 42 -or $job.Process.HasExited) {
        throw 'Syntax error did not preserve the previous UI and job'
    }

    $broken = (Replace-TestFunction $updated 'New-StandardButton' "function New-StandardButton { throw 'refresh failure' }") + "`nfunction Test-NewHelper { 'new' }"
    [System.IO.File]::WriteAllText($uiFunctionsPath, $broken)
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 2 } 'refresh error notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].ToggleButton) -or
        (New-StandardButton 'check').Height -ne 42 -or $job.Process.HasExited -or
        (Get-Command Test-NewHelper -CommandType Function -ErrorAction SilentlyContinue)) {
        throw 'Refresh failure did not restore the previous UI, methods, and job'
    }

    $removed = Replace-TestFunction $updated 'Get-UseBuiltInOutput' ''
    [System.IO.File]::WriteAllText($uiFunctionsPath, $removed)
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 3 } 'removed function notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].ToggleButton) -or
        -not (Get-Command Get-UseBuiltInOutput -CommandType Function -ErrorAction SilentlyContinue)) {
        throw 'Removing a loaded function was not rejected'
    }

    [System.IO.File]::Delete($uiFunctionsPath)
    Wait-TestCondition { $SyncHash.TestMessages.Count -ge 4 } 'deleted file notification'
    if (-not [object]::ReferenceEquals($previousControl, $SyncHash.ScheduledUi[0].ToggleButton) -or $job.Process.HasExited) {
        throw 'Deleting the method file did not preserve the previous UI and job'
    }

    [System.IO.File]::WriteAllText($uiFunctionsPath, $source)
    Wait-TestCondition { $SyncHash.ScheduledUi[0].ToggleButton.Height -eq 25 } 'recovery after failure'
    $SyncHash.ScheduledUi[0].ToggleButton.RaiseEvent(
        [System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)
    )
    if ($SyncHash.JobStates.ContainsKey($definition.JobName) -or
        $SyncHash.ScheduledUi[0].ToggleButton.Content -ne 'start' -or
        $SyncHash.ScheduledUi[0].ToggleButton.Tag -ne 'Stopped') {
        throw 'Rebuilt toggle button did not stop the job and return to the start state'
    }
    'PASS: UI extension discovery, custom control reload, native button template, state backgrounds, hover/pressed colors, active job and output, rollback, recovery'
} finally {
    if ($timer) { $timer.Stop() }
    if ($watcher) { $watcher.Dispose() }
    if ($windowShown -and $SyncHash.Form -and $SyncHash.Form.IsVisible) {
        $SyncHash.Form.Close()
    }
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
