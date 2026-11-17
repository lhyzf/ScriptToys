#Requires -Version 7

# 确保脚本以 UTF-8 编码保存，以支持中文字符

# 设置变量
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Host = $host
$SyncHash.ScriptRoot = $PSScriptRoot
$SyncHash.JobStates = @{}
$SyncHash.JobRecords = @{}
$SyncHash.ScheduledUi = [System.Collections.ArrayList]::new()
$SyncHash.ManualUi = [System.Collections.ArrayList]::new()
$SyncHash.RetiredUi = [System.Collections.ArrayList]::new()
$SyncHash.ConfigLastProcessedUtcTicks = 0L
$SyncHash.UiFunctionsLastProcessedUtcTicks = 0L
$SyncHash.UiFunctionDefinitions = @{}

# UI 运行空间
$UiRunspace = [runspacefactory]::CreateRunspace()
$UiRunspace.ApartmentState = 'STA'
$UiRunspace.ThreadOptions = 'ReuseThread'
$UiRunspace.Open()
$UiRunspace.SessionStateProxy.SetVariable('SyncHash', $SyncHash)

# UI 脚本
$UiPowerShell = [PowerShell]::Create().AddScript({
    # 设置当前线程的文化，以防万一
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'zh-CN'

    # 捕获 UI 线程中的错误
    trap {
        $SyncHash.Host.UI.WriteErrorLine("UI Error: $_`nLine: $($_.InvocationInfo.ScriptLineNumber)`n$($_.InvocationInfo.Line)")
    }

    try {
        foreach ($assemblyName in 'PresentationFramework', 'PresentationCore', 'WindowsBase') {
            [void][System.Reflection.Assembly]::LoadWithPartialName($assemblyName)
        }
    } catch {
        $SyncHash.Host.UI.WriteErrorLine("无法加载 WPF 程序集: $_")
        return
    }

    function Initialize-UiHelperTypes {
        if ('ScriptRunner.OutputCapture' -as [type]) {
            return
        }

        Add-Type -TypeDefinition @'
using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace ScriptRunner
{
    public sealed class OutputCapture
    {
        public readonly ConcurrentQueue<string> Lines = new ConcurrentQueue<string>();

        public void OnDataReceived(object sender, DataReceivedEventArgs eventArgs)
        {
            if (eventArgs.Data != null)
            {
                Lines.Enqueue(eventArgs.Data);
            }
        }
    }

    public sealed class ConfigChangeSignal : IDisposable
    {
        private readonly FileSystemWatcher watcher;
        private long lastChangeUtcTicks;

        public long LastChangeUtcTicks
        {
            get { return Interlocked.Read(ref lastChangeUtcTicks); }
        }

        public ConfigChangeSignal(string directory, string fileName)
        {
            watcher = new FileSystemWatcher(directory, fileName);
            watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size;
            watcher.Changed += OnChanged;
            watcher.Created += OnChanged;
            watcher.Deleted += OnChanged;
            watcher.Renamed += OnRenamed;
            watcher.EnableRaisingEvents = true;
        }

        private void OnChanged(object sender, FileSystemEventArgs eventArgs)
        {
            Interlocked.Exchange(ref lastChangeUtcTicks, System.DateTime.UtcNow.Ticks);
        }

        private void OnRenamed(object sender, RenamedEventArgs eventArgs)
        {
            Interlocked.Exchange(ref lastChangeUtcTicks, System.DateTime.UtcNow.Ticks);
        }

        public void Dispose()
        {
            watcher.Dispose();
        }
    }
}
'@ -ErrorAction Stop
    }

    function Import-UiFunctions {
        param([string]$Path, [switch]$Refresh)

        $ErrorActionPreference = 'Stop'
        $source = [System.IO.File]::ReadAllText($Path)
        $tokens = $null
        $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, $Path, [ref]$tokens, [ref]$parseErrors)
        if ($parseErrors.Count -gt 0) {
            throw ($parseErrors | ForEach-Object { $_.ToString() } | Out-String)
        }
        if ($ast.ParamBlock -or $ast.BeginBlock -or $ast.ProcessBlock -or $ast.DynamicParamBlock -or
            $ast.UsingStatements.Count -gt 0 -or -not $ast.EndBlock) {
            throw 'UI 方法文件只能包含函数定义。'
        }

        $newDefinitions = @{}
        foreach ($statement in $ast.EndBlock.Statements) {
            if ($statement -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -or
                $statement.IsFilter -or $statement.IsWorkflow -or
                $statement.Name -notmatch '^[A-Za-z][A-Za-z0-9_-]*$') {
                throw 'UI 方法文件只能包含未指定作用域的函数定义。'
            }
            if ($newDefinitions.ContainsKey($statement.Name)) {
                throw "UI 方法重复: $($statement.Name)"
            }
            if (-not $SyncHash.UiFunctionDefinitions.ContainsKey($statement.Name) -and
                (Get-Command -Name $statement.Name -CommandType Function -ErrorAction SilentlyContinue)) {
                throw "UI 方法与现有函数冲突: $($statement.Name)"
            }
            $newDefinitions[$statement.Name] = $statement.Body.GetScriptBlock()
        }
        if ($newDefinitions.Count -eq 0) {
            throw 'UI 方法文件不包含函数定义。'
        }
        foreach ($name in $SyncHash.UiFunctionDefinitions.Keys) {
            if (-not $newDefinitions.ContainsKey($name)) {
                throw "不能在运行时移除 UI 方法: $name"
            }
        }

        $previousDefinitions = $SyncHash.UiFunctionDefinitions
        if ($Refresh) {
            $previousChildren = @($SyncHash.OperationsPanel.Children)
            $previousScheduledUi = @($SyncHash.ScheduledUi)
            $previousManualUi = @($SyncHash.ManualUi)
            $previousRetiredUi = @($SyncHash.RetiredUi)
            $previousRetiredGroupBox = $SyncHash.RetiredGroupBox
            $previousOneTimeGroups = $SyncHash.OneTimeGroups
            $previousConfig = $SyncHash.Config
        }

        $installedNames = [System.Collections.ArrayList]::new()
        try {
            foreach ($name in $newDefinitions.Keys) {
                Set-Item -Path "Function:\global:$name" -Value $newDefinitions[$name]
                [void]$installedNames.Add($name)
            }
            if ($Refresh) {
                $SyncHash.Config = Read-ValidatedConfig -Path (Join-Path $SyncHash.ScriptRoot 'config.json')
                Update-OperationsUi
            }
            $SyncHash.UiFunctionDefinitions = $newDefinitions
        } catch {
            foreach ($name in $installedNames) {
                if ($previousDefinitions.ContainsKey($name)) {
                    Set-Item -Path "Function:\global:$name" -Value $previousDefinitions[$name]
                } else {
                    Remove-Item -Path "Function:\$name"
                }
            }
            if ($Refresh) {
                $SyncHash.Config = $previousConfig
                $SyncHash.OperationsPanel.Children.Clear()
                foreach ($child in $previousChildren) {
                    [void]$SyncHash.OperationsPanel.Children.Add($child)
                }
                $SyncHash.ScheduledUi.Clear()
                $SyncHash.ManualUi.Clear()
                $SyncHash.RetiredUi.Clear()
                foreach ($entry in $previousScheduledUi) { [void]$SyncHash.ScheduledUi.Add($entry) }
                foreach ($entry in $previousManualUi) { [void]$SyncHash.ManualUi.Add($entry) }
                foreach ($entry in $previousRetiredUi) { [void]$SyncHash.RetiredUi.Add($entry) }
                $SyncHash.RetiredGroupBox = $previousRetiredGroupBox
                $SyncHash.OneTimeGroups = $previousOneTimeGroups
            }
            throw
        }
    }

    function New-MainWindow {
        $window = [System.Windows.Window]::new()
        $window.Title = '脚本执行器'
        $window.SizeToContent = [System.Windows.SizeToContent]::Manual
        $window.Background = [System.Windows.Media.Brushes]::White

        $grid = [System.Windows.Controls.Grid]::new()
        $grid.Margin = [System.Windows.Thickness]::new(10)
        $window.Content = $grid

        $operationsRow = [System.Windows.Controls.RowDefinition]::new()
        $operationsRow.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        [void]$grid.RowDefinitions.Add($operationsRow)
        $SyncHash.OperationsRow = $operationsRow

        $outputRow = [System.Windows.Controls.RowDefinition]::new()
        $outputRow.Height = [System.Windows.GridLength]::Auto
        [void]$grid.RowDefinitions.Add($outputRow)
        $SyncHash.OutputRow = $outputRow

        $operationsPanel = [System.Windows.Controls.StackPanel]::new()
        $SyncHash.OperationsPanel = $operationsPanel

        $operationsScroller = [System.Windows.Controls.ScrollViewer]::new()
        $operationsScroller.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
        $operationsScroller.HorizontalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Disabled
        $operationsScroller.Content = $operationsPanel

        $operationsExpander = [System.Windows.Controls.Expander]::new()
        $operationsExpander.Header = '任务操作'
        $operationsExpander.IsExpanded = $true
        $operationsExpander.Margin = [System.Windows.Thickness]::new(0, 0, 0, 5)
        $operationsExpander.Content = $operationsScroller
        [System.Windows.Controls.Grid]::SetRow($operationsExpander, 0)
        [void]$grid.Children.Add($operationsExpander)
        $SyncHash.OperationsExpander = $operationsExpander
        Update-OperationsUi

        $outputTabs = [System.Windows.Controls.TabControl]::new()
        $outputTabs.Margin = [System.Windows.Thickness]::new(0, 0, 0, 0)
        $outputTabs.Visibility = [System.Windows.Visibility]::Collapsed
        [System.Windows.Controls.Grid]::SetRow($outputTabs, 1)
        [void]$grid.Children.Add($outputTabs)
        $SyncHash.OutputTabs = $outputTabs
        $operationsExpander.Add_Expanded({
            Set-OutputTabsVisible -IsVisible ($SyncHash.OutputTabs.Visibility -eq [System.Windows.Visibility]::Visible)
        })
        $operationsExpander.Add_Collapsed({
            Set-OutputTabsVisible -IsVisible ($SyncHash.OutputTabs.Visibility -eq [System.Windows.Visibility]::Visible)
        })
        $outputTabs.Add_SelectionChanged({
            $selectedTab = $SyncHash.OutputTabs.SelectedItem
            if ($selectedTab -and $selectedTab.Content -is [System.Windows.Controls.TextBox]) {
                $selectedTab.Content.ScrollToEnd()
            }
        })

        return $window
    }

    $uiFunctionsPath = Join-Path $SyncHash.ScriptRoot 'ui\functions.ps1'
    try {
        Import-UiFunctions -Path $uiFunctionsPath
    } catch {
        $SyncHash.Host.UI.WriteErrorLine("加载 UI 方法失败: $_")
        [void][System.Windows.MessageBox]::Show("加载 UI 方法失败: $_", '提示', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # 读取配置文件
    $configPath = Join-Path $SyncHash.ScriptRoot 'config.json'
    if (-not (Test-Path $configPath)) {
        Show-Message "配置文件不存在: $configPath" ([System.Windows.MessageBoxImage]::Error)
        return
    }

    try {
        $SyncHash.Config = Read-ValidatedConfig -Path $configPath
    } catch {
        Show-Message "读取配置文件失败: $_" ([System.Windows.MessageBoxImage]::Error)
        return
    }

    $SyncHash.ScheduledOnStartCallback = {
        param($JobName)
        Invoke-ScheduledOnStart -JobName $JobName
    }

    $SyncHash.ScheduledOnExitCallback = {
        param($JobName)
        Invoke-ScheduledOnExit -JobName $JobName
    }

    $SyncHash.Form = New-MainWindow

    # 添加窗口 Loaded 事件处理程序
    $SyncHash.Form.Add_Loaded({
        foreach ($definition in $SyncHash.Config.longRunning) {
            if ($definition.Item.autoStart) {
                $description = Get-DisplayDescription -Title ([string]$definition.Item.title) -DisplayName ([string]$definition.Item.displayName)
                Start-ScriptJob -JobName $definition.JobName -ScriptPath ([string]$definition.Item.scriptPath) -JobDescription $description -Item $definition.Item -TaskKey $definition.TaskKey -TaskKind Scheduled -OnStart $SyncHash.ScheduledOnStartCallback -OnExit $SyncHash.ScheduledOnExitCallback | Out-Null
            }
        }
    })

    try {
        Initialize-UiHelperTypes
        $configDirectory = [System.IO.Path]::GetDirectoryName($configPath)
        $configFileName = [System.IO.Path]::GetFileName($configPath)
        $watcher = [ScriptRunner.ConfigChangeSignal]::new($configDirectory, $configFileName)
        $SyncHash.ConfigWatcher = $watcher

        $uiFunctionsDirectory = [System.IO.Path]::GetDirectoryName($uiFunctionsPath)
        $uiFunctionsFileName = [System.IO.Path]::GetFileName($uiFunctionsPath)
        $SyncHash.UiFunctionsWatcher = [ScriptRunner.ConfigChangeSignal]::new($uiFunctionsDirectory, $uiFunctionsFileName)
        Import-UiFunctions -Path $uiFunctionsPath -Refresh

        $configRefreshTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $configRefreshTimer.Interval = [TimeSpan]::FromMilliseconds(400)
        $configRefreshTimer.add_Tick({
            $changeTicks = $SyncHash.ConfigWatcher.LastChangeUtcTicks
            if ($changeTicks -le $SyncHash.ConfigLastProcessedUtcTicks) {
                return
            }
            if (([DateTime]::UtcNow.Ticks - $changeTicks) -lt [TimeSpan]::FromMilliseconds(500).Ticks) {
                return
            }

            $SyncHash.ConfigLastProcessedUtcTicks = $changeTicks
            try {
                $newConfig = Read-ValidatedConfig -Path $configPath
                $SyncHash.Config = $newConfig
                Update-OperationsUi
            } catch {
                Show-Message "热更新配置失败，仍保留上一次有效配置: $_" ([System.Windows.MessageBoxImage]::Error)
            }
        })
        $configRefreshTimer.Start()
        $SyncHash.ConfigRefreshTimer = $configRefreshTimer

        $uiFunctionsRefreshTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $uiFunctionsRefreshTimer.Interval = [TimeSpan]::FromMilliseconds(400)
        $uiFunctionsRefreshTimer.add_Tick({
            $changeTicks = $SyncHash.UiFunctionsWatcher.LastChangeUtcTicks
            if ($changeTicks -le $SyncHash.UiFunctionsLastProcessedUtcTicks) {
                return
            }
            if (([DateTime]::UtcNow.Ticks - $changeTicks) -lt [TimeSpan]::FromMilliseconds(500).Ticks) {
                return
            }

            $SyncHash.UiFunctionsLastProcessedUtcTicks = $changeTicks
            try {
                Import-UiFunctions -Path $uiFunctionsPath -Refresh
            } catch {
                Show-Message "热更新 UI 方法失败，仍保留上一次有效方法和操作界面: $_" ([System.Windows.MessageBoxImage]::Error)
            }
        })
        $uiFunctionsRefreshTimer.Start()
        $SyncHash.UiFunctionsRefreshTimer = $uiFunctionsRefreshTimer
    } catch {
        if ($SyncHash.ConfigRefreshTimer) { $SyncHash.ConfigRefreshTimer.Stop() }
        if ($SyncHash.ConfigWatcher) { $SyncHash.ConfigWatcher.Dispose() }
        if ($SyncHash.UiFunctionsWatcher) { $SyncHash.UiFunctionsWatcher.Dispose() }
        Show-Message "无法监听配置或 UI 方法文件变更: $_" ([System.Windows.MessageBoxImage]::Error)
        return
    }

    # 窗口关闭事件
    $SyncHash.Form.Add_Closing({
        if ($SyncHash.ConfigRefreshTimer) {
            $SyncHash.ConfigRefreshTimer.Stop()
        }
        if ($SyncHash.ConfigWatcher) {
            $SyncHash.ConfigWatcher.Dispose()
        }
        if ($SyncHash.UiFunctionsRefreshTimer) {
            $SyncHash.UiFunctionsRefreshTimer.Stop()
        }
        if ($SyncHash.UiFunctionsWatcher) {
            $SyncHash.UiFunctionsWatcher.Dispose()
        }
        Stop-AllTrackedJobs
    })

    # 显示窗口
    $null = $SyncHash.Form.ShowDialog()

    # --- 清理 ---
    Stop-AllTrackedJobs
})

# 启动 UI 线程
$UiPowerShell.Runspace = $UiRunspace
$UiHandle = $UiPowerShell.BeginInvoke()

# 等待 UI 关闭
$UiPowerShell.EndInvoke($UiHandle)

# 清理流
$UiPowerShell.Dispose()
$UiRunspace.Close()
$UiRunspace.Dispose()
