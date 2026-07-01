#Requires -Version 7

# 确保脚本以 UTF-8 编码保存，以支持中文字符

# 设置变量
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Host = $host
$SyncHash.ScriptRoot = $PSScriptRoot
$SyncHash.JobStates = @{}
$SyncHash.ScheduledUi = [System.Collections.ArrayList]::new()
$SyncHash.ManualUi = [System.Collections.ArrayList]::new()

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

    function Show-Message {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Message,
            [System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Information
        )

        if ($SyncHash.Form) {
            [System.Windows.MessageBox]::Show(
                $SyncHash.Form,
                $Message,
                '提示',
                [System.Windows.MessageBoxButton]::OK,
                $Icon
            ) | Out-Null
            return
        }

        [System.Windows.MessageBox]::Show(
            $Message,
            '提示',
            [System.Windows.MessageBoxButton]::OK,
            $Icon
        ) | Out-Null
    }

    function Resolve-ScriptPath {
        param([string]$Path)

        if ([string]::IsNullOrWhiteSpace($Path)) {
            return $Path
        }

        if ([System.IO.Path]::IsPathRooted($Path)) {
            return $Path
        }

        return Join-Path $SyncHash.ScriptRoot $Path
    }

    function Split-CommandAndArguments {
        param([string]$CommandLine)

        if ([string]::IsNullOrWhiteSpace($CommandLine)) {
            return @{
                CommandPath = $CommandLine
                Arguments   = ''
            }
        }

        $trimmed = $CommandLine.Trim()

        if ($trimmed -match '^"(?<cmd>[^"]+)"\s*(?<args>.*)$') {
            return @{
                CommandPath = $Matches.cmd
                Arguments   = $Matches.args
            }
        }

        if ($trimmed -match "^'(?<cmd>[^']+)'\s*(?<args>.*)$") {
            return @{
                CommandPath = $Matches.cmd
                Arguments   = $Matches.args
            }
        }

        if ($trimmed -match '^(?<cmd>\S+)\s*(?<args>.*)$') {
            return @{
                CommandPath = $Matches.cmd
                Arguments   = $Matches.args
            }
        }

        return @{
            CommandPath = $trimmed
            Arguments   = ''
        }
    }

    function Get-DisplayDescription {
        param(
            [string]$Title,
            [string]$DisplayName
        )

        if ([string]::IsNullOrWhiteSpace($Title)) {
            return $DisplayName
        }

        return $Title
    }

    function Get-OneTimeGroups {
        param([object[]]$Items)

        $groups = [System.Collections.ArrayList]::new()
        $lookup = @{}

        foreach ($item in $Items) {
            $titleKey = [string]$item.title
            if (-not $lookup.ContainsKey($titleKey)) {
                $group = [pscustomobject]@{
                    Name  = $titleKey
                    Items = [System.Collections.ArrayList]::new()
                }
                $null = $groups.Add($group)
                $lookup[$titleKey] = $group
            }

            $null = $lookup[$titleKey].Items.Add($item)
        }

        return $groups
    }

    function New-StandardButton {
        param([string]$Content)

        $button = [System.Windows.Controls.Button]::new()
        $button.Content = $Content
        $button.Height = 25
        $button.Margin = [System.Windows.Thickness]::new(5)
        $button.Padding = [System.Windows.Thickness]::new(5, 2, 5, 2)
        return $button
    }

    function New-GroupPanel {
        $panel = [System.Windows.Controls.StackPanel]::new()
        $panel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $panel.Margin = [System.Windows.Thickness]::new(5)
        return $panel
    }

    function New-GroupBox {
        param(
            [string]$Header,
            [bool]$HasBottomMargin = $true
        )

        $groupBox = [System.Windows.Controls.GroupBox]::new()
        $groupBox.Header = $Header

        if ($HasBottomMargin) {
            $groupBox.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
        }

        return $groupBox
    }

    function New-LongRunningGroup {
        param(
            [psobject]$Item,
            [int]$Index
        )

        $groupBox = New-GroupBox -Header ([string]$Item.title)
        $panel = New-GroupPanel

        $startButton = New-StandardButton -Content ([string]$Item.displayName)
        $stopButton = New-StandardButton -Content ([string]$Item.stopDisplayName)
        $stopButton.IsEnabled = $false

        $statusText = [System.Windows.Controls.TextBlock]::new()
        $statusText.Text = '状态: 已停止'
        $statusText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $statusText.Margin = [System.Windows.Thickness]::new(10, 0, 0, 0)

        [void]$panel.Children.Add($startButton)
        [void]$panel.Children.Add($stopButton)
        [void]$panel.Children.Add($statusText)

        $groupBox.Content = $panel

        $null = $SyncHash.ScheduledUi.Add([pscustomobject]@{
            Index      = $Index
            Item       = $Item
            JobName    = "ScheduledTask$Index"
            StartButton = $startButton
            StopButton  = $stopButton
            StatusText  = $statusText
        })

        return $groupBox
    }

    function New-OneTimeGroup {
        param(
            [psobject]$Group,
            [ref]$ButtonIndex
        )

        $groupBox = New-GroupBox -Header ([string]$Group.Name) -HasBottomMargin $false
        $panel = New-GroupPanel

        foreach ($item in $Group.Items) {
            $currentIndex = $ButtonIndex.Value
            $button = New-StandardButton -Content ([string]$item.displayName)
            [void]$panel.Children.Add($button)

            $null = $SyncHash.ManualUi.Add([pscustomobject]@{
                Index       = $currentIndex
                Item        = $item
                JobName     = "ManualTask$currentIndex"
                Description = [string]$item.displayName
                Button      = $button
            })

            $ButtonIndex.Value = $currentIndex + 1
        }

        $groupBox.Content = $panel
        return $groupBox
    }

    function Set-ManualButtonsState {
        param([bool]$IsEnabled)

        foreach ($entry in $SyncHash.ManualUi) {
            $entry.Button.IsEnabled = $IsEnabled
        }
    }

    function Set-ScheduledState {
        param(
            [int]$Index,
            [bool]$IsRunning
        )

        if ($Index -lt 0 -or $Index -ge $SyncHash.ScheduledUi.Count) {
            return
        }

        $entry = $SyncHash.ScheduledUi[$Index]
        if ($IsRunning) {
            $entry.StartButton.IsEnabled = $false
            $entry.StopButton.IsEnabled = $true
            $entry.StatusText.Text = '状态: 运行中'
            return
        }

        $entry.StartButton.IsEnabled = $true
        $entry.StopButton.IsEnabled = $false
        $entry.StatusText.Text = '状态: 已停止'
    }

    function Get-ScheduledIndexFromJobName {
        param([string]$JobName)

        if ([string]::IsNullOrWhiteSpace($JobName)) {
            return $null
        }

        $match = [regex]::Match($JobName, '\d+$')
        if ($match.Success) {
            return [int]$match.Value
        }

        return $null
    }

    function Invoke-ScheduledOnStart {
        param([string]$JobName)

        $index = Get-ScheduledIndexFromJobName -JobName $JobName
        if ($null -ne $index) {
            Set-ScheduledState -Index $index -IsRunning $true
        }
    }

    function Invoke-ScheduledOnExit {
        param([string]$JobName)

        $index = Get-ScheduledIndexFromJobName -JobName $JobName
        if ($null -ne $index) {
            Set-ScheduledState -Index $index -IsRunning $false
        }
    }

    function Start-ScriptJob {
        param(
            [string]$JobName,
            [string]$ScriptPath,
            [string]$JobDescription,
            [scriptblock]$OnStart,
            [scriptblock]$OnExit
        )

        if ($SyncHash.JobStates.ContainsKey($JobName)) {
            Show-Message "$JobDescription 已在运行。"
            return $false
        }

        $commandSpec = Split-CommandAndArguments -CommandLine $ScriptPath
        $rawCommandPath = $commandSpec.CommandPath
        $rawArguments = $commandSpec.Arguments
        $resolvedScriptPath = Resolve-ScriptPath -Path $rawCommandPath

        if (-not (Test-Path $resolvedScriptPath)) {
            Show-Message "错误: 未找到脚本 '$resolvedScriptPath'" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        try {
            $extension = [System.IO.Path]::GetExtension($resolvedScriptPath)
            if ($extension -ieq '.ps1') {
                # 统一通过 pwsh 运行 PowerShell 脚本，保持行为一致。
                $pwshArguments = "-File `"$resolvedScriptPath`""
                if (-not [string]::IsNullOrWhiteSpace($rawArguments)) {
                    $pwshArguments = "$pwshArguments $rawArguments"
                }
                $process = Start-Process -FilePath 'pwsh.exe' -ArgumentList $pwshArguments -PassThru
            } else {
                if ([string]::IsNullOrWhiteSpace($rawArguments)) {
                    $process = Start-Process -FilePath $resolvedScriptPath -PassThru
                } else {
                    $process = Start-Process -FilePath $resolvedScriptPath -ArgumentList $rawArguments -PassThru
                }
            }
        } catch {
            Show-Message "启动 $JobDescription 失败: $_" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        $SyncHash.JobStates[$JobName] = @{
            Process     = $process
            Description = $JobDescription
            ProcessId   = $process.Id
            Timer       = $null
            OnExit      = $OnExit
        }

        if ($OnStart) {
            & $OnStart $JobName
        }

        $timer = [System.Windows.Threading.DispatcherTimer]::new()
        $timer.Interval = [TimeSpan]::FromSeconds(1)
        $timer | Add-Member -NotePropertyName JobName -NotePropertyValue $JobName
        $timer.add_Tick({
            param($sender, $args)

            $jobName = $sender.JobName
            if (-not $jobName) {
                $sender.Stop()
                return
            }

            $jobState = $SyncHash.JobStates[$jobName]
            if (-not $jobState) {
                $sender.Stop()
                return
            }

            $jobState.Process.Refresh()
            if ($jobState.Process.HasExited) {
                $sender.Stop()
                $SyncHash.JobStates.Remove($jobName)
                if ($jobState.OnExit) {
                    & $jobState.OnExit $jobName
                }
            }
        })
        $timer.Start()
        $SyncHash.JobStates[$JobName].Timer = $timer

        return $true
    }

    function Stop-ScriptJob {
        param([string]$JobName)

        if (-not $SyncHash.JobStates.ContainsKey($JobName)) {
            Show-Message "任务 '$JobName' 未在运行。"
            return $false
        }

        $job = $SyncHash.JobStates[$JobName]
        $onExit = $job.OnExit

        try {
            if ($job.Process) {
                $job.Process.Kill($true)
            } elseif ($job.ProcessId) {
                Stop-Process -Id $job.ProcessId -Force -ErrorAction Stop
                Wait-Process -Id $job.ProcessId -ErrorAction SilentlyContinue
            }
        } catch {
            Show-Message "终止 $($job.Description) 时出错: $_" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        $SyncHash.JobStates.Remove($JobName)

        if ($job.Timer) {
            $job.Timer.Stop()
        }

        if ($onExit) {
            & $onExit $JobName
        }

        return $true
    }

    function Stop-AllTrackedJobs {
        $jobNames = @($SyncHash.JobStates.Keys)
        foreach ($jobName in $jobNames) {
            if ($SyncHash.JobStates.ContainsKey($jobName)) {
                Stop-ScriptJob -JobName $jobName | Out-Null
            }
        }
    }

    function Register-ManualRun {
        param([psobject]$Entry)

        $button = $Entry.Button
        $jobName = $Entry.JobName
        $scriptPath = [string]$Entry.Item.scriptPath
        $description = $Entry.Description

        $handler = {
            $onStart = { Set-ManualButtonsState -IsEnabled $false }
            $onExit = { Set-ManualButtonsState -IsEnabled $true }

            if (-not (Start-ScriptJob -JobName $jobName -ScriptPath $scriptPath -JobDescription $description -OnStart $onStart -OnExit $onExit)) {
                Set-ManualButtonsState -IsEnabled $true
            }
        }.GetNewClosure()

        $button.Add_Click($handler)
    }

    function New-MainWindow {
        $window = [System.Windows.Window]::new()
        $window.Title = '脚本执行器'
        $window.SizeToContent = [System.Windows.SizeToContent]::WidthAndHeight
        $window.Background = [System.Windows.Media.Brushes]::White

        $grid = [System.Windows.Controls.Grid]::new()
        $grid.Margin = [System.Windows.Thickness]::new(10)
        $window.Content = $grid

        $oneTimeGroups = Get-OneTimeGroups -Items $SyncHash.Config.oneTime
        $SyncHash.OneTimeGroups = $oneTimeGroups

        $totalRows = $SyncHash.Config.longRunning.Count + $oneTimeGroups.Count
        for ($i = 0; $i -lt $totalRows; $i++) {
            $rowDefinition = [System.Windows.Controls.RowDefinition]::new()
            $rowDefinition.Height = [System.Windows.GridLength]::Auto
            [void]$grid.RowDefinitions.Add($rowDefinition)
        }

        for ($i = 0; $i -lt $SyncHash.Config.longRunning.Count; $i++) {
            $groupBox = New-LongRunningGroup -Item $SyncHash.Config.longRunning[$i] -Index $i
            [System.Windows.Controls.Grid]::SetRow($groupBox, $i)
            [void]$grid.Children.Add($groupBox)
        }

        $rowIndex = $SyncHash.Config.longRunning.Count
        $buttonIndex = 0
        foreach ($group in $oneTimeGroups) {
            $groupBox = New-OneTimeGroup -Group $group -ButtonIndex ([ref]$buttonIndex)
            [System.Windows.Controls.Grid]::SetRow($groupBox, $rowIndex)
            [void]$grid.Children.Add($groupBox)
            $rowIndex++
        }

        return $window
    }

    # 读取配置文件
    $configPath = Join-Path $SyncHash.ScriptRoot 'config.json'
    if (-not (Test-Path $configPath)) {
        Show-Message "配置文件不存在: $configPath" ([System.Windows.MessageBoxImage]::Error)
        return
    }

    try {
        $config = Get-Content $configPath -Raw | ConvertFrom-Json
    } catch {
        Show-Message "读取配置文件失败: $_" ([System.Windows.MessageBoxImage]::Error)
        return
    }

    $SyncHash.Config = [pscustomobject]@{
        longRunning = @($config.longRunning)
        oneTime     = @($config.oneTime)
    }

    $SyncHash.Form = New-MainWindow

    $ScheduledOnStartCallback = {
        param($JobName)
        Invoke-ScheduledOnStart -JobName $JobName
    }

    $ScheduledOnExitCallback = {
        param($JobName)
        Invoke-ScheduledOnExit -JobName $JobName
    }

    # 添加窗口 Loaded 事件处理程序
    $SyncHash.Form.Add_Loaded({
        foreach ($entry in $SyncHash.ScheduledUi) {
            if ($entry.Item.autoStart) {
                $description = Get-DisplayDescription -Title ([string]$entry.Item.title) -DisplayName ([string]$entry.Item.displayName)
                Start-ScriptJob -JobName $entry.JobName -ScriptPath ([string]$entry.Item.scriptPath) -JobDescription $description -OnStart $ScheduledOnStartCallback -OnExit $ScheduledOnExitCallback | Out-Null
            }
        }
    })

    # 注册定时任务按钮
    foreach ($entry in $SyncHash.ScheduledUi) {
        $index = $entry.Index
        $jobName = $entry.JobName
        $scheduledItem = $entry.Item
        $startButton = $entry.StartButton
        $stopButton = $entry.StopButton

        $startHandler = {
            $description = Get-DisplayDescription -Title ([string]$scheduledItem.title) -DisplayName ([string]$scheduledItem.displayName)
            Start-ScriptJob -JobName $jobName -ScriptPath ([string]$scheduledItem.scriptPath) -JobDescription $description -OnStart $ScheduledOnStartCallback -OnExit $ScheduledOnExitCallback | Out-Null
        }.GetNewClosure()

        $stopHandler = {
            if (Stop-ScriptJob -JobName $jobName) {
                Set-ScheduledState -Index $index -IsRunning $false
            }
        }.GetNewClosure()

        $startButton.Add_Click($startHandler)
        $stopButton.Add_Click($stopHandler)
    }

    # 注册手动执行按钮
    foreach ($entry in $SyncHash.ManualUi) {
        Register-ManualRun -Entry $entry
    }

    # 窗口关闭事件
    $SyncHash.Form.Add_Closing({
        foreach ($entry in $SyncHash.ScheduledUi) {
            if ($SyncHash.JobStates.ContainsKey($entry.JobName)) {
                Stop-ScriptJob -JobName $entry.JobName | Out-Null
            }
        }
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
