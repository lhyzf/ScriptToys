#Requires -Version 7

# 确保脚本以 UTF-8 编码保存，以支持中文字符

# 设置变量
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.Host = $host
$SyncHash.ScriptRoot = $PSScriptRoot

# UI 运行空间
$UiRunspace = [runspacefactory]::CreateRunspace()
$UiRunspace.ApartmentState = 'STA'
$UiRunspace.ThreadOptions = 'ReuseThread'
$UiRunspace.Open()
$UiRunspace.SessionStateProxy.SetVariable('syncHash', $SyncHash)

# UI 脚本
$UiPowerShell = [PowerShell]::Create().AddScript({
    # 设置当前线程的文化，以防万一
    [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'zh-CN'

    # 捕获UI线程中的错误
    trap {
        $SyncHash.host.ui.WriteErrorLine("UI Error: $_`nLine: $($_.InvocationInfo.ScriptLineNumber)`n$($_.InvocationInfo.Line)")
        # 可以在这里添加更详细的错误记录或显示
    }

    function Show-Message {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Message,
            [System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Information
        )
        [System.Windows.MessageBox]::Show($SyncHash.Form, $Message, "提示", [System.Windows.MessageBoxButton]::OK, $Icon) | Out-Null
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

    # --- 核心逻辑函数 ---
    function Start-ScriptJob {
        param(
            [string]$JobName,
            [string]$ScriptPath,
            [string]$JobDescription,
            [scriptblock]$OnStart,
            [scriptblock]$OnExit
        )

        if ($SyncHash.Contains($JobName)) {
            Show-Message "$JobDescription 已在运行。"
            return $false
        }

        $resolvedScriptPath = Resolve-ScriptPath -Path $ScriptPath

        if (-not (Test-Path $resolvedScriptPath)) {
            Show-Message "错误: 未找到脚本 '$resolvedScriptPath'" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        try {
            $arguments = @('-File', "`"$resolvedScriptPath`"")
            $process = Start-Process -FilePath "pwsh.exe" -ArgumentList $arguments -PassThru
        } catch {
            Show-Message "启动 $JobDescription 失败: $_" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        $SyncHash[$JobName] = @{
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
            $jobState = $SyncHash[$jobName]
            if (-not $jobState) {
                $sender.Stop()
                return
            }
            $jobState.Process.Refresh()
            if ($jobState.Process.HasExited) {
                $sender.Stop()
                $SyncHash.Remove($jobName)
                if ($jobState.OnExit) {
                    & $jobState.OnExit $jobName
                }
            }
        })
        $timer.Start()
        $SyncHash[$JobName].Timer = $timer

        return $true
    }

    function Stop-ScriptJob {
        param([string]$JobName)

        if (-not $SyncHash.Contains($JobName)) {
            Show-Message "任务 '$JobName' 未在运行。"
            return $false
        }

        $job = $SyncHash[$JobName]
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

        $SyncHash.Remove($JobName)

        if ($job.Timer) {
            $job.Timer.Stop()
        }

        if ($onExit) {
            & $onExit $JobName
        }

        return $true
    }

    # 读取配置文件
    $configPath = Join-Path $SyncHash.ScriptRoot "config.json"
    if (-not (Test-Path $configPath)) {
        Show-Message "配置文件不存在: $configPath" ([System.Windows.MessageBoxImage]::Error)
        return
    }
    $config = Get-Content $configPath -Raw | ConvertFrom-Json
    $SyncHash.Config = $config

    # 拼接 XAML
    $oneTimeGroups = $config.oneTime | Group-Object -Property title
    $totalRows = $config.longRunning.Count + $oneTimeGroups.Count
    $InputXml = @"
<Window x:Class="ScriptRunner.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="脚本执行器"
        SizeToContent="WidthAndHeight"
        ResizeMode="CanMinimize"
        ThemeMode="Light"
        Background="White">
    <Grid Margin="10">
        <Grid.RowDefinitions>
"@

    for ($i = 0; $i -lt $totalRows; $i++) {
        $InputXml += @"

            <RowDefinition Height="Auto"/>
"@
    }

    $InputXml += @"

        </Grid.RowDefinitions>
"@

    for ($i = 0; $i -lt $config.longRunning.Count; $i++) {
        $item = $config.longRunning[$i]
        $InputXml += @"

        <GroupBox Grid.Row="$i" Header="$($item.title)" Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal" Margin="5">
                <Button x:Name="StartScheduledButton$i" Content="$($item.displayName)" Height="25" Margin="5" Padding="5,2"/>
                <Button x:Name="StopScheduledButton$i" Content="$($item.stopDisplayName)" Height="25" Margin="5" Padding="5,2" IsEnabled="False"/>
                <TextBlock Text="状态: 已停止" x:Name="ScheduledStatusText$i" VerticalAlignment="Center" Margin="10,0"/>
            </StackPanel>
        </GroupBox>
"@
    }

    $rowIndex = $config.longRunning.Count
    $globalButtonIndex = 0
    foreach ($group in $oneTimeGroups) {
        $title = $group.Name
        $InputXml += @"

        <GroupBox Grid.Row="$rowIndex" Header="$title">
            <StackPanel Orientation="Horizontal" Margin="5">
"@

        foreach ($item in $group.Group) {
            $buttonName = "RunButton$globalButtonIndex"
            $InputXml += @"

                <Button x:Name="$buttonName" Content="$($item.displayName)" Height="25" Margin="5" Padding="5,2"/>
"@
            $globalButtonIndex++
        }

        $InputXml += @"

            </StackPanel>
        </GroupBox>
"@
        $rowIndex++
    }

    $InputXml += @"

    </Grid>
</Window>
"@

    # 清理 XAML 兼容性问题
    $InputXml = $InputXml -replace 'mc:Ignorable="d"', '' -replace 'x:N', 'N' -replace '^<Win.*', '<Window'

    # 加载 XAML
    try {
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [xml]$Xaml = $InputXml
        $XamlReader = (New-Object System.Xml.XmlNodeReader($Xaml))
        $SyncHash.Form = [Windows.Markup.XamlReader]::Load($XamlReader)
    } catch {
        $SyncHash.host.ui.WriteErrorLine("无法加载 XAML. 请检查 .NET Framework 是否已安装以及 XAML 语法。")
        return
    }

    # 将 XAML 控件加载到 SyncHash 中
    $Xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $SyncHash.Add($_.Name, $SyncHash.Form.FindName($_.Name))
    }

    function Set-ManualButtonsState {
        param([bool]$IsEnabled)
        for ($i = 0; $i -lt $SyncHash.Config.oneTime.Count; $i++) {
            $SyncHash["RunButton$i"].IsEnabled = $IsEnabled
        }
    }

    function Set-ScheduledState {
        param([int]$Index, [bool]$IsRunning)
        if ($IsRunning) {
            $SyncHash["StartScheduledButton$Index"].IsEnabled = $false
            $SyncHash["StopScheduledButton$Index"].IsEnabled = $true
            $SyncHash["ScheduledStatusText$Index"].Text = "状态: 运行中"
        } else {
            $SyncHash["StartScheduledButton$Index"].IsEnabled = $true
            $SyncHash["StopScheduledButton$Index"].IsEnabled = $false
            $SyncHash["ScheduledStatusText$Index"].Text = "状态: 已停止"
        }
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

    $ScheduledOnStartCallback = {
        param($JobName)
        Invoke-ScheduledOnStart -JobName $JobName
    }

    $ScheduledOnExitCallback = {
        param($JobName)
        Invoke-ScheduledOnExit -JobName $JobName
    }

    function Register-ManualRun {
        param(
            [System.Windows.Controls.Button]$Button,
            [string]$ScriptPath,
            [string]$JobName,
            [string]$Description
        )

        $handler = {
            $onStart = { Set-ManualButtonsState -IsEnabled $false }
            $onExit = { Set-ManualButtonsState -IsEnabled $true }

            if (-not (Start-ScriptJob -JobName $JobName -ScriptPath $ScriptPath -JobDescription $Description -OnStart $onStart -OnExit $onExit)) {
                Set-ManualButtonsState -IsEnabled $true
            }
        }.GetNewClosure()

        $Button.Add_Click($handler)
    }

    # 添加窗口 Loaded 事件处理程序
    $SyncHash.Form.Add_Loaded({
        for ($i = 0; $i -lt $SyncHash.Config.longRunning.Count; $i++) {
            $jobName = "ScheduledTask$i"
            $item = $SyncHash.Config.longRunning[$i]
            if ($item.autoStart) {
                $scriptPath = $item.scriptPath
                $description = if ([string]::IsNullOrWhiteSpace($item.title)) { $item.displayName } else { $item.title }
                Start-ScriptJob -JobName $jobName -ScriptPath $scriptPath -JobDescription $description -OnStart $ScheduledOnStartCallback -OnExit $ScheduledOnExitCallback | Out-Null
            }
        }
    })

    # --- 事件处理程序 ---

    # 注册定时任务按钮
    for ($i = 0; $i -lt $config.longRunning.Count; $i++) {
        $indexCopy = $i
        $jobName = "ScheduledTask$indexCopy"
        $startButton = $SyncHash["StartScheduledButton$indexCopy"]
        $stopButton = $SyncHash["StopScheduledButton$indexCopy"]

        $startHandlerJobName = $jobName
        $startHandlerIndex = $indexCopy

        $startHandler = {
            $jobNameLocal = $startHandlerJobName
            $indexLocal = $startHandlerIndex
            $item = $SyncHash.Config.longRunning[$indexLocal]
            $scriptPath = $item.scriptPath
            $description = if ([string]::IsNullOrWhiteSpace($item.title)) { $item.displayName } else { $item.title }
            Start-ScriptJob -JobName $jobNameLocal -ScriptPath $scriptPath -JobDescription $description -OnStart $ScheduledOnStartCallback -OnExit $ScheduledOnExitCallback | Out-Null
        }.GetNewClosure()

        $stopHandlerJobName = $jobName
        $stopHandlerIndex = $indexCopy

        $stopHandler = {
            $jobNameLocal = $stopHandlerJobName
            $indexLocal = $stopHandlerIndex
            if (Stop-ScriptJob -JobName $jobNameLocal) {
                Set-ScheduledState -Index $indexLocal -IsRunning $false
            }
        }.GetNewClosure()

        $startButton.Add_Click($startHandler)
        $stopButton.Add_Click($stopHandler)
    }

    # 注册手动执行按钮
    $buttonIndex = 0
    foreach ($item in $config.oneTime) {
        $buttonName = "RunButton$buttonIndex"
        Register-ManualRun -Button $SyncHash[$buttonName] -ScriptPath $item.scriptPath -JobName "Manual$item.displayName" -Description $item.displayName
        $buttonIndex++
    }

    # 窗口关闭事件
    $SyncHash.Form.Add_Closing({
        # 停止仍在运行的任何作业
        for ($j = 0; $j -lt $SyncHash.Config.longRunning.Count; $j++) {
            if ($SyncHash.Contains("ScheduledTask$j")) {
                Stop-ScriptJob -JobName "ScheduledTask$j"
            }
        }
    })

    # 显示窗口
    $null = $SyncHash.Form.ShowDialog()

    # --- 清理 ---
    # 确保所有作业都已停止
    $runningJobs = $SyncHash.Keys | Where-Object {
        $value = $SyncHash[$_]
        $value -is [hashtable] -and $value.ContainsKey('ProcessId')
    }
    foreach ($jobName in $runningJobs) {
        Stop-ScriptJob -JobName $jobName | Out-Null
    }
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
