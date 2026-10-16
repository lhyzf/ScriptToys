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

    function Get-OptionalBoolean {
        param(
            [object]$Object,
            [string]$PropertyName,
            [bool]$DefaultValue
        )

        if ($null -eq $Object) {
            return $DefaultValue
        }

        $property = $Object.PSObject.Properties[$PropertyName]
        if ($null -eq $property -or $null -eq $property.Value) {
            return $DefaultValue
        }

        if ($property.Value -isnot [bool]) {
            throw "配置项 '$PropertyName' 必须是 true 或 false。"
        }

        return $property.Value
    }

    function Get-OptionalOutputEncoding {
        param([object]$Object)

        if ($null -eq $Object) {
            return $null
        }

        $property = $Object.PSObject.Properties['outputEncoding']
        if ($null -eq $property -or $null -eq $property.Value) {
            return $null
        }

        if ($property.Value -is [string]) {
            if ([string]::IsNullOrWhiteSpace($property.Value)) {
                return $null
            }

            return $property.Value.Trim()
        }

        if ($property.Value -is [int] -or $property.Value -is [long]) {
            return [long]$property.Value
        }

        throw "配置项 'outputEncoding' 必须是编码名称或代码页编号。"
    }

    function ConvertTo-TextEncoding {
        param([object]$EncodingName)

        if ($EncodingName -is [int] -or $EncodingName -is [long]) {
            return [System.Text.Encoding]::GetEncoding([int]$EncodingName)
        }

        if ($null -ne $EncodingName) {
            if ([string]$EncodingName -match '^\d+$') {
                return [System.Text.Encoding]::GetEncoding([int]$EncodingName)
            }

            return [System.Text.Encoding]::GetEncoding([string]$EncodingName)
        }

        return [System.Text.Encoding]::UTF8
    }

    function Set-OutputTabState {
        param(
            [System.Windows.Controls.TabItem]$Tab,
            [string]$TabLabel,
            [string]$Status,
            [bool]$CanClose
        )

        $Tab.Header.Children[0].Text = "$TabLabel - $Status"
        $Tab.Header.Children[1].IsEnabled = $CanClose
    }

    function New-JobOutputTab {
        param(
            [string]$TabKey,
            [string]$TabLabel
        )

        foreach ($existingTab in $SyncHash.OutputTabs.Items) {
            if ($existingTab.Tag -eq $TabKey) {
                Set-OutputTabState -Tab $existingTab -TabLabel $TabLabel -Status '运行中' -CanClose $false
                $SyncHash.OutputTabs.SelectedItem = $existingTab
                Add-OutputLine -TextBox $existingTab.Content -Line "--- 新一轮运行: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ---"

                return [pscustomobject]@{
                    Tab     = $existingTab
                    TextBox = $existingTab.Content
                    Created = $false
                }
            }
        }

        $textBox = [System.Windows.Controls.TextBox]::new()
        $textBox.IsReadOnly = $true
        $textBox.AcceptsReturn = $true
        $textBox.TextWrapping = [System.Windows.TextWrapping]::NoWrap
        $textBox.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
        $textBox.HorizontalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
        $textBox.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
        $textBox.Margin = [System.Windows.Thickness]::new(4)
        $textBox.Tag = 0

        $header = [System.Windows.Controls.StackPanel]::new()
        $header.Orientation = [System.Windows.Controls.Orientation]::Horizontal
        $headerLabel = [System.Windows.Controls.TextBlock]::new()
        $headerLabel.Text = "$TabLabel - 运行中"
        $headerLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $closeButton = [System.Windows.Controls.Button]::new()
        $closeButton.Content = '×'
        $closeButton.Width = 20
        $closeButton.Height = 20
        $closeButton.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
        $closeButton.Padding = [System.Windows.Thickness]::new(0)
        $closeButton.IsEnabled = $false
        $closeButton.ToolTip = '任务结束或停止后可关闭此标签'
        [void]$header.Children.Add($headerLabel)
        [void]$header.Children.Add($closeButton)

        $tab = [System.Windows.Controls.TabItem]::new()
        $tab.Header = $header
        $tab.Tag = $TabKey
        $tab.Content = $textBox

        $closeHandler = {
            foreach ($jobState in $SyncHash.JobStates.Values) {
                if ([object]::ReferenceEquals($jobState.OutputTab, $tab)) {
                    return
                }
            }

            [void]$SyncHash.OutputTabs.Items.Remove($tab)
            if ($SyncHash.OutputTabs.Items.Count -eq 0) {
                $SyncHash.OutputTabs.Visibility = [System.Windows.Visibility]::Collapsed
            }
        }.GetNewClosure()
        $closeButton.Add_Click($closeHandler)

        [void]$SyncHash.OutputTabs.Items.Add($tab)
        $SyncHash.OutputTabs.Visibility = [System.Windows.Visibility]::Visible
        $SyncHash.OutputTabs.SelectedItem = $tab

        return [pscustomobject]@{
            Tab     = $tab
            TextBox = $textBox
            Created = $true
        }
    }

    function Add-OutputLine {
        param(
            [System.Windows.Controls.TextBox]$TextBox,
            [string]$Line
        )

        $TextBox.AppendText($Line + [Environment]::NewLine)

        $addedLines = 1
        $lineBreakIndex = $Line.IndexOf("`n")
        while ($lineBreakIndex -ge 0) {
            $addedLines++
            $lineBreakIndex = $Line.IndexOf("`n", $lineBreakIndex + 1)
        }

        $lineCount = [int]$TextBox.Tag + $addedLines
        $maxLines = 1000
        $maxCharacters = 50000
        while ($lineCount -gt $maxLines -or $TextBox.Text.Length -gt $maxCharacters) {
            $lineEnd = $TextBox.Text.IndexOf("`n")
            if ($lineEnd -lt 0) {
                $TextBox.Clear()
                $lineCount = 0
                break
            }

            $TextBox.Select(0, $lineEnd + 1)
            $TextBox.SelectedText = ''
            $lineCount--
        }
        $TextBox.Tag = $lineCount

        if ($SyncHash.OutputTabs.SelectedItem -and
            [object]::ReferenceEquals($SyncHash.OutputTabs.SelectedItem.Content, $TextBox)) {
            $TextBox.ScrollToEnd()
        }
    }

    function Initialize-OutputCaptureType {
        if ('ScriptRunner.OutputCapture' -as [type]) {
            return
        }

        Add-Type -TypeDefinition @'
using System.Collections.Concurrent;
using System.Diagnostics;

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
}
'@ -ErrorAction Stop
    }

    function Flush-JobOutput {
        param([hashtable]$JobState)

        if (-not $JobState.OutputCapture) {
            return
        }

        $line = $null
        while ($JobState.OutputCapture.Lines.TryDequeue([ref]$line)) {
            Add-OutputLine -TextBox $JobState.OutputTextBox -Line $line
        }
    }

    function Get-UseBuiltInOutput {
        param([psobject]$Item)

        return (Get-OptionalBoolean -Object $Item -PropertyName 'useBuiltInOutput' -DefaultValue $SyncHash.Config.UseBuiltInOutput)
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
                OutputTabKey = "ManualGroup:$($Group.Name)"
                OutputTabLabel = [string]$Group.Name
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
            [psobject]$Item,
            [string]$OutputTabKey,
            [string]$OutputTabLabel,
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

        $useBuiltInOutput = Get-UseBuiltInOutput -Item $Item
        $outputTab = $null
        $outputTextBox = $null
        $outputTabCreated = $false
        if ($useBuiltInOutput) {
            if ([string]::IsNullOrWhiteSpace($OutputTabKey)) {
                $OutputTabKey = $JobName
            }
            if ([string]::IsNullOrWhiteSpace($OutputTabLabel)) {
                $OutputTabLabel = $JobDescription
            }

            $output = New-JobOutputTab -TabKey $OutputTabKey -TabLabel $OutputTabLabel
            $outputTab = $output.Tab
            $outputTextBox = $output.TextBox
            $outputTabCreated = $output.Created
        }

        $process = $null
        $outputCapture = $null
        $processStarted = $false
        try {
            $extension = [System.IO.Path]::GetExtension($resolvedScriptPath)
            if ($useBuiltInOutput) {
                $processPath = $resolvedScriptPath
                $processArguments = $rawArguments
                $processWorkingDirectory = $null

                if ($extension -ieq '.lnk') {
                    $shell = $null
                    $shortcut = $null
                    try {
                        $shell = New-Object -ComObject WScript.Shell
                        $shortcut = $shell.CreateShortcut($resolvedScriptPath)
                        $processPath = [string]$shortcut.TargetPath
                        $shortcutArguments = [string]$shortcut.Arguments
                        $processWorkingDirectory = [string]$shortcut.WorkingDirectory
                    } finally {
                        if ($shortcut -and [System.Runtime.InteropServices.Marshal]::IsComObject($shortcut)) {
                            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shortcut)
                        }
                        if ($shell -and [System.Runtime.InteropServices.Marshal]::IsComObject($shell)) {
                            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($shell)
                        }
                    }

                    if ([string]::IsNullOrWhiteSpace($processPath)) {
                        throw "快捷方式没有有效的目标程序: $resolvedScriptPath"
                    }
                    if (-not [string]::IsNullOrWhiteSpace($shortcutArguments)) {
                        $processArguments = $shortcutArguments
                        if (-not [string]::IsNullOrWhiteSpace($rawArguments)) {
                            $processArguments += " $rawArguments"
                        }
                    }
                }

                $extension = [System.IO.Path]::GetExtension($processPath)
                Initialize-OutputCaptureType
                $startInfo = [System.Diagnostics.ProcessStartInfo]::new()
                $startInfo.UseShellExecute = $false
                $startInfo.CreateNoWindow = $true
                $startInfo.RedirectStandardOutput = $true
                $startInfo.RedirectStandardError = $true
                if (-not [string]::IsNullOrWhiteSpace($processWorkingDirectory)) {
                    $startInfo.WorkingDirectory = $processWorkingDirectory
                }

                if ($extension -ieq '.ps1') {
                    $startInfo.FileName = 'pwsh.exe'
                    $startInfo.Arguments = "-File `"$processPath`""
                    if (-not [string]::IsNullOrWhiteSpace($processArguments)) {
                        $startInfo.Arguments += " $processArguments"
                    }
                } elseif ($extension -iin '.bat', '.cmd') {
                    $startInfo.FileName = if ([string]::IsNullOrWhiteSpace($env:ComSpec)) { 'cmd.exe' } else { $env:ComSpec }
                    $commandLine = '""{0}"' -f $processPath
                    if (-not [string]::IsNullOrWhiteSpace($processArguments)) {
                        $commandLine += " $processArguments"
                    }
                    $commandLine += '"'
                    $startInfo.Arguments = "/d /s /c $commandLine"
                } else {
                    $startInfo.FileName = $processPath
                    $startInfo.Arguments = $processArguments
                }

                $outputEncodingName = Get-OptionalOutputEncoding -Object $Item
                if ($null -eq $outputEncodingName) {
                    $outputEncodingName = $SyncHash.Config.OutputEncoding
                }
                $outputEncoding = ConvertTo-TextEncoding -EncodingName $outputEncodingName
                $startInfo.StandardOutputEncoding = $outputEncoding
                $startInfo.StandardErrorEncoding = $outputEncoding
                $process = [System.Diagnostics.Process]::new()
                $process.StartInfo = $startInfo
                $outputCapture = [ScriptRunner.OutputCapture]::new()
                $outputHandler = [System.Diagnostics.DataReceivedEventHandler]$outputCapture.OnDataReceived
                $process.add_OutputDataReceived($outputHandler)
                $process.add_ErrorDataReceived($outputHandler)

                if (-not $process.Start()) {
                    throw "无法启动进程 '$processPath'。"
                }
                $processStarted = $true

                $process.BeginOutputReadLine()
                $process.BeginErrorReadLine()
            } elseif ($extension -ieq '.ps1') {
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
            $startupError = $_.Exception.Message
            if ($process -and $useBuiltInOutput) {
                try {
                    if ($processStarted -and -not $process.HasExited) {
                        $process.Kill($true)
                        $process.WaitForExit()
                    }
                } catch {
                    $startupError += "`n清理启动失败的进程时出错: $($_.Exception.Message)"
                } finally {
                    $process.Dispose()
                }
            }
            if ($outputTabCreated) {
                [void]$SyncHash.OutputTabs.Items.Remove($outputTab)
                if ($SyncHash.OutputTabs.Items.Count -eq 0) {
                    $SyncHash.OutputTabs.Visibility = [System.Windows.Visibility]::Collapsed
                }
            } elseif ($outputTab) {
                Set-OutputTabState -Tab $outputTab -TabLabel $OutputTabLabel -Status '启动失败' -CanClose $true
                Add-OutputLine -TextBox $outputTextBox -Line "启动失败: $startupError"
            }
            Show-Message "启动 $JobDescription 失败: $startupError" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        $SyncHash.JobStates[$JobName] = @{
            Process       = $process
            Description   = $JobDescription
            ProcessId     = $process.Id
            Timer         = $null
            OnExit        = $OnExit
            OutputTab     = $outputTab
            OutputTabKey  = $OutputTabKey
            OutputTabLabel = $OutputTabLabel
            OutputTextBox = $outputTextBox
            OutputCapture = $outputCapture
        }

        if ($OnStart) {
            & $OnStart $JobName
        }

        $timer = [System.Windows.Threading.DispatcherTimer]::new()
        if ($useBuiltInOutput) {
            $timer.Interval = [TimeSpan]::FromMilliseconds(100)
        } else {
            $timer.Interval = [TimeSpan]::FromSeconds(1)
        }
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

            Flush-JobOutput -JobState $jobState
            $jobState.Process.Refresh()
            if ($jobState.Process.HasExited) {
                $jobState.Process.WaitForExit()
                Flush-JobOutput -JobState $jobState
                $exitCode = $jobState.Process.ExitCode
                $sender.Stop()
                $SyncHash.JobStates.Remove($jobName)
                if ($jobState.OutputTab) {
                    Set-OutputTabState -Tab $jobState.OutputTab -TabLabel $jobState.OutputTabLabel -Status "已结束 ($exitCode)" -CanClose $true
                }
                $jobState.Process.Dispose()
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
                if (-not $job.Process.HasExited) {
                    $job.Process.Kill($true)
                }
                $job.Process.WaitForExit()
            } elseif ($job.ProcessId) {
                Stop-Process -Id $job.ProcessId -Force -ErrorAction Stop
                Wait-Process -Id $job.ProcessId -ErrorAction SilentlyContinue
            }
        } catch {
            Show-Message "终止 $($job.Description) 时出错: $_" ([System.Windows.MessageBoxImage]::Error)
            return $false
        }

        $SyncHash.JobStates.Remove($JobName)

        Flush-JobOutput -JobState $job
        if ($job.OutputTab) {
            Set-OutputTabState -Tab $job.OutputTab -TabLabel $job.OutputTabLabel -Status '已停止' -CanClose $true
        }

        if ($job.Timer) {
            $job.Timer.Stop()
        }

        if ($job.Process) {
            $job.Process.Dispose()
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
        $outputTabKey = $Entry.OutputTabKey
        $outputTabLabel = $Entry.OutputTabLabel

        $handler = {
            $onStart = { Set-ManualButtonsState -IsEnabled $false }
            $onExit = { Set-ManualButtonsState -IsEnabled $true }

            if (-not (Start-ScriptJob -JobName $jobName -ScriptPath $scriptPath -JobDescription $description -Item $Entry.Item -OutputTabKey $outputTabKey -OutputTabLabel $outputTabLabel -OnStart $onStart -OnExit $onExit)) {
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
        for ($i = 0; $i -le $totalRows; $i++) {
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

        $outputTabs = [System.Windows.Controls.TabControl]::new()
        $outputTabs.Height = 300
        $outputTabs.Margin = [System.Windows.Thickness]::new(0, 5, 0, 0)
        $outputTabs.Visibility = [System.Windows.Visibility]::Collapsed
        [System.Windows.Controls.Grid]::SetRow($outputTabs, $totalRows)
        [void]$grid.Children.Add($outputTabs)
        $SyncHash.OutputTabs = $outputTabs
        $outputTabs.Add_SelectionChanged({
            $selectedTab = $SyncHash.OutputTabs.SelectedItem
            if ($selectedTab -and $selectedTab.Content -is [System.Windows.Controls.TextBox]) {
                $selectedTab.Content.ScrollToEnd()
            }
        })

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
        $globalUseBuiltInOutput = Get-OptionalBoolean -Object $config.ui -PropertyName 'useBuiltInOutput' -DefaultValue $false
        $globalOutputEncoding = Get-OptionalOutputEncoding -Object $config.ui
        if ($null -ne $globalOutputEncoding) {
            $null = ConvertTo-TextEncoding -EncodingName $globalOutputEncoding
        }

        foreach ($item in @($config.longRunning) + @($config.oneTime)) {
            $null = Get-OptionalBoolean -Object $item -PropertyName 'useBuiltInOutput' -DefaultValue $globalUseBuiltInOutput
            $itemOutputEncoding = Get-OptionalOutputEncoding -Object $item
            if ($null -ne $itemOutputEncoding) {
                $null = ConvertTo-TextEncoding -EncodingName $itemOutputEncoding
            }
        }
    } catch {
        Show-Message "读取配置文件失败: $_" ([System.Windows.MessageBoxImage]::Error)
        return
    }

    $SyncHash.Config = [pscustomobject]@{
        UseBuiltInOutput = $globalUseBuiltInOutput
        OutputEncoding   = $globalOutputEncoding
        longRunning      = @($config.longRunning)
        oneTime          = @($config.oneTime)
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
                Start-ScriptJob -JobName $entry.JobName -ScriptPath ([string]$entry.Item.scriptPath) -JobDescription $description -Item $entry.Item -OnStart $ScheduledOnStartCallback -OnExit $ScheduledOnExitCallback | Out-Null
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
            Start-ScriptJob -JobName $jobName -ScriptPath ([string]$scheduledItem.scriptPath) -JobDescription $description -Item $scheduledItem -OnStart $ScheduledOnStartCallback -OnExit $ScheduledOnExitCallback | Out-Null
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
