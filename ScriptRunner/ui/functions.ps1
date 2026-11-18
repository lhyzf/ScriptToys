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

function Set-OutputTabsVisible {
    param([bool]$IsVisible)

    if ($IsVisible) {
        if ($SyncHash.OperationsExpander.IsExpanded) {
            $SyncHash.OperationsRow.Height = [System.Windows.GridLength]::new(2, [System.Windows.GridUnitType]::Star)
            $SyncHash.PanelSplitter.Visibility = [System.Windows.Visibility]::Visible
        } else {
            $SyncHash.OperationsRow.Height = [System.Windows.GridLength]::Auto
            $SyncHash.PanelSplitter.Visibility = [System.Windows.Visibility]::Collapsed
        }
        $SyncHash.SplitterRow.Height = [System.Windows.GridLength]::Auto
        $SyncHash.OutputRow.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $SyncHash.OutputTabs.Visibility = [System.Windows.Visibility]::Visible
        return
    }

    $SyncHash.OutputTabs.Visibility = [System.Windows.Visibility]::Collapsed
    $SyncHash.PanelSplitter.Visibility = [System.Windows.Visibility]::Collapsed
    $SyncHash.SplitterRow.Height = [System.Windows.GridLength]::Auto
    $SyncHash.OutputRow.Height = [System.Windows.GridLength]::Auto
    $SyncHash.OperationsRow.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
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
            Set-OutputTabsVisible -IsVisible $true
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
            Set-OutputTabsVisible -IsVisible $false
        }
    }.GetNewClosure()
    $closeButton.Add_Click($closeHandler)

    [void]$SyncHash.OutputTabs.Items.Add($tab)
    Set-OutputTabsVisible -IsVisible $true
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

    $Line = [regex]::Replace($Line, '\x1B\[[0-9;:]*m', '')
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

    foreach ($definition in $Items) {
        $titleKey = [string]$definition.Item.title
        if (-not $lookup.ContainsKey($titleKey)) {
            $group = [pscustomobject]@{
                Name  = $titleKey
                Items = [System.Collections.ArrayList]::new()
            }
            $null = $groups.Add($group)
            $lookup[$titleKey] = $group
        }

        $null = $lookup[$titleKey].Items.Add($definition)
    }

    return $groups
}

function New-TaskDefinitions {
    param(
        [object[]]$Items,
        [ValidateSet('Scheduled', 'Manual')]
        [string]$Kind,
        [hashtable]$SeenExplicitIds,
        [hashtable]$SeenFallbackIds
    )

    $definitions = [System.Collections.ArrayList]::new()
    foreach ($item in $Items) {
        if ($null -eq $item -or $item -isnot [pscustomobject]) {
            throw "$Kind 任务必须是 JSON 对象。"
        }

        $idProperty = $item.PSObject.Properties['id']
        $explicitId = if ($idProperty) { ([string]$idProperty.Value).Trim() } else { '' }
        if (-not [string]::IsNullOrWhiteSpace($explicitId)) {
            $identity = "$Kind`nid:$explicitId"
            if ($SeenExplicitIds.ContainsKey($identity)) {
                throw "$Kind 任务的 id '$explicitId' 重复。"
            }
            $SeenExplicitIds[$identity] = $true
        } else {
            $identityData = [ordered]@{
                Kind           = $Kind
                Title          = [string]$item.title
                DisplayName    = [string]$item.displayName
                StopDisplayName = [string]$item.stopDisplayName
                ScriptPath     = [string]$item.scriptPath
            }
            $identity = ConvertTo-Json -InputObject $identityData -Compress
            $occurrence = 0
            if ($SeenFallbackIds.ContainsKey($identity)) {
                $occurrence = [int]$SeenFallbackIds[$identity]
            }
            $SeenFallbackIds[$identity] = $occurrence + 1
            $identity = "$identity`n$occurrence"
        }

        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        try {
            $digest = [Convert]::ToHexString(
                $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($identity))
            ).Substring(0, 20)
        } finally {
            $sha256.Dispose()
        }

        $jobPrefix = if ($Kind -eq 'Scheduled') { 'ScheduledTask' } else { 'ManualTask' }
        $null = $definitions.Add([pscustomobject]@{
            Item    = $item
            TaskKey = "$Kind`:$digest"
            JobName = "${jobPrefix}_$digest"
        })
    }

    return ,$definitions
}

function Read-ValidatedConfig {
    param([string]$Path)

    $config = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    if ($null -eq $config -or $config -isnot [pscustomobject]) {
        throw '配置文件根节点必须是 JSON 对象。'
    }

    $globalUseBuiltInOutput = Get-OptionalBoolean -Object $config.ui -PropertyName 'useBuiltInOutput' -DefaultValue $false
    $globalOutputEncoding = Get-OptionalOutputEncoding -Object $config.ui
    if ($null -ne $globalOutputEncoding) {
        $null = ConvertTo-TextEncoding -EncodingName $globalOutputEncoding
    }

    $longRunningItems = @()
    $oneTimeItems = @()
    if ($null -ne $config.longRunning) {
        $longRunningItems = @($config.longRunning)
    }
    if ($null -ne $config.oneTime) {
        $oneTimeItems = @($config.oneTime)
    }
    foreach ($item in $longRunningItems + $oneTimeItems) {
        if ($null -eq $item -or $item -isnot [pscustomobject]) {
            throw '每个任务配置都必须是 JSON 对象。'
        }
        $null = Get-OptionalBoolean -Object $item -PropertyName 'useBuiltInOutput' -DefaultValue $globalUseBuiltInOutput
        $itemOutputEncoding = Get-OptionalOutputEncoding -Object $item
        if ($null -ne $itemOutputEncoding) {
            $null = ConvertTo-TextEncoding -EncodingName $itemOutputEncoding
        }
    }

    $seenExplicitIds = @{}
    $seenFallbackIds = @{}
    return [pscustomobject]@{
        UseBuiltInOutput = $globalUseBuiltInOutput
        OutputEncoding   = $globalOutputEncoding
        longRunning      = New-TaskDefinitions -Items $longRunningItems -Kind Scheduled -SeenExplicitIds $seenExplicitIds -SeenFallbackIds $seenFallbackIds
        oneTime          = New-TaskDefinitions -Items $oneTimeItems -Kind Manual -SeenExplicitIds $seenExplicitIds -SeenFallbackIds $seenFallbackIds
    }
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
        [psobject]$Definition
    )

    $item = $Definition.Item
    $groupBox = New-GroupBox -Header ([string]$Item.title)
    $panel = New-GroupPanel

    $toggleButton = New-ScheduledToggleButton -Content ([string]$Item.displayName)

    $statusText = [System.Windows.Controls.TextBlock]::new()
    $statusText.Text = '状态: 已停止'
    $statusText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $statusText.Margin = [System.Windows.Thickness]::new(10, 0, 0, 0)

    [void]$panel.Children.Add($toggleButton)
    [void]$panel.Children.Add($statusText)

    $groupBox.Content = $panel

    $null = $SyncHash.ScheduledUi.Add([pscustomobject]@{
        TaskKey     = $Definition.TaskKey
        Item        = $item
        JobName     = $Definition.JobName
        ToggleButton = $toggleButton
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
        $button = New-StandardButton -Content ([string]$item.Item.displayName)
        [void]$panel.Children.Add($button)

        $definition = $item
        $null = $SyncHash.ManualUi.Add([pscustomobject]@{
            Index       = $currentIndex
            TaskKey     = $definition.TaskKey
            Item        = $definition.Item
            JobName     = $definition.JobName
            Description = [string]$definition.Item.displayName
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
        [string]$JobName,
        [bool]$IsRunning
    )

    foreach ($entry in $SyncHash.ScheduledUi) {
        if ($entry.JobName -eq $JobName) {
            if ($IsRunning) {
                $entry.ToggleButton.Content = [string]$entry.Item.stopDisplayName
            } else {
                $entry.ToggleButton.Content = [string]$entry.Item.displayName
            }
            Set-ScheduledToggleVisualState -Button $entry.ToggleButton -IsRunning $IsRunning
            $entry.StatusText.Text = if ($IsRunning) { '状态: 运行中' } else { '状态: 已停止' }
            break
        }
    }
    foreach ($entry in $SyncHash.RetiredUi) {
        if ($entry.JobName -eq $JobName) {
            $record = $SyncHash.JobRecords[$JobName]
            $entry.StatusText.Text = if ($IsRunning) { '状态: 运行中' } elseif ($record) { $record.State } else { '状态: 已停止' }
            $entry.StopButton.IsEnabled = $IsRunning
            break
        }
    }
}

function Invoke-ScheduledOnStart {
    param([string]$JobName)

    Set-ScheduledState -JobName $JobName -IsRunning $true
}

function Invoke-ScheduledOnExit {
    param([string]$JobName)

    Set-ScheduledState -JobName $JobName -IsRunning $false
}

function Set-JobRecordState {
    param(
        [string]$JobName,
        [string]$State
    )

    if ($SyncHash.JobRecords.ContainsKey($JobName)) {
        $SyncHash.JobRecords[$JobName].State = $State
    }
    foreach ($entry in $SyncHash.RetiredUi) {
        if ($entry.JobName -eq $JobName) {
            $entry.StatusText.Text = $State
            $entry.StopButton.IsEnabled = $SyncHash.JobStates.ContainsKey($JobName)
            break
        }
    }
    if ($SyncHash.OperationsPanel) {
        Update-RetiredJobsUi
    }
}

function Start-ScriptJob {
    param(
        [string]$JobName,
        [string]$ScriptPath,
        [string]$JobDescription,
        [psobject]$Item,
        [string]$TaskKey,
        [ValidateSet('Scheduled', 'Manual')]
        [string]$TaskKind,
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
            Initialize-UiHelperTypes
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
                Set-OutputTabsVisible -IsVisible $false
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
        TaskKey        = $TaskKey
        TaskKind       = $TaskKind
        Timer          = $null
        OnExit        = $OnExit
        OutputTab     = $outputTab
        OutputTabKey  = $OutputTabKey
        OutputTabLabel = $OutputTabLabel
        OutputTextBox = $outputTextBox
        OutputCapture = $outputCapture
    }
    $SyncHash.JobRecords[$JobName] = @{
        JobName      = $JobName
        TaskKey      = $TaskKey
        TaskKind     = $TaskKind
        Description  = $JobDescription
        Item         = $Item
        ProcessId    = $process.Id
        StartedAt    = Get-Date
        State        = '状态: 运行中'
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
            Set-JobRecordState -JobName $jobName -State "状态: 已结束 ($exitCode)"
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
    Set-JobRecordState -JobName $JobName -State '状态: 已停止'

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

        if (-not (Start-ScriptJob -JobName $jobName -ScriptPath $scriptPath -JobDescription $description -Item $Entry.Item -TaskKey $Entry.TaskKey -TaskKind Manual -OutputTabKey $outputTabKey -OutputTabLabel $outputTabLabel -OnStart $onStart -OnExit $onExit)) {
            Set-ManualButtonsState -IsEnabled $true
        }
    }.GetNewClosure()

    $button.Add_Click($handler)
}

function Get-ConfiguredTaskKeys {
    $keys = @{}
    foreach ($definition in @($SyncHash.Config.longRunning) + @($SyncHash.Config.oneTime)) {
        $keys[$definition.TaskKey] = $true
    }
    return $keys
}

function Update-RetiredJobsUi {
    if ($SyncHash.RetiredGroupBox) {
        [void]$SyncHash.OperationsPanel.Children.Remove($SyncHash.RetiredGroupBox)
        $SyncHash.RetiredGroupBox = $null
    }
    $SyncHash.RetiredUi.Clear()
    $configuredKeys = Get-ConfiguredTaskKeys
    $retiredRecords = @(
        $SyncHash.JobRecords.Values |
            Where-Object { -not $configuredKeys.ContainsKey($_.TaskKey) -and $SyncHash.JobStates.ContainsKey($_.JobName) } |
            Sort-Object StartedAt
    )
    if ($retiredRecords.Count -eq 0) {
        return
    }

    $groupBox = New-GroupBox -Header '配置中已移除的任务'
    $panel = [System.Windows.Controls.StackPanel]::new()
    foreach ($record in $retiredRecords) {
        $row = New-GroupPanel
        $description = [System.Windows.Controls.TextBlock]::new()
        $description.Text = "$($record.Description) (PID: $($record.ProcessId))"
        $description.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $statusText = [System.Windows.Controls.TextBlock]::new()
        $statusText.Text = $record.State
        $statusText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $statusText.Margin = [System.Windows.Thickness]::new(10, 0, 0, 0)
        $stopButton = New-StandardButton -Content '停止'
        $stopButton.IsEnabled = $SyncHash.JobStates.ContainsKey($record.JobName)
        [void]$row.Children.Add($description)
        [void]$row.Children.Add($statusText)
        [void]$row.Children.Add($stopButton)
        [void]$panel.Children.Add($row)

        $entry = [pscustomobject]@{
            JobName    = $record.JobName
            StatusText = $statusText
            StopButton = $stopButton
        }
        $null = $SyncHash.RetiredUi.Add($entry)
        $jobName = $record.JobName
        $stopButton.Add_Click({
            Stop-ScriptJob -JobName $jobName | Out-Null
        }.GetNewClosure())
    }
    $groupBox.Content = $panel
    [void]$SyncHash.OperationsPanel.Children.Add($groupBox)
    $SyncHash.RetiredGroupBox = $groupBox
}

function Register-ScheduledTaskHandlers {
    foreach ($entry in $SyncHash.ScheduledUi) {
        $scheduledEntry = $entry
        $jobName = $scheduledEntry.JobName
        $scheduledItem = $scheduledEntry.Item

        $toggleHandler = {
            if ($SyncHash.JobStates.ContainsKey($jobName)) {
                Stop-ScriptJob -JobName $jobName | Out-Null
            } else {
                $description = Get-DisplayDescription -Title ([string]$scheduledItem.title) -DisplayName ([string]$scheduledItem.displayName)
                Start-ScriptJob -JobName $jobName -ScriptPath ([string]$scheduledItem.scriptPath) -JobDescription $description -Item $scheduledItem -TaskKey $scheduledEntry.TaskKey -TaskKind Scheduled -OnStart $SyncHash.ScheduledOnStartCallback -OnExit $SyncHash.ScheduledOnExitCallback | Out-Null
            }
        }.GetNewClosure()

        $scheduledEntry.ToggleButton.Add_Click($toggleHandler)
    }
}

function Update-OperationsUi {
    $SyncHash.ScheduledUi.Clear()
    $SyncHash.ManualUi.Clear()
    $SyncHash.OperationsPanel.Children.Clear()

    foreach ($definition in $SyncHash.Config.longRunning) {
        $groupBox = New-LongRunningGroup -Definition $definition
        [void]$SyncHash.OperationsPanel.Children.Add($groupBox)
        $running = $SyncHash.JobStates.ContainsKey($definition.JobName)
        Set-ScheduledState -JobName $definition.JobName -IsRunning $running
    }

    $oneTimeGroups = Get-OneTimeGroups -Items $SyncHash.Config.oneTime
    $SyncHash.OneTimeGroups = $oneTimeGroups
    $buttonIndex = 0
    foreach ($group in $oneTimeGroups) {
        $groupBox = New-OneTimeGroup -Group $group -ButtonIndex ([ref]$buttonIndex)
        [void]$SyncHash.OperationsPanel.Children.Add($groupBox)
    }

    Register-ScheduledTaskHandlers
    foreach ($entry in $SyncHash.ManualUi) {
        Register-ManualRun -Entry $entry
    }
    $hasRunningManualJob = @($SyncHash.JobStates.Values | Where-Object TaskKind -eq 'Manual').Count -gt 0
    Set-ManualButtonsState -IsEnabled (-not $hasRunningManualJob)
    Update-RetiredJobsUi
}
