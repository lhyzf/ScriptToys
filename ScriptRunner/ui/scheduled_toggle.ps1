function Get-ScheduledTogglePalette {
    param([System.Windows.Controls.Button]$Button)

    $foreground = $Button.Foreground
    if ($foreground -isnot [System.Windows.Media.SolidColorBrush]) {
        throw '无法从按钮前景色判断系统主题。'
    }

    $color = $foreground.Color
    $red = [double]$color.R / 255
    $green = [double]$color.G / 255
    $blue = [double]$color.B / 255
    $red = if ($red -le 0.04045) { $red / 12.92 } else { [Math]::Pow(($red + 0.055) / 1.055, 2.4) }
    $green = if ($green -le 0.04045) { $green / 12.92 } else { [Math]::Pow(($green + 0.055) / 1.055, 2.4) }
    $blue = if ($blue -le 0.04045) { $blue / 12.92 } else { [Math]::Pow(($blue + 0.055) / 1.055, 2.4) }
    $luminance = (0.2126 * $red) + (0.7152 * $green) + (0.0722 * $blue)

    if ($luminance -ge 0.5) {
        return @{
            Stopped = '#2E6B45'
            StoppedHover = '#3A8054'
            StoppedPressed = '#245638'
            Running = '#7A3535'
            RunningHover = '#924646'
            RunningPressed = '#682C2C'
        }
    }

    return @{
        Stopped = '#DCEBDD'
        StoppedHover = '#C6DFC9'
        StoppedPressed = '#B0D1B5'
        Running = '#F0DDDD'
        RunningHover = '#E4C5C5'
        RunningPressed = '#D6AAAA'
    }
}

function Set-ScheduledToggleVisualState {
    param(
        [System.Windows.Controls.Button]$Button,
        [bool]$IsRunning
    )

    $Button.Tag = if ($IsRunning) { 'Running' } else { 'Stopped' }
    $palette = Get-ScheduledTogglePalette -Button $Button
    $state = if ($IsRunning) { 'Running' } else { 'Stopped' }
    $Button.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($palette[$state])
    $Button.Resources['ScriptRunner.Toggle.HoverBackground'] = [System.Windows.Media.BrushConverter]::new().ConvertFromString($palette["${state}Hover"])
    $Button.Resources['ScriptRunner.Toggle.PressedBackground'] = [System.Windows.Media.BrushConverter]::new().ConvertFromString($palette["${state}Pressed"])
    Set-ScheduledToggleBackground -Button $Button
}

function Set-ScheduledToggleBackground {
    param(
        [System.Windows.Controls.Button]$Button,
        [ValidateSet('Auto', 'Hover', 'Pressed')]
        [string]$InteractionState = 'Auto'
    )

    $border = $Button.Resources['ScriptRunner.Toggle.BackgroundElement']
    if (-not $border) {
        return
    }

    $stateKey = switch ($InteractionState) {
        'Pressed' { 'ScriptRunner.Toggle.PressedBackground'; break }
        'Hover' { 'ScriptRunner.Toggle.HoverBackground'; break }
        default {
            if ($Button.IsPressed) {
                'ScriptRunner.Toggle.PressedBackground'
            } elseif ($Button.IsMouseOver) {
                'ScriptRunner.Toggle.HoverBackground'
            }
        }
    }

    if ($stateKey) {
        $border.Background = $Button.Resources[$stateKey]
    } else {
        $border.Background = $Button.Background
    }
}

function Initialize-ScheduledToggleBackground {
    param([System.Windows.Controls.Button]$Button)

    $Button.ApplyTemplate()
    $border = $null
    foreach ($trigger in $Button.Template.Triggers) {
        foreach ($setter in $trigger.Setters) {
            if ($setter.Property -eq [System.Windows.Controls.Border]::BackgroundProperty -and $setter.TargetName) {
                $border = $Button.Template.FindName($setter.TargetName, $Button)
                if ($border -is [System.Windows.Controls.Border]) {
                    break
                }
            }
        }
        if ($border) { break }
    }

    if (-not $border) {
        $pending = [System.Collections.Stack]::new()
        $pending.Push($Button)
        while ($pending.Count -gt 0 -and -not $border) {
            $current = $pending.Pop()
            $childCount = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($current)
            for ($index = $childCount - 1; $index -ge 0; $index--) {
                $child = [System.Windows.Media.VisualTreeHelper]::GetChild($current, $index)
                if ($child -is [System.Windows.Controls.Border]) {
                    $border = $child
                    break
                }
                $pending.Push($child)
            }
        }
    }

    if (-not $border) {
        throw '无法找到系统按钮模板的背景元素。'
    }

    $Button.Resources['ScriptRunner.Toggle.BackgroundElement'] = $border
    $listenersKey = 'ScriptRunner.Toggle.StateListeners'
    if ($Button.Resources.Contains($listenersKey)) {
        foreach ($listener in $Button.Resources[$listenersKey]) {
            $listener.Descriptor.RemoveValueChanged($Button, $listener.Handler)
        }
        [void]$Button.Resources.Remove($listenersKey)
    }

    $listeners = [System.Collections.ArrayList]::new()
    foreach ($property in @(
        [System.Windows.UIElement]::IsMouseOverProperty,
        [System.Windows.Controls.Primitives.ButtonBase]::IsPressedProperty,
        [System.Windows.Controls.Control]::ForegroundProperty
    )) {
        $descriptor = [System.ComponentModel.DependencyPropertyDescriptor]::FromProperty($property, $Button.GetType())
        $buttonForState = $Button
        $propertyForState = $property
        $handler = [System.EventHandler]{
            if ($propertyForState -eq [System.Windows.Controls.Control]::ForegroundProperty) {
                Set-ScheduledToggleVisualState -Button $buttonForState -IsRunning ($buttonForState.Tag -eq 'Running')
            } else {
                Set-ScheduledToggleBackground -Button $buttonForState
            }
        }.GetNewClosure()
        $descriptor.AddValueChanged($Button, $handler)
        [void]$listeners.Add([pscustomobject]@{
            Descriptor = $descriptor
            Handler    = $handler
        })
    }
    $Button.Resources[$listenersKey] = $listeners
    Set-ScheduledToggleVisualState -Button $Button -IsRunning ($Button.Tag -eq 'Running')

    $unloadedHandlerKey = 'ScriptRunner.Toggle.UnloadedHandler'
    if (-not $Button.Resources.Contains($unloadedHandlerKey)) {
        $buttonForUnload = $Button
        $Button.Add_Unloaded({
            $unloadedListenersKey = 'ScriptRunner.Toggle.StateListeners'
            if ($buttonForUnload.Resources.Contains($unloadedListenersKey)) {
                foreach ($listener in $buttonForUnload.Resources[$unloadedListenersKey]) {
                    $listener.Descriptor.RemoveValueChanged($buttonForUnload, $listener.Handler)
                }
                [void]$buttonForUnload.Resources.Remove($unloadedListenersKey)
            }
            [void]$buttonForUnload.Resources.Remove('ScriptRunner.Toggle.BackgroundElement')
        }.GetNewClosure())
        $Button.Resources[$unloadedHandlerKey] = $true
    }
}

function New-ScheduledToggleButton {
    param([string]$Content)

    $button = New-StandardButton -Content $Content
    Set-ScheduledToggleVisualState -Button $button -IsRunning $false
    $buttonForLoaded = $button
    $button.Add_Loaded({
        try {
            Initialize-ScheduledToggleBackground -Button $buttonForLoaded
        } catch {
            Show-Message "无法自定义任务按钮背景: $_" ([System.Windows.MessageBoxImage]::Error)
        }
    }.GetNewClosure())

    return $button
}
