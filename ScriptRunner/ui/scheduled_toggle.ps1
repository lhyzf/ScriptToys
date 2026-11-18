function Set-ScheduledToggleVisualState {
    param(
        [System.Windows.Controls.Button]$Button,
        [bool]$IsRunning
    )

    $Button.Tag = if ($IsRunning) { 'Running' } else { 'Stopped' }
    if ($IsRunning) {
        $Button.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#F0DDDD')
        $Button.Resources['ScriptRunner.Toggle.HoverBackground'] = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#E4C5C5')
        $Button.Resources['ScriptRunner.Toggle.PressedBackground'] = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D6AAAA')
    } else {
        $Button.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#DCEBDD')
        $Button.Resources['ScriptRunner.Toggle.HoverBackground'] = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#C6DFC9')
        $Button.Resources['ScriptRunner.Toggle.PressedBackground'] = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#B0D1B5')
    }
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
        [System.Windows.Controls.Primitives.ButtonBase]::IsPressedProperty
    )) {
        $descriptor = [System.ComponentModel.DependencyPropertyDescriptor]::FromProperty($property, $Button.GetType())
        $buttonForState = $Button
        $handler = [System.EventHandler]{
            Set-ScheduledToggleBackground -Button $buttonForState
        }.GetNewClosure()
        $descriptor.AddValueChanged($Button, $handler)
        [void]$listeners.Add([pscustomobject]@{
            Descriptor = $descriptor
            Handler    = $handler
        })
    }
    $Button.Resources[$listenersKey] = $listeners
    Set-ScheduledToggleBackground -Button $Button

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
