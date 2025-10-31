#Requires -Version 7

# 定时执行 Task.ps1 的脚本
# 每天 18:00 自动运行一次

# 存储已运行日期的变量
$ExecutedDates = @()

# 设置执行时间
$ScheduledHour = 18
$ScheduledMinute = 00

Write-Host "定时任务已启动，将在每天 ${ScheduledHour}:${ScheduledMinute} 执行 Task.ps1"
Write-Host "按 Ctrl+C 停止运行"
Write-Host "当前时间: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host "----------------------------------------"

# 如果脚本在模块中执行，$PSScriptRoot 可直接提供所在目录
if ($PSScriptRoot) {
    Set-Location -Path $PSScriptRoot
} else {
    Set-Location -Path (Split-Path -Parent $PSCommandPath)
}

try {
    while ($true) {
        $CurrentTime = Get-Date
        $CurrentDate = $CurrentTime.ToString("yyyy-MM-dd")

        # 检查是否到达执行时间
        if ($CurrentTime.Hour -eq $ScheduledHour -and $CurrentTime.Minute -eq $ScheduledMinute) {

            # 检查今天是否已经执行过
            if ($ExecutedDates -notcontains $CurrentDate) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - 开始执行 Task.ps1"

                # 检查 Task.ps1 文件是否存在
                $ExportScript = Join-Path $PSScriptRoot "Task.ps1"
                if (Test-Path $ExportScript) {
                    try {
                        # 执行 Task.ps1
                        & $ExportScript
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Task.ps1 执行完成"
                    }
                    catch {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - 执行 Task.ps1 时出错: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - 错误: 找不到 Task.ps1 文件 ($ExportScript)" -ForegroundColor Red
                }

                # 记录已执行的日期
                $ExecutedDates += $CurrentDate
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - 今日任务已完成，记录执行日期: $CurrentDate"
            }
        }

        # 等待1秒
        Start-Sleep -Seconds 1
    }
}
catch {
    Write-Host "脚本运行出错: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Write-Host "定时任务已停止"
}
