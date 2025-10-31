#Requires -Version 7

# 执行特定任务

# 如果脚本在模块中执行，$PSScriptRoot 可直接提供所在目录
if ($PSScriptRoot) {
    Set-Location -Path $PSScriptRoot
} else {
    Set-Location -Path (Split-Path -Parent $PSCommandPath)
}

Write-Host "任务执行完成！" -ForegroundColor Green
