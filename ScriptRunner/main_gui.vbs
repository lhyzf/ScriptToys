Option Explicit

Dim shell
Dim fileSystem
Dim scriptDir
Dim mainUiPath
Dim command

Set shell = CreateObject("WScript.Shell")
Set fileSystem = CreateObject("Scripting.FileSystemObject")

scriptDir = fileSystem.GetParentFolderName(WScript.ScriptFullName)
mainUiPath = fileSystem.BuildPath(scriptDir, "main_ui.ps1")

If Not fileSystem.FileExists(mainUiPath) Then
    MsgBox "未找到脚本: " & mainUiPath, vbCritical, "启动失败"
    WScript.Quit 1
End If

shell.CurrentDirectory = scriptDir
command = "pwsh.exe -NoLogo -NoProfile -File """ & mainUiPath & """"
shell.Run command, 0, False
