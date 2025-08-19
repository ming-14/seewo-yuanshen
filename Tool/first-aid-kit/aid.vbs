Option Explicit

Msgbox "hhh被骂了吗"
Msgbox "祝你好运，Good Luck!~"

RunVBS(".\bin\replace.vbs")
RunVBS(".\bin\shortcut.vbs")

' @brief 调用 VBS 脚本
' @note 工作目录是被执行脚本的目录
Function RunVBS(path)
    Dim WshShell, fso, baseFolder, targetScript, scriptFolder
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    targetScript = fso.BuildPath(baseFolder, path)
    scriptFolder = fso.GetParentFolderName(targetScript)
    WshShell.Run "cmd /c cd /d """ & scriptFolder & """ && wscript """ & targetScript & """", 0, True
End Function