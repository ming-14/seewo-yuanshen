Option Explicit

Msgbox "hhh��������"
Msgbox "ף����ˣ�Good Luck!~"

RunVBS(".\bin\replace.vbs")
RunVBS(".\bin\shortcut.vbs")

' @brief ���� VBS �ű�
' @note ����Ŀ¼�Ǳ�ִ�нű���Ŀ¼
Function RunVBS(path)
    Dim WshShell, fso, baseFolder, targetScript, scriptFolder
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    targetScript = fso.BuildPath(baseFolder, path)
    scriptFolder = fso.GetParentFolderName(targetScript)
    WshShell.Run "cmd /c cd /d """ & scriptFolder & """ && wscript """ & targetScript & """", 0, True
End Function