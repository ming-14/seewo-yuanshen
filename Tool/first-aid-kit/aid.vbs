Option Explicit
On Error Resume Next

Msgbox "hhh��������"
Msgbox "ף����ˣ�Good Luck! ~"

RunVBS("bin\replace.vbs")
RunVBS("bin\shortcut.vbs")

' @brief ���� VBS �ű�
' @note ����Ŀ¼�Ǳ�ִ�нű���Ŀ¼
Function RunVBS(path)
    Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    Dim baseFolder : baseFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    Dim targetScript : targetScript = fso.BuildPath(baseFolder, path)
    Dim scriptFolder : scriptFolder = fso.GetParentFolderName(targetScript)
    WshShell.Run "cmd /c cd /d """ & scriptFolder & """ && wscript """ & targetScript & """", 0, True
End Function