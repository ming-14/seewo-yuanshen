Option Explicit
On Error Resume Next
include("base.vbs")

' ɾ�� "ϣ�ְװ� 5 .lnk"
Msgbox "Try to delete 'ϣ�ְװ� 5 .lnk'"
Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim desktopPath : desktopPath = WshShell.SpecialFolders("Desktop")
Dim filePath : filePath = desktopPath & "\" & "ϣ�ְװ� 5 .lnk"
CreateObject("Scripting.FileSystemObject").DeleteFile filePath

' ���� "ϣ�ְװ� 5.lnk"
Msgbox "Try to create 'ϣ�ְװ� 5.lnk'"
Dim swenlauncher
swenlauncher = FindInstallPathIcon("ϣ�ְװ� 5")
Msgbox "This is the path of swenlauncher: " & swenlauncher
Call CreateShortcutOnDesktop("ϣ�ְװ� 5", swenlauncher, 1, swenlauncher, "", swenlauncher) '������ݷ�ʽ
Msgbox "Successfully"

' @brief �������� VBS
Function include(sInstFile)
    Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = oFSO.OpenTextFile(sInstFile)
    Dim s : s = f.ReadAll
    f.Close
    ExecuteGlobal s
End Function