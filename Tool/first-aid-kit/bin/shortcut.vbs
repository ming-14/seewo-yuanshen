Option Explicit
On Error Resume Next
include("base.vbs")

' É¾³ý "Ï£ÎÖ°×°å 5 .lnk"
Msgbox "Try to delete 'Ï£ÎÖ°×°å 5 .lnk'"
Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
Dim desktopPath : desktopPath = WshShell.SpecialFolders("Desktop")
Dim filePath : filePath = desktopPath & "\" & "Ï£ÎÖ°×°å 5 .lnk"
CreateObject("Scripting.FileSystemObject").DeleteFile filePath

' ´´½¨ "Ï£ÎÖ°×°å 5.lnk"
Msgbox "Try to create 'Ï£ÎÖ°×°å 5.lnk'"
Dim swenlauncher
swenlauncher = FindInstallPathIcon("Ï£ÎÖ°×°å 5")
Msgbox "This is the path of swenlauncher: " & swenlauncher
Call CreateShortcutOnDesktop("Ï£ÎÖ°×°å 5", swenlauncher, 1, swenlauncher, "", swenlauncher) '´´½¨¿ì½Ý·½Ê½
Msgbox "Successfully"

' @brief ÒýÓÃÆäËü VBS
Function include(sInstFile)
    Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = oFSO.OpenTextFile(sInstFile)
    Dim s : s = f.ReadAll
    f.Close
    ExecuteGlobal s
End Function