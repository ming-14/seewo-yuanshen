' ###################################################
' # @file CreateShortcuts.vbs
' # @brief 劫持快捷方式
' # @author Rikka Github/ming-14
' ###################################################
Option Explicit
include("bin\base.vbs")
include("bin\checkEnvironment.vbs")

' +-+-+-+-+-+-+-+-+-+-+-+-+-
' 主程序 Start
' +-+-+-+-+-+-+-+-+-+-+-+-+-
Check() ' 检查环境

Dim swenlauncher : swenlauncher = FindInstallPathIcon("希沃白板 5")
Dim thisFolder : thisFolder = createobject("Scripting.FileSystemObject").GetFolder(".").Path ' 获取当前目录
Call CreateShortcutOnDesktop("希沃白板 5 ", thisFolder & "\bin\yuanshen.vbs", 7, swenlauncher, "", thisFolder & "\bin\")

Msgbox("Successfully")
' +-+-+-+-+-+-+-+-+-+-+-+-+-
' 主程序 End
' +-+-+-+-+-+-+-+-+-+-+-+-+-

' @brief 检查环境
Function Check()
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	' 检查程序文件是否完整
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	Dim fileList : fileList = Array( _
		"bin/base.vbs", _
		"bin/checkEnvironment.vbs", _
		"bin/yuanshen.vbs", _
		"bin/img.pngx" _
	)
	Dim missingFiles : missingFiles = ""
	Dim allExist : allExist = True
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim filePath
	For Each filePath In fileList
		If Not fso.FileExists(filePath) Then
			missingFiles = missingFiles & " - " & filePath & vbCrLf
			allExist = False
		End If
	Next

	If Not allExist Then
		Msgbox("以下文件缺失：" & vbCrLf & missingFiles)
		WScript.Quit(1)
	End If
	
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	' 检查是否安装希沃白板5
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	If Not isEasiNoteInstalled() Then
		Msgbox "未找到希沃白板5"
		WScript.Quit(1)
	End if
	
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	' 检查系统是否是64位
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	If Not isWindows64bit() Then
		Msgbox "该程序只适用于64位"
		WScript.Quit(1)
	End IF
End Function

' @brief 引用其它 VBS
Function include(sInstFile)
    Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = oFSO.OpenTextFile(sInstFile)
    Dim s : s = f.ReadAll
    f.Close
    ExecuteGlobal s
End Function
