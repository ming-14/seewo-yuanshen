' ###################################################
' # @file CreateShortcuts.vbs
' # @brief �ٳֿ�ݷ�ʽ
' # @author Rikka Github/ming-14
' ###################################################
Option Explicit
include("bin\base.vbs")
include("bin\checkEnvironment.vbs")

' +-+-+-+-+-+-+-+-+-+-+-+-+-
' ������ Start
' +-+-+-+-+-+-+-+-+-+-+-+-+-
Check() ' ��黷��

Dim swenlauncher : swenlauncher = FindInstallPathIcon("ϣ�ְװ� 5")
Dim thisFolder : thisFolder = createobject("Scripting.FileSystemObject").GetFolder(".").Path ' ��ȡ��ǰĿ¼
Call CreateShortcutOnDesktop("ϣ�ְװ� 5 ", thisFolder & "\bin\yuanshen.vbs", 7, swenlauncher, "", thisFolder & "\bin\")

Msgbox("Successfully")
' +-+-+-+-+-+-+-+-+-+-+-+-+-
' ������ End
' +-+-+-+-+-+-+-+-+-+-+-+-+-

' @brief ��黷��
Function Check()
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	' �������ļ��Ƿ�����
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
		Msgbox("�����ļ�ȱʧ��" & vbCrLf & missingFiles)
		WScript.Quit(1)
	End If
	
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	' ����Ƿ�װϣ�ְװ�5
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	If Not isEasiNoteInstalled() Then
		Msgbox "δ�ҵ�ϣ�ְװ�5"
		WScript.Quit(1)
	End if
	
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	' ���ϵͳ�Ƿ���64λ
	' +-+-+-+-+-+-+-+-+-+-+-+-+-
	If Not isWindows64bit() Then
		Msgbox "�ó���ֻ������64λ"
		WScript.Quit(1)
	End IF
End Function

' @brief �������� VBS
Function include(sInstFile)
    Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = oFSO.OpenTextFile(sInstFile)
    Dim s : s = f.ReadAll
    f.Close
    ExecuteGlobal s
End Function
