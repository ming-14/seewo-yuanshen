' ###################################################
' # @file yuanshen.vbs
' # @brief �滻����ͼƬ����ϣ�ְװ�
' # @author Rikka Github/ming-14
' ###################################################
Option Explicit
On Error Resume Next
include("base.vbs")
include("checkEnvironment.vbs")

Dim swenlauncher : swenlauncher = FindInstallPathIcon("ϣ�ְװ� 5")

' +-+-+-+-+-+-+-+-+-+-+-+-+-
' ������ Start
' +-+-+-+-+-+-+-+-+-+-+-+-+-
Check()
ReplaceImg(swenlauncher)
createobject("wscript.shell").run swenlauncher, 0
' +-+-+-+-+-+-+-+-+-+-+-+-+-
' ������ End
' +-+-+-+-+-+-+-+-+-+-+-+-+-

' @brief �滻ͼƬ
Function ReplaceImg(swenlauncher)
	' UserprofileBanner																									' UserprofileBanner ����ͼƬ·��
	Dim AssetsSplashScreenPath : AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)				' Assets 			����ͼƬ·��
	Dim ResourcesSplashScreenPath : ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)	' Resources 		����ͼƬ·��

	Dim ws : Set ws = createobject("wscript.shell")
	If FileExists(AssetsSplashScreenPath) Then '�滻 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png
		If Not FileExists(AssetsSplashScreenPath & "Backup") Then
			ws.run "xcopy " & AssetsSplashScreenPath & " " & AssetsSplashScreenPath & "Backup" & " /Y", 0 ' ����
		End if
		ws.run "xcopy img.pngx " & AssetsSplashScreenPath & " /Y", 0
	End if
	If FileExists(ResourcesSplashScreenPath) Then '�滻 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png
		If Not FileExists(ResourcesSplashScreenPath & "Backup") Then
			ws.run "xcopy " & ResourcesSplashScreenPath & " " & ResourcesSplashScreenPath & "Backup" & " /Y", 0 ' ����
		End if
		ws.run "xcopy img.pngx " & ResourcesSplashScreenPath & " /Y", 0
	End if
	If FileExists(UserprofileBanner) Then '�滻 %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png 
		If Not FileExists(UserprofileBanner & "Backup") Then
			ws.run "xcopy " & UserprofileBanner  & " " & UserprofileBanner & "Backup" & " /Y", 0 ' ����
		End if
		ws.run "xcopy img.pngx " & UserprofileBanner & " /Y", 0
	End if
End Function

' @brief ����ǰ���
' @note ��� 64bit ���ʧ�ܣ�����Ҳ������ϣ�ְװ�
Function Check()
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
		ws.run swenlauncher, 0
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

