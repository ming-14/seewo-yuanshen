' ###################################################
' # @file yuanshen.vbs
' # @brief 替换启动图片并打开希沃白板
' # @author Rikka Github/ming-14
' ###################################################
Option Explicit
On Error Resume Next
include("base.vbs")
include("checkEnvironment.vbs")

Dim swenlauncher : swenlauncher = FindInstallPathIcon("希沃白板 5")

' +-+-+-+-+-+-+-+-+-+-+-+-+-
' 主程序 Start
' +-+-+-+-+-+-+-+-+-+-+-+-+-
Check()
ReplaceImg(swenlauncher)
createobject("wscript.shell").run swenlauncher, 0
' +-+-+-+-+-+-+-+-+-+-+-+-+-
' 主程序 End
' +-+-+-+-+-+-+-+-+-+-+-+-+-

' @brief 替换图片
Function ReplaceImg(swenlauncher)
	' UserprofileBanner																									' UserprofileBanner 启动图片路径
	Dim AssetsSplashScreenPath : AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)				' Assets 			启动图片路径
	Dim ResourcesSplashScreenPath : ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)	' Resources 		启动图片路径

	Dim ws : Set ws = createobject("wscript.shell")
	If FileExists(AssetsSplashScreenPath) Then '替换 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png
		If Not FileExists(AssetsSplashScreenPath & "Backup") Then
			ws.run "xcopy " & AssetsSplashScreenPath & " " & AssetsSplashScreenPath & "Backup" & " /Y", 0 ' 备份
		End if
		ws.run "xcopy img.pngx " & AssetsSplashScreenPath & " /Y", 0
	End if
	If FileExists(ResourcesSplashScreenPath) Then '替换 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png
		If Not FileExists(ResourcesSplashScreenPath & "Backup") Then
			ws.run "xcopy " & ResourcesSplashScreenPath & " " & ResourcesSplashScreenPath & "Backup" & " /Y", 0 ' 备份
		End if
		ws.run "xcopy img.pngx " & ResourcesSplashScreenPath & " /Y", 0
	End if
	If FileExists(UserprofileBanner) Then '替换 %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png 
		If Not FileExists(UserprofileBanner & "Backup") Then
			ws.run "xcopy " & UserprofileBanner  & " " & UserprofileBanner & "Backup" & " /Y", 0 ' 备份
		End if
		ws.run "xcopy img.pngx " & UserprofileBanner & " /Y", 0
	End if
End Function

' @brief 运行前检查
' @note 如果 64bit 检查失败，程序也会启动希沃白板
Function Check()
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
		ws.run swenlauncher, 0
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

