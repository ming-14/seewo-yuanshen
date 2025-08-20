Option Explicit
On Error Resume Next
include("base.vbs")

Dim swenlauncher : swenlauncher = FindInstallPathIcon("Ï£ÎÖ°×°å 5")
Dim AssetsSplashScreenPath : AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)
Dim ResourcesSplashScreenPath : ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)

Msgbox "Try to recover the SplashScreen"
Dim ws : Set ws = createobject("wscript.shell")
If FileExists(AssetsSplashScreenPath) Then 'Ìæ»» {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png
	ws.run "xcopy backup.png.bak " & AssetsSplashScreenPath & " /Y", 0
End if
If FileExists(ResourcesSplashScreenPath) Then 'Ìæ»» {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png
	ws.run "xcopy backup.png.bak " & ResourcesSplashScreenPath & " /Y", 0
End if
If FileExists(UserprofileBanner) Then 'Ìæ»» %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png 
	ws.run "xcopy backup.png.bak " & UserprofileBanner & " /Y", 0
End if
Msgbox "Successfully"

' @brief ÒýÓÃÆäËü VBS
Function include(sInstFile)
    Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim f : Set f = oFSO.OpenTextFile(sInstFile)
    Dim s : s = f.ReadAll
    f.Close
    ExecuteGlobal s
End Function