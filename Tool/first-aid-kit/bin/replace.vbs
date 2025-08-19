Option Explicit
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

Dim swenlauncher
swenlauncher = FindInstallPathIcon("希沃白板 5")

Dim AssetsSplashScreenPath, ResourcesSplashScreenPath, UserprofileBanner
AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)
ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)
UserprofileBanner = "%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png"

Dim ws
Set ws = createobject("wscript.shell")
If FileExists(AssetsSplashScreenPath) Then '替换 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png
	ws.run "xcopy backup.png.bak " & AssetsSplashScreenPath & " /Y", 0
End if
If FileExists(ResourcesSplashScreenPath) Then '替换 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png
	ws.run "xcopy backup.png.bak " & ResourcesSplashScreenPath & " /Y", 0
End if
If FileExists(UserprofileBanner) Then '替换 %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png 
	ws.run "xcopy backup.png.bak %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png /Y", 0
End if
Msgbox "Replace Success"

' @brief 将 swenlauncher\swenlauncher.exe 路径转换成 \Main\Assets\SplashScreen.png
' @note "{IstallPath}\EasiNote5\swenlauncher\swenlauncher.exe" To "{IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png"
' 			1. 遍布 {IstallPath}\EasiNote5\* 找到以 "EasiNote5" 开头的文件夹
' 			2. 拼接 {IstallPath}\EasiNote5\{找到的文件夹}\Main\Assets\SplashScreen.png
Function MainExePathToAssetsSplashScreenPath(mainExePath)
    Dim fso, basePath, parentFolder, subFolder, targetFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 获取主程序所在目录的上级目录（即 {IstallPath}\EasiNote5\）
    basePath = fso.GetParentFolderName(fso.GetParentFolderName(mainExePath))
    
    ' 遍历基础目录下的所有子文件夹
    Set parentFolder = fso.GetFolder(basePath)
    For Each subFolder In parentFolder.SubFolders
        ' 检查文件夹名是否以 "EasiNote5" 开头
        If Left(subFolder.Name, 9) = "EasiNote5" Then
            targetFolder = subFolder.Name
            Exit For
        End If
    Next
    
    ' 拼接目标路径
    If targetFolder <> "" Then
        MainExePathToAssetsSplashScreenPath = fso.BuildPath(basePath, fso.BuildPath(targetFolder, "Main\Assets\SplashScreen.png"))
    Else
        MainExePathToAssetsSplashScreenPath = "" ' 未找到匹配文件夹时返回空
    End If
End Function

' @brief 将 swenlauncher\swenlauncher.exe 路径转换成 \Main\Resources\Startup\SplashScreen.png
' @note "{IstallPath}\EasiNote5\swenlauncher\swenlauncher.exe" To "{IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png"
' 			1. 遍布 {IstallPath}\EasiNote5\* 找到以 "EasiNote5" 开头的文件夹
' 			2. 拼接 {IstallPath}\EasiNote5\{找到的文件夹}\Main\Resources\Startup\SplashScreen.png
Function MainExePathToResourcesSplashScreenPath(mainExePath)
    Dim fso, basePath, parentFolder, subFolder, targetFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 获取主程序所在目录的上级目录（即 {IstallPath}\EasiNote5\）
    basePath = fso.GetParentFolderName(fso.GetParentFolderName(mainExePath))
    
    ' 遍历基础目录下的所有子文件夹
    Set parentFolder = fso.GetFolder(basePath)
    For Each subFolder In parentFolder.SubFolders
        ' 检查文件夹名是否以 "EasiNote5" 开头
        If Left(subFolder.Name, 9) = "EasiNote5" Then
            targetFolder = subFolder.Name
            Exit For
        End If
    Next
    
    ' 拼接目标路径
    If targetFolder <> "" Then
        MainExePathToResourcesSplashScreenPath = fso.BuildPath(basePath, fso.BuildPath(targetFolder, "Main\Resources\Startup\SplashScreen.png"))
    Else
        MainExePathToResourcesSplashScreenPath = "" ' 未找到匹配文件夹时返回空
    End If
End Function

' @brief 判断文件是否存在
' @note 支持环境变量
Function FileExists(path)
    Dim WshShell, expandedPath, fso
    Set WshShell = CreateObject("WScript.Shell")
    expandedPath = WshShell.ExpandEnvironmentStrings(path)
    Set WshShell = Nothing
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(expandedPath)
    Set fso = Nothing
End Function

' @brief 根据程序名查询卸载程序的显示图标
Function FindInstallPathIcon(Name)
    Dim oReg, basePaths, basePath, subKeys, subKey
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    ' 定义要搜索的注册表路径
    basePaths = Array( _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" _
    )
    
    ' 遍历所有注册表路径
    For Each basePath In basePaths
        If oReg.EnumKey(HKEY_LOCAL_MACHINE, basePath, subKeys) = 0 Then
            For Each subKey In subKeys
                Dim fullPath, displayName, installPath
                fullPath = basePath & "\" & subKey
                
                ' 检查DisplayName是否包含目标标识
                If ReadRegistryValue(oReg, HKEY_LOCAL_MACHINE, fullPath, "DisplayName", displayName) Then
                    If InStr(displayName, Name) > 0 Then
                        ' 获取安装路径
                        If ReadRegistryValue(oReg, HKEY_LOCAL_MACHINE, fullPath, "DisplayIcon", installPath) Then
                            If installPath <> "" Then FindInstallPathIcon = installPath : Exit Function
                        End If
                    End If
                End If
            Next
        End If
    Next
    FindInstallPathIcon = ""  ' 未找到时返回空字符串
End Function
Function ReadRegistryValue(oReg, hive, path, valueName, ByRef value)
    Dim valueData
    If oReg.GetStringValue(hive, path, valueName, valueData) = 0 Then
        value = valueData : ReadRegistryValue = True : Exit Function
    End If
    If oReg.GetExpandedStringValue(hive, path, valueName, valueData) = 0 Then
        value = valueData : ReadRegistryValue = True : Exit Function
    End If
    value = "" : ReadRegistryValue = False
End Function

