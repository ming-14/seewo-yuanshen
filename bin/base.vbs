Option Explicit
On Error Resume Next

Dim UserprofileBanner
UserprofileBanner = "%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png"

' @brief 在桌面创建一个快捷方式
' @param name 快捷方式名称
' @param targetPath 快捷方式的执行路径
' @param runStyle 运行方式（参数‘1’默认窗口激活，参数‘3’最大化激活，参数‘7’最小化激活）
' @param Icon 快捷方式图标
' @param description 快捷方式的描述
' @param workingDirectory 起始位置
Function CreateShortcutOnDesktop(name, targetPath, runStyle, Icon, description, workingDirectory)
    Dim WshShell : Set WshShell = Wscript.CreateObject("Wscript.Shell") 
    Dim strDesktop: strDesktop = WshShell.SpecialFolders("Desktop") '在桌面创建快捷方式
    Dim oShellLink : Set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & name & ".lnk") '创建一个快捷方式对象
    oShellLink.TargetPath  = targetPath '设置快捷方式的执行路径 
    oShellLink.WindowStyle = runStyle '运行方式
    oShellLink.IconLocation= Icon '设置快捷方式的图标
    oShellLink.Description = description  '设置快捷方式的描述 
    oShellLink.WorkingDirectory = workingDirectory '起始位置
    oShellLink.Save
End Function

' @brief 根据程序名查询卸载程序的显示图标
Const HKEY_LOCAL_MACHINE = &H80000002
Function FindInstallPathIcon(Name)
	On Error Resume Next
	If Name = "" Then
		FindInstallPathIcon = ""
		Exit Function
	End If

    Dim basePath, subKeys, subKey
    Dim oReg : Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    ' 定义要搜索的注册表路径
    Dim basePaths : basePaths = Array( _
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

' @brief 判断文件是否存在
' @note 支持环境变量
Function FileExists(path)
    Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
    Dim expandedPath : expandedPath = WshShell.ExpandEnvironmentStrings(path)
    Set WshShell = Nothing
    
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(expandedPath)
End Function

' @brief 将 swenlauncher\swenlauncher.exe 路径转换成 \Main\Assets\SplashScreen.png
' @note "{IstallPath}\EasiNote5\swenlauncher\swenlauncher.exe" To "{IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png"
' 			1. 遍布 {IstallPath}\EasiNote5\* 找到以 "EasiNote5" 开头的文件夹
' 			2. 拼接 {IstallPath}\EasiNote5\{找到的文件夹}\Main\Assets\SplashScreen.png
Function MainExePathToAssetsSplashScreenPath(mainExePath)
    On Error Resume Next
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 获取主程序所在目录的上级目录（即 {IstallPath}\EasiNote5\）
    Dim basePath : basePath = fso.GetParentFolderName(fso.GetParentFolderName(mainExePath))
    
    Dim subFolder, targetFolder
    ' 遍历基础目录下的所有子文件夹
    Dim parentFolder : Set parentFolder = fso.GetFolder(basePath)
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
    On Error Resume Next
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 获取主程序所在目录的上级目录（即 {IstallPath}\EasiNote5\）
    Dim basePath : basePath = fso.GetParentFolderName(fso.GetParentFolderName(mainExePath))
    
    Dim subFolder, targetFolder
    ' 遍历基础目录下的所有子文件夹
    Dim parentFolder : Set parentFolder = fso.GetFolder(basePath)
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

Function GetAssetsSplashScreenPath()
	GetAssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(FindInstallPathIcon("希沃白板 5"))
End Function
Function GetResourcesSplashScreenPath()
	GetResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(FindInstallPathIcon("希沃白板 5"))
End Function