' ###################################################
' # @file yuanshen.vbs
' # @brief 替换启动图片并打开希沃白板
' # @author Rikka Github/ming-14
' ###################################################

Option Explicit
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

Dim swenlauncher
swenlauncher = FindInstallPathIcon("希沃白板 5")

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
	Dim AssetsSplashScreenPath, ResourcesSplashScreenPath, UserprofileBanner
	AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)
	ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)
	UserprofileBanner = "%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png"

	Dim ws
	Set ws = createobject("wscript.shell")
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
		If Not FileExists("%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png" & "Backup") Then
			ws.run "xcopy %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner_backup.png /Y", 0 ' 备份
		End if
		ws.run "xcopy img.pngx %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png /Y", 0
	End if
End Function

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

' @brief 运行前检查
' @note 如果 64bit 检查失败，程序也会启动希沃白板
Function Check()
	' 希沃白板 检查
	If swenlauncher = "" Then
		Msgbox "Cannot find 希沃白板 5"
		WScript.Quit(1)
	End if
	
	' 64bit 检查
	If Not isWindows64bit() Then
		Msgbox "64bit Only"
		ws.run swenlauncher, 0
		WScript.Quit(1)
	End IF
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

' @brief 判断是否为 64bit
Function isWindows64bit()
	Dim is64bit, wmiService, processors, processor

    ' 方法1：通过 WScript.Shell 获取环境变量（兼容所有版本）
    Set shell = CreateObject("WScript.Shell")
    Dim procArch, procArch6432
    procArch = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    procArch6432 = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    On Error Goto 0
    If InStr(1, procArch, "64", vbTextCompare) > 0 Then
        is64bit = True
    ElseIf procArch6432 <> "" And InStr(1, procArch6432, "64", vbTextCompare) > 0 Then
        is64bit = True
    End If

	' 方法2：通过WMI检测（更可靠）
	If Not is64bit Then
		Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
		If Err.Number = 0 Then  ' 确保WMI服务可用
			Set processors = wmiService.ExecQuery("SELECT AddressWidth FROM Win32_Processor")
			If Err.Number = 0 And processors.Count > 0 Then
				For Each processor In processors
					If processor.AddressWidth = 64 Then
						is64bit = True
						Exit For
					End If
				Next
			End If
		End If
	End If

	' 方法3：备用注册表检测
	If Not is64bit Then
		Dim shell, registryValue
		Set shell = CreateObject("WScript.Shell")
		registryValue = ""
		
		On Error Resume Next  ' 防止注册表访问错误
		registryValue = shell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
		If Err.Number = 0 Then
			If InStr(1, registryValue, "64", vbTextCompare) > 0 Then
				is64bit = True
			End If
		End If
		Err.Clear
	End If

	isWindows64bit = is64bit
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

