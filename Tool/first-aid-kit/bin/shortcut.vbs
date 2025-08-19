Option Explicit
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

' 删除 "希沃白板 5 .lnk"
Dim WshShell, desktopPath, fileName, filePath
Set WshShell = CreateObject("WScript.Shell")
desktopPath = WshShell.SpecialFolders("Desktop")
fileName = "希沃白板 5 .lnk"
filePath = desktopPath & "\" & fileName
CreateObject("Scripting.FileSystemObject").DeleteFile filePath

' 创建 "希沃白板 5.lnk"
Dim swenlauncher
swenlauncher = FindInstallPathIcon("希沃白板 5")
Msgbox "To " & swenlauncher
Call CreateShortcutOnDesktop("希沃白板 5", swenlauncher, 1, swenlauncher, "", swenlauncher) '创建快捷方式
Msgbox "Success"

' @brief 在桌面创建一个快捷方式
' @param name 快捷方式名称
' @param targetPath 快捷方式的执行路径
' @param runStyle 运行方式（参数‘1’默认窗口激活，参数‘3’最大化激活，参数‘7’最小化激活）
' @param Icon 快捷方式图标
' @param description 快捷方式的描述
' @param workingDirectory 起始位置
Function CreateShortcutOnDesktop(name, targetPath, runStyle, Icon, description, workingDirectory)
	Dim WshShell, strDesktop, oShellLink
    set WshShell = Wscript.CreateObject("Wscript.Shell") 
    strDesktop = WshShell.SpecialFolders("Desktop") '在桌面创建快捷方式
    set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & name & ".lnk") '创建一个快捷方式对象
    oShellLink.TargetPath  = targetPath '设置快捷方式的执行路径 
    oShellLink.WindowStyle = runStyle '运行方式
    oShellLink.IconLocation= Icon '设置快捷方式的图标
    oShellLink.Description = description  '设置快捷方式的描述 
    oShellLink.WorkingDirectory = workingDirectory '起始位置
    oShellLink.Save
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