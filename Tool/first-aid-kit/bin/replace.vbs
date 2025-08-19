Option Explicit
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

Dim swenlauncher
swenlauncher = FindInstallPathIcon("ϣ�ְװ� 5")

Dim AssetsSplashScreenPath, ResourcesSplashScreenPath, UserprofileBanner
AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)
ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)
UserprofileBanner = "%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png"

Dim ws
Set ws = createobject("wscript.shell")
If FileExists(AssetsSplashScreenPath) Then '�滻 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png
	ws.run "xcopy backup.png.bak " & AssetsSplashScreenPath & " /Y", 0
End if
If FileExists(ResourcesSplashScreenPath) Then '�滻 {IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png
	ws.run "xcopy backup.png.bak " & ResourcesSplashScreenPath & " /Y", 0
End if
If FileExists(UserprofileBanner) Then '�滻 %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png 
	ws.run "xcopy backup.png.bak %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png /Y", 0
End if
Msgbox "Replace Success"

' @brief �� swenlauncher\swenlauncher.exe ·��ת���� \Main\Assets\SplashScreen.png
' @note "{IstallPath}\EasiNote5\swenlauncher\swenlauncher.exe" To "{IstallPath}\EasiNote5\EasiNote5_{version}\Main\Assets\SplashScreen.png"
' 			1. �鲼 {IstallPath}\EasiNote5\* �ҵ��� "EasiNote5" ��ͷ���ļ���
' 			2. ƴ�� {IstallPath}\EasiNote5\{�ҵ����ļ���}\Main\Assets\SplashScreen.png
Function MainExePathToAssetsSplashScreenPath(mainExePath)
    Dim fso, basePath, parentFolder, subFolder, targetFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ��ȡ����������Ŀ¼���ϼ�Ŀ¼���� {IstallPath}\EasiNote5\��
    basePath = fso.GetParentFolderName(fso.GetParentFolderName(mainExePath))
    
    ' ��������Ŀ¼�µ��������ļ���
    Set parentFolder = fso.GetFolder(basePath)
    For Each subFolder In parentFolder.SubFolders
        ' ����ļ������Ƿ��� "EasiNote5" ��ͷ
        If Left(subFolder.Name, 9) = "EasiNote5" Then
            targetFolder = subFolder.Name
            Exit For
        End If
    Next
    
    ' ƴ��Ŀ��·��
    If targetFolder <> "" Then
        MainExePathToAssetsSplashScreenPath = fso.BuildPath(basePath, fso.BuildPath(targetFolder, "Main\Assets\SplashScreen.png"))
    Else
        MainExePathToAssetsSplashScreenPath = "" ' δ�ҵ�ƥ���ļ���ʱ���ؿ�
    End If
End Function

' @brief �� swenlauncher\swenlauncher.exe ·��ת���� \Main\Resources\Startup\SplashScreen.png
' @note "{IstallPath}\EasiNote5\swenlauncher\swenlauncher.exe" To "{IstallPath}\EasiNote5\EasiNote5_{version}\Main\Resources\Startup\SplashScreen.png"
' 			1. �鲼 {IstallPath}\EasiNote5\* �ҵ��� "EasiNote5" ��ͷ���ļ���
' 			2. ƴ�� {IstallPath}\EasiNote5\{�ҵ����ļ���}\Main\Resources\Startup\SplashScreen.png
Function MainExePathToResourcesSplashScreenPath(mainExePath)
    Dim fso, basePath, parentFolder, subFolder, targetFolder
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' ��ȡ����������Ŀ¼���ϼ�Ŀ¼���� {IstallPath}\EasiNote5\��
    basePath = fso.GetParentFolderName(fso.GetParentFolderName(mainExePath))
    
    ' ��������Ŀ¼�µ��������ļ���
    Set parentFolder = fso.GetFolder(basePath)
    For Each subFolder In parentFolder.SubFolders
        ' ����ļ������Ƿ��� "EasiNote5" ��ͷ
        If Left(subFolder.Name, 9) = "EasiNote5" Then
            targetFolder = subFolder.Name
            Exit For
        End If
    Next
    
    ' ƴ��Ŀ��·��
    If targetFolder <> "" Then
        MainExePathToResourcesSplashScreenPath = fso.BuildPath(basePath, fso.BuildPath(targetFolder, "Main\Resources\Startup\SplashScreen.png"))
    Else
        MainExePathToResourcesSplashScreenPath = "" ' δ�ҵ�ƥ���ļ���ʱ���ؿ�
    End If
End Function

' @brief �ж��ļ��Ƿ����
' @note ֧�ֻ�������
Function FileExists(path)
    Dim WshShell, expandedPath, fso
    Set WshShell = CreateObject("WScript.Shell")
    expandedPath = WshShell.ExpandEnvironmentStrings(path)
    Set WshShell = Nothing
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(expandedPath)
    Set fso = Nothing
End Function

' @brief ���ݳ�������ѯж�س������ʾͼ��
Function FindInstallPathIcon(Name)
    Dim oReg, basePaths, basePath, subKeys, subKey
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    ' ����Ҫ������ע���·��
    basePaths = Array( _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", _
        "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" _
    )
    
    ' ��������ע���·��
    For Each basePath In basePaths
        If oReg.EnumKey(HKEY_LOCAL_MACHINE, basePath, subKeys) = 0 Then
            For Each subKey In subKeys
                Dim fullPath, displayName, installPath
                fullPath = basePath & "\" & subKey
                
                ' ���DisplayName�Ƿ����Ŀ���ʶ
                If ReadRegistryValue(oReg, HKEY_LOCAL_MACHINE, fullPath, "DisplayName", displayName) Then
                    If InStr(displayName, Name) > 0 Then
                        ' ��ȡ��װ·��
                        If ReadRegistryValue(oReg, HKEY_LOCAL_MACHINE, fullPath, "DisplayIcon", installPath) Then
                            If installPath <> "" Then FindInstallPathIcon = installPath : Exit Function
                        End If
                    End If
                End If
            Next
        End If
    Next
    FindInstallPathIcon = ""  ' δ�ҵ�ʱ���ؿ��ַ���
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

