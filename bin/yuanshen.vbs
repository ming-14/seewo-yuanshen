' ###################################################
' # @file yuanshen.vbs
' # @brief �滻����ͼƬ����ϣ�ְװ�
' # @author Rikka Github/ming-14
' ###################################################

Option Explicit
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

Dim swenlauncher
swenlauncher = FindInstallPathIcon("ϣ�ְװ� 5")

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
	Dim AssetsSplashScreenPath, ResourcesSplashScreenPath, UserprofileBanner
	AssetsSplashScreenPath = MainExePathToAssetsSplashScreenPath(swenlauncher)
	ResourcesSplashScreenPath = MainExePathToResourcesSplashScreenPath(swenlauncher)
	UserprofileBanner = "%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png"

	Dim ws
	Set ws = createobject("wscript.shell")
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
		If Not FileExists("%userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png" & "Backup") Then
			ws.run "xcopy %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner_backup.png /Y", 0 ' ����
		End if
		ws.run "xcopy img.pngx %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png /Y", 0
	End if
End Function

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

' @brief ����ǰ���
' @note ��� 64bit ���ʧ�ܣ�����Ҳ������ϣ�ְװ�
Function Check()
	' ϣ�ְװ� ���
	If swenlauncher = "" Then
		Msgbox "Cannot find ϣ�ְװ� 5"
		WScript.Quit(1)
	End if
	
	' 64bit ���
	If Not isWindows64bit() Then
		Msgbox "64bit Only"
		ws.run swenlauncher, 0
		WScript.Quit(1)
	End IF
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

' @brief �ж��Ƿ�Ϊ 64bit
Function isWindows64bit()
	Dim is64bit, wmiService, processors, processor

    ' ����1��ͨ�� WScript.Shell ��ȡ�����������������а汾��
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

	' ����2��ͨ��WMI��⣨���ɿ���
	If Not is64bit Then
		Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
		If Err.Number = 0 Then  ' ȷ��WMI�������
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

	' ����3������ע�����
	If Not is64bit Then
		Dim shell, registryValue
		Set shell = CreateObject("WScript.Shell")
		registryValue = ""
		
		On Error Resume Next  ' ��ֹע�����ʴ���
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

