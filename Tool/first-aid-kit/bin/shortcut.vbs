Option Explicit
On Error Resume Next
Const HKEY_LOCAL_MACHINE = &H80000002

' ɾ�� "ϣ�ְװ� 5 .lnk"
Dim WshShell, desktopPath, fileName, filePath
Set WshShell = CreateObject("WScript.Shell")
desktopPath = WshShell.SpecialFolders("Desktop")
fileName = "ϣ�ְװ� 5 .lnk"
filePath = desktopPath & "\" & fileName
CreateObject("Scripting.FileSystemObject").DeleteFile filePath

' ���� "ϣ�ְװ� 5.lnk"
Dim swenlauncher
swenlauncher = FindInstallPathIcon("ϣ�ְװ� 5")
Msgbox "To " & swenlauncher
Call CreateShortcutOnDesktop("ϣ�ְװ� 5", swenlauncher, 1, swenlauncher, "", swenlauncher) '������ݷ�ʽ
Msgbox "Success"

' @brief �����洴��һ����ݷ�ʽ
' @param name ��ݷ�ʽ����
' @param targetPath ��ݷ�ʽ��ִ��·��
' @param runStyle ���з�ʽ��������1��Ĭ�ϴ��ڼ��������3����󻯼��������7����С�����
' @param Icon ��ݷ�ʽͼ��
' @param description ��ݷ�ʽ������
' @param workingDirectory ��ʼλ��
Function CreateShortcutOnDesktop(name, targetPath, runStyle, Icon, description, workingDirectory)
	Dim WshShell, strDesktop, oShellLink
    set WshShell = Wscript.CreateObject("Wscript.Shell") 
    strDesktop = WshShell.SpecialFolders("Desktop") '�����洴����ݷ�ʽ
    set oShellLink = WshShell.CreateShortcut(strDesktop & "\" & name & ".lnk") '����һ����ݷ�ʽ����
    oShellLink.TargetPath  = targetPath '���ÿ�ݷ�ʽ��ִ��·�� 
    oShellLink.WindowStyle = runStyle '���з�ʽ
    oShellLink.IconLocation= Icon '���ÿ�ݷ�ʽ��ͼ��
    oShellLink.Description = description  '���ÿ�ݷ�ʽ������ 
    oShellLink.WorkingDirectory = workingDirectory '��ʼλ��
    oShellLink.Save
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