' name����ݷ�ʽ����
' targetPath����ݷ�ʽ��ִ��·��
' Icon����ݷ�ʽͼ��
' description����ݷ�ʽ������
' workingDirectory����ʼλ��
Function CreateShortcutOnDesktop(name, targetPath, Icon, description, workingDirectory)
    set WshShell    = Wscript.CreateObject("Wscript.Shell") 
    strDesktop  = WshShell.SpecialFolders("Desktop") '�����洴����ݷ�ʽ
    set oShellLink  = WshShell.CreateShortcut(strDesktop&"\"&name&".lnk") '����һ����ݷ�ʽ����
    oShellLink.TargetPath  = targetPath '���ÿ�ݷ�ʽ��ִ��·�� 
    oShellLink.WindowStyle = 7 '���з�ʽ
    oShellLink.IconLocation= Icon '���ÿ�ݷ�ʽ��ͼ��
    oShellLink.Description = description  '���ÿ�ݷ�ʽ������ 
    oShellLink.WorkingDirectory = workingDirectory '��ʼλ��
    oShellLink.Save
End Function

thisFolder = createobject("Scripting.FileSystemObject").GetFolder(".").Path '��ȡ��ǰĿ¼

set ws = createobject("wscript.shell")
ws.run "xcopy backup.png.bak %userprofile%\AppData\Roaming\Seewo\EasiNote5\Resources\Banner\Banner.png /Y",vbhide
Call CreateShortcutOnDesktop("ϣ�ְװ� 5 ", "C:\Program Files (x86)\Seewo\EasiNote5\swenlauncher\swenlauncher.exe", "C:\Program Files (x86)\Seewo\EasiNote5\swenlauncher\swenlauncher.exe", "", "C:\Program Files (x86)\Seewo\EasiNote5\swenlauncher\") '������ݷ�ʽ