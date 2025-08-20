' ���� base.vbs
Option Explicit
On Error Resume Next

' @brief �ж��Ƿ�Ϊ 64bit
Function isWindows64bit()
	Dim is64bit, processors, processor

    ' ����1��ͨ�� WScript.Shell ��ȡ�����������������а汾��
    Dim shell : Set shell = CreateObject("WScript.Shell")
    Dim procArch : procArch = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    Dim procArch6432 : procArch6432 = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    On Error Goto 0
    If InStr(1, procArch, "64", vbTextCompare) > 0 Then
        is64bit = True
    ElseIf procArch6432 <> "" And InStr(1, procArch6432, "64", vbTextCompare) > 0 Then
        is64bit = True
    End If

	' ����2��ͨ��WMI��⣨���ɿ���
	If Not is64bit Then
		Dim wmiService : Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
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
		Dim registryValue : registryValue = ""
		
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

' @brief �ж��Ƿ�װϣ�ְװ�
Function isEasiNoteInstalled()
	isEasiNoteInstalled = (FindInstallPathIcon("ϣ�ְװ� 5") <> "")
End Function
