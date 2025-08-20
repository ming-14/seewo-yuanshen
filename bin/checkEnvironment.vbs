' 依赖 base.vbs
Option Explicit
On Error Resume Next

' @brief 判断是否为 64bit
Function isWindows64bit()
	Dim is64bit, processors, processor

    ' 方法1：通过 WScript.Shell 获取环境变量（兼容所有版本）
    Dim shell : Set shell = CreateObject("WScript.Shell")
    Dim procArch : procArch = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%")
    Dim procArch6432 : procArch6432 = shell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITEW6432%")
    On Error Goto 0
    If InStr(1, procArch, "64", vbTextCompare) > 0 Then
        is64bit = True
    ElseIf procArch6432 <> "" And InStr(1, procArch6432, "64", vbTextCompare) > 0 Then
        is64bit = True
    End If

	' 方法2：通过WMI检测（更可靠）
	If Not is64bit Then
		Dim wmiService : Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
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
		Dim registryValue : registryValue = ""
		
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

' @brief 判断是否安装希沃白板
Function isEasiNoteInstalled()
	isEasiNoteInstalled = (FindInstallPathIcon("希沃白板 5") <> "")
End Function
