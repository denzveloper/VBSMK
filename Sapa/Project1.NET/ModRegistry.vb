Option Strict Off
Option Explicit On
Module ModRegistry
	'***********************************************************
	'-- Ini Modul untuk mendapatkan Akses Registry
	'***********************************************************
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function GetPrivateProfileString Lib "KERNEL32"  Alias "GetPrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Public Declare Function WritePrivateProfileString Lib "KERNEL32"  Alias "WritePrivateProfileStringA"(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
	'UPGRADE_NOTE: Beep was upgraded to Beep_Renamed. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Declare Function Beep_Renamed Lib "KERNEL32" (ByVal dwFreq As Integer, ByVal dwDuration As Integer) As Integer
	Private lReg As Integer
	Private KeyHandle As Integer
	Private Const ERROR_SUCCESS As Short = 0
	Private Const REG_SZ As Short = 1
	Private Const REG_DWORD As Short = 4
	Const KEY_READ As Boolean = ((&H20000 Or &H1s Or &H8s Or &H10s) And (Not &H100000))
	Private Const KEY_QUERY_VALUE As Short = &H1s
	Private Const REG_BINARY As Short = 3
	Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Integer) As Integer
	Private Declare Function RegCreateKey Lib "advapi32.dll"  Alias "RegCreateKeyA"(ByVal hkey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	Private Declare Function RegDeleteValue Lib "advapi32.dll"  Alias "RegDeleteValueA"(ByVal hkey As Integer, ByVal lpValueName As String) As Integer
	Private Declare Function RegOpenKey Lib "advapi32.dll"  Alias "RegOpenKeyA"(ByVal hkey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RegQueryValueEx Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hkey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Any, ByRef lpcbData As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Function RegSetValueEx Lib "advapi32.dll"  Alias "RegSetValueExA"(ByVal hkey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpData As Any, ByVal cbData As Integer) As Integer
	Public Enum MainKey
		HKEY_CLASSES_ROOT = &H80000000
		HKEY_CURRENT_USER = &H80000001
		HKEY_LOCAL_MACHINE = &H80000002
		HKEY_USERS = &H80000003
	End Enum
	
	Public Function SetStringValue(ByRef hkey As MainKey, ByRef sPath As String, ByRef sValue As String, ByRef sData As String) As Integer
		lReg = RegCreateKey(hkey, sPath, KeyHandle)
		lReg = RegSetValueEx(KeyHandle, sValue, 0, REG_SZ, sData, Len(sData))
		lReg = RegCloseKey(KeyHandle)
	End Function
	
	Public Function DeleteValue(ByVal hkey As MainKey, ByVal sPath As String, ByVal sValue As String) As Integer
		lReg = RegOpenKey(hkey, sPath, KeyHandle)
		lReg = RegDeleteValue(KeyHandle, sValue)
		lReg = RegCloseKey(KeyHandle)
	End Function
	
	Public Function JalanStartUp(ByRef Index As Short) As Object
		If Index = 1 Then
			'UPGRADE_WARNING: App property App.EXEName has a new behavior. Click for more: 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			SetStringValue(MainKey.HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SapaBoot", My.Application.Info.DirectoryPath & "\" & My.Application.Info.AssemblyName & ".exe")
		ElseIf Index = 0 Then 
			DeleteValue(MainKey.HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SapaBoot")
		End If
	End Function
End Module