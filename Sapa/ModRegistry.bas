Attribute VB_Name = "ModRegistry"
'***********************************************************
'-- Ini Modul untuk mendapatkan Akses Registry
'***********************************************************
Option Explicit

Public Declare Function GetPrivateProfileString Lib "KERNEL32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "KERNEL32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    Private Declare Function Beep Lib "KERNEL32" _
(ByVal dwFreq As Long, _
ByVal dwDuration As Long) As Long
Private lReg As Long
Private KeyHandle As Long
Private Const ERROR_SUCCESS = 0&
Private Const REG_SZ = 1
Private Const REG_DWORD = 4
Const KEY_READ = ((&H20000 Or &H1 Or &H8 Or &H10) And (Not &H100000))
Private Const KEY_QUERY_VALUE = &H1
Private Const REG_BINARY = 3
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Enum MainKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Function SetStringValue(hkey As MainKey, sPath As String, sValue As String, sData As String) As Long
    lReg = RegCreateKey(hkey, sPath, KeyHandle)
    lReg = RegSetValueEx(KeyHandle, sValue, 0, REG_SZ, ByVal sData, Len(sData))
    lReg = RegCloseKey(KeyHandle)
End Function

Public Function DeleteValue(ByVal hkey As MainKey, ByVal sPath As String, ByVal sValue As String) As Long
    lReg = RegOpenKey(hkey, sPath, KeyHandle)
    lReg = RegDeleteValue(KeyHandle, sValue)
    lReg = RegCloseKey(KeyHandle)
End Function

Public Function JalanStartUp(Index As Integer)
If Index = 1 Then
    SetStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SapaBoot", App.Path & "\" & App.EXEName & ".exe"
ElseIf Index = 0 Then
    DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "SapaBoot"
End If
End Function

