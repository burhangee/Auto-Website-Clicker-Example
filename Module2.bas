Attribute VB_Name = "Module2"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, ByRef phkResult As Long) As Long


Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long


Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
    ' Note that if you declare the lpData pa
    '     rameter as String in RegSetValueEx, you
    '     must pass it ByVal.


Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long


Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
    ByVal hKey As Long, ByVal lpValueName As String) As Long


Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
    ByVal hKey As Long, ByVal lpSubKey As String) As Long


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_OPTION_NON_VOLATILE = 0
Public Const SYNCHRONIZE = &H100000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_WRITE = &H20006
Public Const KEY_ALL_ACCESS = &H2003F
Public Const KEY_READ = _
((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
' Registry location
Public Const gREGKEYLocation = "SOFTWARE\Your Company Name\Your App Name\Your Current Version"
Public Const gREGKEYXPos = "XPos"
Public Const gREGKEYYPos = "YPos"
Public Const gREGKEYWidth = "Width"
Public Const gREGKEYHeight = "Height"
Public Const gREGKEYWindowState = "WindowState"
Public Const ERROR_SUCCESS = 0&


Public Sub GetRegistryKeys()
    Dim strXPos$, strYPos$, strHeight$, strWidth$, strWindowState$
    
    GetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, gREGKEYXPos, strXPos
    GetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, gREGKEYYPos, strYPos
    GetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, gREGKEYWidth, strWidth
    GetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, gREGKEYHeight, strHeight
    GetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, gREGKEYWindowState, strWindowState
    'Checking the Len is faster than checkin
    '     g if the value equals ""
    'MIN_WIDTH and MIN_HEIGHT can be defined
    '     as the smallest value the window is allo
    '     wed


    If Len(strWidth) <> 0 Then
        frmMain.Width = IIf(CInt(strWidth) > MIN_WIDTH, CInt(strWidth), MIN_WIDTH)
    Else: frmMain.Width = MIN_WIDTH
    End If
    


    If Len(strHeight) <> 0 Then
        frmMain.Height = IIf(CInt(strHeight) > MIN_HEIGHT, CInt(strHeight), MIN_HEIGHT)
    Else: frmMain.Height = MIN_HEIGHT
    End If
    
    'This sets the location of the window to
    '     what is saved in the registry.
    'IF the value does not exist, then it wi
    '     ll place it in the center of the screen.
    '


    If Len(strXPos) <> 0 Then
        frmMain.Left = IIf(CInt(strXPos) > 0, CInt(strXPos), (Screen.Width - frmMain.Width) / 2)
    End If


    If Len(strYPos) <> 0 Then
        frmMain.Top = IIf(CInt(strYPos) > 0, CInt(strYPos), (Screen.Height - frmMain.Height) / 2)
    Else: frmMain.Top = (Screen.Height - frmMain.Height) / 2
    End If
    
    'Sets the app up to be either normal, or
    '     maximized, based on how the user left it
    '     last


    If Len(strWindowState) > 0 Then


        Select Case CInt(strWindowState)
            Case vbMaximized
            frmMain.WindowState = vbMaximized
            Case Else
            frmMain.WindowState = vbNormal
        End Select
End If
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, _
    ByRef KeyVal As String) As Boolean
    Dim i As Long
    Dim rc As Long
    Dim hKey As Long
    Dim KeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    


    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else
        tmpVal = Left(tmpVal, KeyValSize)
    End If
    


    Select Case KeyValType
        Case REG_DWORD


        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Format(Hex(Asc(Mid(tmpVal, i, 1))), "00")
        Next
        KeyVal = Format$("&h" + KeyVal)
        Case REG_SZ
        KeyVal = tmpVal
    End Select

GetKeyValue = True
rc = RegCloseKey(hKey)
Exit Function

GetKeyError:
GetKeyValue = False
rc = RegCloseKey(hKey)
End Function


Public Sub SetRegistryKeys()
    Dim strF1Prefixes$, strF1PrefixesEnabled$
    'Deletes the entire key so it can re-wri
    '     te it. This is an easy way
    'to manage values that may need to be sa
    '     ved with less data. For
    'example, if an MRU list upon opening th
    '     e app has 4 entries, and
    'when the app is closed only has three,
    '     you don't need to worry about
    'determining if there is one extra in th
    '     e registry and deleting it.
    DeleteRegKey gREGKEYLocation
    
    'If the window is minimized, then set it
    '     to a normal size before saving.
    'This way it will not be opened in a min
    '     imized state.
    If frmMain.WindowState = vbMinimized Then frmMain.WindowState = vbNormal
    
    'Save the windowstate to the registry
    SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYWindowState, frmMain.WindowState
    
    'Put the window at a normal state to set
    '     the correct window sizes in the registry
    '
    frmMain.WindowState = vbNormal
    'Set all the window positions.


    If frmMain.Left >= 0 Then
        SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYXPos, frmMain.Left
    Else: SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYXPos, 0
    End If


    If frmMain.Top >= 0 Then
        SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYYPos, frmMain.Top
    Else: SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYYPos, 0
    End If
    SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYWidth, frmMain.Width
    SetKeyValue HKEY_LOCAL_MACHINE, gREGKEYLocation, REG_DWORD, gREGKEYHeight, frmMain.Height
End Sub


Public Function SetKeyValue(KeyRoot As Long, KeyName As String, lType As Long, SubKeyRef As String, KeyVal As Variant) As Boolean
    Dim rc As Long
    Dim hKey As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    If (rc <> ERROR_SUCCESS) Then
        Call RegCreateKey(KeyRoot, KeyName, hKey)
    End If
  
    Select Case lType
        Case REG_SZ
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, REG_SZ, ByVal CStr(KeyVal & Chr$(0)), Len(KeyVal))
        Case REG_BINARY
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, REG_BINARY, ByVal CStr(KeyVal & Chr$(0)), Len(KeyVal))
        Case REG_DWORD
        rc = RegSetValueEx(hKey, SubKeyRef, 0&, REG_DWORD, CLng(KeyVal), 4)
    End Select
If (rc <> ERROR_SUCCESS) Then GoTo SetKeyError

SetKeyValue = True
rc = RegCloseKey(hKey)

Exit Function
SetKeyError:
KeyVal = ""
SetKeyValue = False
rc = RegCloseKey(hKey)
End Function


Public Function DeleteRegValue(KeyName As String, SubKeyRef As String) As Boolean
    Dim rc As Long
    Dim hKey As Long
    rc = RegOpenKeyEx(HKEY_LOCAL_MACHINE, KeyName, 0, KEY_ALL_ACCESS, hKey)
    If (rc <> ERROR_SUCCESS) Then GoTo DeleteKeyError
    rc = RegDeleteValue(hKey, SubKeyRef)
    If (rc <> ERROR_SUCCESS) Then GoTo DeleteKeyError
    DeleteRegValue = True
    Exit Function
DeleteKeyError:
    DeleteRegValue = False
    
End Function


Public Function DeleteRegKey(KeyName As String) As Boolean
    
    Dim rc As Long
    'All sub keys must be deleted for this t
    '     o work.
    'If you create key under your original k
    '     ey, you
    'need to delete it forst.
    rc = RegDeleteKey(HKEY_LOCAL_MACHINE, KeyName)
    DeleteRegKey = IIf(rc = ERROR_SUCCESS, True, False)
End Function



