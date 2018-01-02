VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1485
   ClientLeft      =   5790
   ClientTop       =   4095
   ClientWidth     =   2865
   LinkTopic       =   "Form2"
   ScaleHeight     =   1485
   ScaleWidth      =   2865
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

     Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, _
        ByVal dwMilliseconds As Long) As Long

     Private Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" _
        (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

     Private Declare Function PostMessage Lib "user32" _
        Alias "PostMessageA" _
        (ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

     Private Declare Function IsWindow Lib "user32" _
        (ByVal hwnd As Long) As Long

     'Constants used by the API functions
     Const WM_CLOSE = &H10
     Const INFINITE = &HFFFFFFFF

Private Sub Form_Load()
App.TaskVisible = False
End Sub


