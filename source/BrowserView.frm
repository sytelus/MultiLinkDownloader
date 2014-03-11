VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowserView 
   Caption         =   "Browser View"
   ClientHeight    =   6255
   ClientLeft      =   3465
   ClientTop       =   2460
   ClientWidth     =   8505
   Icon            =   "BrowserView.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   8505
   Begin VB.TextBox txtFileLocation 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "File Location"
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton btnHide 
      Caption         =   "<<"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin SHDocVwCtl.WebBrowser wbrBrowser 
      Height          =   5655
      Left            =   -120
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   9975
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmBrowserView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnHide_Click()
    Unload Me
End Sub

Public Sub ShowURL(ByVal vsURL As String)

    txtFileLocation.Text = vsURL
    
    Call wbrBrowser.Navigate(vsURL)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Me.Hide
    
    Call frmMain.UpdateBrowserButtonStatus(True)
    
End Sub

Private Sub Form_Resize()
    
    wbrBrowser.Left = 0
    wbrBrowser.Width = Me.Width
    wbrBrowser.Height = Me.Height - wbrBrowser.Top

End Sub
