VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planet Source Downloader"
   ClientHeight    =   2820
   ClientLeft      =   585
   ClientTop       =   855
   ClientWidth     =   3375
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3375
   Begin ComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2445
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "ConnectedStatus"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "MessageStatus"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "INETStatus"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton btnBrowser 
      Caption         =   ">>"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      ToolTipText     =   "Open browser view"
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop!"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtDestinationDir 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "C:\Shital\PlanetSource"
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox txtEnd 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "55"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtListName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "aishwarya"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton btnGet 
      Caption         =   "Get Messages Now"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin InetCtlsObjects.Inet InetCtrl 
      Left            =   3120
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
   Begin VB.Label lblHyperlink 
      AutoSize        =   -1  'True
      Caption         =   "http://i.am/shital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1920
      MouseIcon       =   "main.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   1200
      Width           =   1185
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Start Msg#"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "End Msg#"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Save In:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Listbot ID:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'Purpose:   Main Form
'Auther:    Shital Shah (http://i.am/shital)
'Date:      12-Mar-1999
'Version:   0.2
'Copyright: Freeware
'


Private WithEvents moFetcher As clsListFetcher
Attribute moFetcher.VB_VarHelpID = -1

Private Const msSTATUS_PANEL_KEY_CONNECTED As String = "ConnectedStatus"
Private Const msSTATUS_PANEL_KEY_PROGRESS As String = "MessageStatus"
Private Const msSTATUS_PANEL_KEY_INET As String = "INETStatus"

'API declaration for ShellExecute function which can execute any file or URL.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Constant to pass in ShellExecute to indicate that open the browser window in normal position
Private Const SW_SHOWNORMAL = 1


Private Sub btnBrowser_Click()
    
    Call UpdateBrowserButtonStatus(frmBrowserView.Visible)
    
End Sub

Private Sub btnGet_Click()

    On Error GoTo ERR_btnGet_Click
    
    'Set the parameters
    With moFetcher
        .StartMessageNumber = txtStart
        .EndMessageNumber = txtEnd
        .ListBotAccountID = txtListName
    End With
    
    'Fetch the messages. This generates the event after
    'every message is fetched.
    Call moFetcher.GetAllMessages

Exit Sub
ERR_btnGet_Click:

    'Cancel the transfer
    Call moFetcher.StopMessageFetch
    
    'Show the error
    MsgBox "Error " & Err.Number & " : " & Err.Description
    
End Sub

Private Sub btnStop_Click()
    
    Call moFetcher.StopMessageFetch
    
End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    'Createt he fetcher class
    Set moFetcher = New clsListFetcher
    
    'Set the INet control that will be used by fetcher class
    Set moFetcher.INetControl = InetCtrl
    
    'Set the detault status by calling the event handler directly
    Call InetCtrl_StateChanged(icNone)
    
    'Update browser button status
    Call UpdateBrowserButtonStatus(True)

End Sub

Private Function SaveMessageToFile(ByVal vsFileName As String, ByVal vsMessageContent As String)

    'Error handling would be done by caller

    Dim nFreeFile As Integer
    
    'Get the handle for free file
    nFreeFile = FreeFile

    'Open the file and save the content
    Open vsFileName For Binary Access Write As #FreeFile
    Put #nFreeFile, , vsMessageContent
    Close #nFreeFile

End Function

Private Sub Form_Unload(Cancel As Integer)
    
    'Free the refrence
    Set moFetcher = Nothing
    
    End
    
End Sub

Private Sub InetCtrl_StateChanged(ByVal State As Integer)
    
    Dim sConnectedStatusText As String
    Dim sINetStatusText As String
    
    'Update status according to INet's status
    Select Case State
    
        Case icNone
            
            sConnectedStatusText = "Ready"
            sINetStatusText = ""
            
            Call UpdateStatus(sConnectedStatusText, sINetStatusText)
            
        Case icConnecting
        
            sConnectedStatusText = "Connecting..."
            
            Call UpdateStatus(sConnectedStatusText)
    
        Case icConnected
        
            sConnectedStatusText = "Connected"
            
            Call UpdateStatus(sConnectedStatusText)
    
        Case icError
        
            sINetStatusText = "Error: " & InetCtrl.ResponseInfo
            
            Call UpdateStatus(, sINetStatusText)
            
        Case icDisconnected
        
            sConnectedStatusText = "Disconnected"
            
            Call UpdateStatus(sConnectedStatusText)
            
        Case icDisconnected
        
            sConnectedStatusText = "Disconnected"
            
            Call UpdateStatus(sConnectedStatusText)
            
        Case icDisconnecting
        
            sConnectedStatusText = "Disconnecting"
            
            Call UpdateStatus(sConnectedStatusText)
            
'        Case icResponseReceived
'
'            sINetStatusText = "Receiving..."
'
'            Call UpdateStatus(, sINetStatusText)
        
        Case icResponseCompleted
        
            sINetStatusText = "Transfer Completed"
            
            Call UpdateStatus(, sINetStatusText)
            
        'Case icDialogPending
        
            'sINetStatusText = "Dialog pending"
            
            'Call UpdateStatus(, sINetStatusText)

    End Select
    
    Call moFetcher.mctlINet_StateChanged(State)
    
End Sub

Private Sub UpdateStatus(Optional vsConnectedStatus As Variant, Optional vsINetStatus As Variant, Optional vlMessageNumberUpdate As Long)

    If Not IsMissing(vsConnectedStatus) Then
    
        'Update the status bar
        sbMain.Panels(msSTATUS_PANEL_KEY_CONNECTED) = vsConnectedStatus
        
    End If

    If Not IsMissing(vsINetStatus) Then
    
        'Update the status bar
        sbMain.Panels(msSTATUS_PANEL_KEY_INET) = vsINetStatus
        
    End If

    If Not IsMissing(vlMessageNumberUpdate) Then
    
        If vlMessageNumberUpdate <> 0 Then
    
            'Update the status bar
            sbMain.Panels(msSTATUS_PANEL_KEY_PROGRESS) = "Msg# " & vlMessageNumberUpdate & " fetched"
            
        Else
        
            sbMain.Panels(msSTATUS_PANEL_KEY_PROGRESS) = "<no message fetched>"
        
        End If
        
    End If

End Sub

Private Sub lblHyperlink_Click()
    
    'Open the URL
    Call ShellExecute(0, "open", lblHyperlink.Caption, &O0, "", SW_SHOWNORMAL)

End Sub

Private Sub moFetcher_Error(ByVal ErrorNumber As Long, ByVal ErrorMessage As String)
    
    'Display the error message
    Call MsgBox("Error " & ErrorNumber & " : " & ErrorMessage)
    
End Sub

Private Sub moFetcher_MessageFetched(ByVal MessageID As Long, ByVal MessageContent As String)

    Call UpdateStatus(, , MessageID)
    
    Dim sFileName As String
    
    'Prepare the filename
    sFileName = GetPathWithSlash(txtDestinationDir) & "ps" & Format(MessageID, "00000") & ".htm"

    'Store the message in the file
    Call SaveMessageToFile(sFileName, MessageContent)
    
    'Show in browser view
    If frmBrowserView.Visible Then
    
        Call frmBrowserView.ShowURL(sFileName)
    
    End If
    
End Sub

Private Sub moFetcher_ProgressNotification(ByVal BytesFetched As Long)
    
    'Update the status
    Call UpdateStatus(, BytesFetched & "bytes fetched")
    
End Sub


Public Sub UpdateBrowserButtonStatus(ByVal vblnMakeOpen As Boolean)

    'Show the browser form
    If vblnMakeOpen Then
    
        frmBrowserView.Hide
        
        btnBrowser.Caption = ">>"
        
        btnBrowser.ToolTipText = "Open Browser View"
        
    Else
    
        frmBrowserView.Show
        
        btnBrowser.Caption = "<<"
        
        btnBrowser.ToolTipText = "Close Browser View"
    
    End If

End Sub
