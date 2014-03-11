VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListFetcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'Purpose:   This class encapsultes the functionality to fetch messages from ListBot (www.listbot.com) mailing list archive.
'Auther:    Shital Shah (http://i.am/shital)
'Date:      12-Mar-1999
'Version:   0.2
'Copyright: Freeware
'


Private mlStartMessageNumber As Long   'Numeric ID of the first message to be fetched
Private mlEndMessageNumber As Long     'Numeric ID of the end message to be fetched
Private msListBotAccountID As String 'Name of the mailing list in ListBot
Private mctlINet As Inet             'INet control that will be used by this class to perform the Internet related functions
Private msFetchedContent As String      'Holds contents of the message temporarily because it is comming part by part
Private mlBytesFetched As Long          'Holds length of the message fetched so far
Private mblnUserStopped As Boolean  'Holds flag whether user stopped the message fetching
Private mblnErrorOccured As Boolean     'Flag to indicate that some error occured


'This event is raised when the message fech is completed
Public Event MessageFetched(ByVal MessageID As Long, ByVal MessageContent As String)

'This event is fired when any error occures
Public Event Error(ByVal ErrorNumber As Long, ByVal ErrorMessage As String)

'This event is raised whenever new chunk is arrived
Public Event ProgressNotification(ByVal BytesFetched As Long)


'Get all the messages from mlStartMessageNumber to mlEndMessageNumber
'After each message is fetched, event is raised
Public Sub GetAllMessages()

    'Clear the flags
    mblnUserStopped = False
    mblnErrorOccured = False
    
    Dim lNextMessageNumber As Long   'Holds the message number that will be fetched next
    
    'Set the first message as the next message to be fetched
    lNextMessageNumber = mlStartMessageNumber
    
    'While next message to be fetched doesn't goes beyond last message, keep om fetching.
    'While loop is used instead of For so that if in midway last message to be fetched is
    'changed, the change will be reflected.
    Do While (lNextMessageNumber <= mlEndMessageNumber) And (Not mblnUserStopped) And (Not mblnErrorOccured)
    
        'Fetch the message. This function raises the event which
        'informs user when message content is available
        Call GetAMessage(lNextMessageNumber)
        
        'Update the message number to be fetched next
        lNextMessageNumber = lNextMessageNumber + 1
    
    Loop
    
End Sub

'Fetches the specified message and returns the message content
Public Function GetAMessage(ByVal MessageNumber As Long) As String
        
    Dim sMessageURL As String 'Complete URL of ListBot site where the message will be found
        
    'This was the old URL
    'sURL = "http://aishwarya.listbot.com/cgi-bin/view_archive?Act=view_message&from_subscriber=false&list_id=" & txtListName & "&msg_num=" & i & "&start_num=" & txtEnd
    
    'Form the URL
    sMessageURL = "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId="
    sMessageURL = sMessageURL & MessageNumber
    
    'Clear the buffer where fetched string will be placed
    msFetchedContent = ""
    mlBytesFetched = 0
    
    'Start fetching
    msFetchedContent = mctlINet.OpenURL(sMessageURL, icString)
    
    Dim sNextChunk As String    'Holds the next comming chunk
    Dim lChunkLength As Long    'Holds the legnth of the chunk
    
    
    Do
    
        sNextChunk = mctlINet.GetChunk(32000, icString)
        
        lChunkLength = Len(sNextChunk)
        
        msFetchedContent = msFetchedContent & sNextChunk
        
        DoEvents
        
    Loop While lChunkLength > 0
    
    
    
'    'While file is getting downloaded
'    Do While mctlINet.StillExecuting
'
'        DoEvents
'
'    Loop
'
'    DoEvents
        
    'Raise the event informing that the transfer is complete
    RaiseEvent MessageFetched(MessageNumber, msFetchedContent)
       
    'Return the message content
    GetAMessage = msFetchedContent

End Function

Private Sub Class_Terminate()
    
    'If INet control is set
    If Not (mctlINet Is Nothing) Then
    
        'If some request is still pending
        If mctlINet.StillExecuting Then
        
            'Cancel the request
            Call mctlINet.Cancel
        
        End If
    
    End If
    
    'Free the refrence
    Set mctlINet = Nothing
    
End Sub

Public Sub mctlINet_StateChanged(ByVal State As Integer)

    On Error GoTo ERR_mctlINet_StateChanged

    Select Case State

       'Occures when any error happens
       Case icError

            mblnErrorOccured = True

            Dim lErrorNumber As Long    'Holds the error number
            Dim sErrorMessage As String  'Holds the error message

            'Set the error number and message
            lErrorNumber = mctlINet.ResponseCode
            sErrorMessage = mctlINet.ResponseInfo

            'Raise the event informing that the error is occured
            RaiseEvent Error(lErrorNumber, sErrorMessage)

    End Select

Exit Sub
ERR_mctlINet_StateChanged:

    'If operation still in progress error occures
    If Err.Number = 33 Then

        DoEvents

        Resume

    End If

End Sub

Public Sub StopMessageFetch()

    'Set the flag indicating that the user stopped the process
    mblnUserStopped = True

    'Try to stop transfer in anyway
    On Error Resume Next

    'Stop the request
    Call mctlINet.Cancel

End Sub



'Property wrappers
Public Property Get StartMessageNumber() As Long
    StartMessageNumber = mlStartMessageNumber
End Property

Public Property Get EndMessageNumber() As Long
    EndMessageNumber = mlEndMessageNumber
End Property

Public Property Get ListBotAccountID() As String
    ListBotAccountID = msListBotAccountID
End Property

Public Property Get INetControl() As Inet
    Set INetControl = mctlINet
End Property

Public Property Let StartMessageNumber(ByVal vlStartMessageNumber As Long)
    mlStartMessageNumber = vlStartMessageNumber
End Property

Public Property Let EndMessageNumber(ByVal vlEndMessageNumber As Long)
    mlEndMessageNumber = vlEndMessageNumber
End Property

Public Property Let ListBotAccountID(ByVal vsListBotAccountID As String)
    msListBotAccountID = vsListBotAccountID
End Property

Public Property Set INetControl(ByVal vctlINetControl As Inet)
    Set mctlINet = vctlINetControl
End Property

