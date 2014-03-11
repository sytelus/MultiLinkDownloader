Attribute VB_Name = "modUtils"
Option Explicit

'
'==========================================================================================
' Routine Name : GetPathWithSlash
' Purpose      : Checks if slash already present at the end of path. If not append it.
' Parameters   : a string containing path
' Return       : Path with slash
' Effects      : None
' Assumes      : None
' Author       : shital
' Date         : 09-Apr-1998 02:43 PM
' Template     : Ver.11   Author: Shital Shah   Date: 07 Apr, 1998
' Revision History :
' Date          Person      Details.
'==========================================================================================
'

Public Function GetPathWithSlash(ByVal vsPath As String) As String

    If Trim$(vsPath) <> "" Then
    
        If Right$(vsPath, 1) <> "\" Then
            
            GetPathWithSlash = vsPath + "\"
            
        Else
        
            GetPathWithSlash = vsPath
        
        End If

    Else
        
        GetPathWithSlash = ""
        
    End If

End Function

