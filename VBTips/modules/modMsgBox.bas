Attribute VB_Name = "modMsgBox"
Option Explicit

Public Function YesNo(strMessage As String) As Integer
'*****************************************************
' Purpose:  This is a generic YesNo box
' Inputs:   The message as a string, an alternate title
' Returns:  the int value to be handled.
' Comment:
'*****************************************************
On Error GoTo myErrHandler
    Dim intResponse As Integer
    intResponse = MsgBox(strMessage, vbYesNo + vbQuestion, App.Title & " " & App.Major & "." & App.Minor)
    If intResponse = vbYes Then
        YesNo = vbYes
    Else
        YesNo = vbNo                  ' skip it!
    End If                          ' end the nested if
    Exit Function                   ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Function

Public Sub InputInfoBox(strInfo As String)
'*****************************************************
' Purpose:  This is the sub that launches the message box template
'           for error messages
' Inputs:   The message as a string
' Returns:  None
' Comments: This is used for communicating input errors to the user. It doesn't write to a log.
'*****************************************************
    Beep
    MsgBox strInfo, vbInformation, App.Title & " " & App.Major & "." & App.Minor
End Sub
