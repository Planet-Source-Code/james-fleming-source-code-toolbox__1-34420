Attribute VB_Name = "modError"
' the purpose of this module is to contain all of the error handling message boxes.
' all other message boxes are under modMsgBox

Public Sub ErrMsgBox(myString As String)
'*****************************************************
' Purpose:  This is the sub that launches the message box template
'           for error messages
' Inputs:   The message as a string
' Returns:  None
' Comments: This is used for uniformity of the message box
'           All errors are written to a log file
'*****************************************************
    Static strLast As String
    
    If strLast <> myString Then
        strLast = myString
        Beep
        MsgBox myString, vbCritical, App.Title & " " & App.Major & "." & App.Minor
        Call LogErr(myString)  '// James 1/27/2000 this works fine, just commented out for now.
    End If
End Sub
Public Sub InputErrBox(strErr As String)
'*****************************************************
' Purpose:  This is the sub that launches the message box template
'           for error messages
' Inputs:   The message as a string
' Returns:  None
' Comments: This is used for communicating input errors to the user. It doesn't write to a log.
'*****************************************************
    Beep
    MsgBox "User Input Error: " & strErr, vbCritical, App.Title & " " & App.Major & "." & App.Minor
End Sub

Public Sub LogErr(errMsg As String)
'**************************************************************
'Purpose:   Logs an error in ERRORS.LOG in your applications
'           directory, and continues running your program.
'Arguments: An optional strings, which will be logged for
'           additional information, if desired.
'Dependancies: None.
'Assumes: the file exists App.Path & "\log\errors.log"
'Side Effects: If the error can not be logged, the application
'               will be terminated.
'**************************************************************

    Dim strMsg As String        'The msg in a msgbox explaining the error to the user.
    Dim strTitle As String      'The title of that msgbox
    Dim OldErrDesc As String
    Dim OldErrNum As Long       'The old info is in case the is an error while logging the

    ' error, so that the old error info is not erased.

    Dim intFile As Integer
    'This is the file number, a handle for VB.
    OldErrDesc = Err.Description
    OldErrNum = Err.Number
    
    If lFatal = True Then
        strMsg = "Fatal"
    Else
        strMsg = "Unexpected"
    End If
    
    On Error GoTo ErrWhileLogging:
    'That's in case logging the error generates an error.

 
    'Log the error in error.log

    intFile = FreeFile
    Open App.path & "\log\errors.log" For Append As #intFile
    Print #intFile, ""      ' blank line
    Print #intFile, ""      ' blank line
    Print #intFile, "Date: " & Date$ & "        Time: " & Time$ & "     Error #" & OldErrNum
    Print #intFile, "----------------------------------------------------"
    If Not IsNull(errMsg) Then
        Print #intFile, "Desc: " & errMsg & "!"
    End If
    ' document the error type and path.
    If lFatal Then
        Print #intFile, "Fatal";
    Else
        Print #intFile, "Non-fatal";
    End If
    ' document the app path.
    Print #intFile, " Error in " & App.path & "\";
    Print #intFile, App.Title & " v" & App.Major & "." & App.Minor

    Close #intFile
    Exit Sub
ErrWhileLogging: ' an error in the error log routine!
    strMsg = "Fatal Error: Could not log error." & vbCrLf & _
    "Please contact the program vendor With the following " & _
    "error information:" & vbCrLf & vbCrLf & _
    "Err #" & OldErrNum & vbCrLf & OldErrDesc

    If Not IsNull(errMsg) Then
        strMsg = strMsg & vbCrLf & strInput1
    End If

    MsgBox strMsg
    End
End Sub
Public Function ErrorCode() As String
' here we can print out all the non fatal errors
Dim ErrorNum As Long
Dim ErrorDescription As String
Dim FatalError As String, strReturn As String

On Error GoTo Error_Exit
FatalError = Error(1)
strReturn = ""
    ErrorDescription = vbTab & "Non Fatal Errors:"
      ErrorDescription = Format$(0) & vbTab & "(Nothing) Check to see if you included your exit before your error handler." & vbCrLf
      strReturn = ErrorDescription
For ErrorNum = 1 To 1000

   If Error(ErrorNum) <> FatalError Then
      ErrorDescription = Format$(ErrorNum) & vbTab & Error$(ErrorNum) & vbCrLf
      strReturn = strReturn & ErrorDescription
   End If

   Next ErrorNum
    ErrorCode = strReturn
   Exit Function
Error_Exit:
    Resume Next
End Function
