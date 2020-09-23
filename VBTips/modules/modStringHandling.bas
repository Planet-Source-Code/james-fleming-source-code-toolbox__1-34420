Attribute VB_Name = "modStringHandling"
Option Explicit
Public Function StringSplit(strPath As String, iCounter As Integer) As String
'*******************************************************
' Purpose:  Takes a string as input and removes the file name from the path
' Inputs:   The string being changed. the number of steps back up the path (# of \ being removed)
' Returns:  a path without its file
'*******************************************************
    Dim arrPath() As String
    Dim i As Integer
    Dim strNew As String
    '-Parse String
    arrPath = Split(strPath, "\")
    
    '-Loop through array
    For i = LBound(arrPath) To (UBound(arrPath) - iCounter)
    ' MsgBox arrPath(iCounter)
    strNew = strNew & arrPath(i)
    strNew = strNew & "\"   ' this particular function requires the path to end on \
    Next
    strNew = Left(strNew, Len(strNew) - 1)
    StringSplit = strNew
End Function
'

Function ExtStrip(InString As String) As String
'*******************************************************
' Purpose:  Takes a string as input and strips away any 1,2,3 or 4 letter extension
' Inputs:   The string being changed.   Returns: a path and file without its . suffix
'*******************************************************
    
    Dim OutString As String, CurrentLetter As String
    Dim iCount As Integer, dotFound As Boolean
    OutString = InString
    
    If InString = "" Then   ' The input string is blank
         Exit Function      ' Then so shall be the output!
    End If
    
    For iCount = 1 To Len(InString)
    CurrentLetter = Right(OutString, 1)
    If CurrentLetter <> "." Then
        OutString = Left(OutString, Len(OutString) - 1)
    Else
        dotFound = True
        OutString = Left(OutString, Len(OutString) - 1) ' lop off the (.)
        Exit For
    End If
    Next
    
    If dotFound = True Then         ' there was an extension
        ExtStrip = OutString        ' return the string without it
    Else                            ' there was no extension
        ExtStrip = InString         ' return the original string.
    End If

End Function

Function ExtReturn(InString As String) As String
'*******************************************************
' Purpose:  Takes a string as input and returns any 1,2,3 or 4 letter extension
' Inputs:   The string being changed.   Returns: a string of a file suffix
'*******************************************************
On Error GoTo ExtReturnErr
    Dim strExt As String
    Dim iCount As Integer

    For iCount = 0 To Len(InString)
        strExt = Mid(InString, Len(InString) - iCount, 1 + iCount)
        If Left(strExt, 1) = "." Then
            ExtReturn = Right(strExt, Len(strExt) - 1)
            Exit For
        End If
    Next
    Exit Function
ExtReturnErr:
    Select Case Err.Number
        Case 5
            Exit Function
        Case Else
        Resume Next
    End Select
End Function

Function FileNameReturn(InString As String) As String
'*******************************************************
' Purpose:  Takes a string as input and returns any Filename & extension (removes the path)
' Inputs:   The string being changed.   Returns: a string of a file suffix
'*******************************************************
On Error GoTo ExtReturnErr
    Dim strExt As String
    Dim iCount As Integer

    For iCount = 0 To Len(InString)
        strExt = Mid(InString, Len(InString) - iCount, 1 + iCount)
        If Left(strExt, 1) = "\" Then
            FileNameReturn = Right(strExt, Len(strExt) - 1)
            Exit For
        End If
    Next
    Exit Function
ExtReturnErr:
    Select Case Err.Number
        Case 5
            Exit Function
        Case Else
        Resume Next
    End Select
End Function
Public Function TitleCaps(InString As String) As String
'*******************************************************
' Purpose:  Takes a string as input and returns It in Title Case
' Inputs:   The string being changed.   Returns: a formatted string
'*******************************************************
Dim OutString As String, CurrentLetter As String
Dim CurrentWord As String, TCaps As String
Dim StrCount As Integer, i As Byte
    
    If InString = "" Then
         TitleCaps = ""
         Exit Function
    End If
    
    For StrCount = 1 To Len(InString)
    CurrentLetter = Mid(InString, StrCount, 1)
    CurrentWord = CurrentWord + CurrentLetter
    If InStr(" .,/\;:-!?[]()#", CurrentLetter) <> 0 Or StrCount = Len(InString) Then
         TCaps = UCase(Left(CurrentWord, 1))
         For i = 2 To Len(CurrentWord)
         TCaps = TCaps & Mid(CurrentWord, i, 1)
         Next
         OutString = OutString & TCaps
         CurrentWord = ""
    End If
    Next
    
    TitleCaps = OutString
End Function

Public Function URLValidate(strURL As String) As Boolean
On Error GoTo URLValErr
    Dim strTest As String
    strTest = ExtReturn(strURL)
    If strTest = "com" Or strTest = "org" Or strTest = "net" Or strTest = "gov" Or strTest = "edu" Or strTest = "ivb" Or strTest = "asp" Or strTest = "htm" Or strTest = "html" Then
        URLValidate = False
    Else
        URLValidate = True
        Exit Function
    End If
    Exit Function
URLValErr:
    URLValidate = False
End Function
