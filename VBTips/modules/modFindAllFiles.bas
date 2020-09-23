Attribute VB_Name = "modFindAllFiles"
    Public Const MAX_PATH As Long = 260
    Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
    Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
    Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
    Public Const FILE_ATTRIBUTE_HIDDEN = &H2
    Public Const FILE_ATTRIBUTE_NORMAL = &H80
    Public Const FILE_ATTRIBUTE_READONLY = &H1
    Public Const FILE_ATTRIBUTE_SYSTEM = &H4
    Public Const FILE_ATTRIBUTE_TEMPORARY = &H100


    Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
    End Type
    
    Type SaveF
        StingToSave As String
    End Type

    Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
    End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public NbFile As Long
Public FileFSToOpen As String
Public StringToFind As String
Public ProgressCancel As Boolean
Public TypeView As Integer
Public EndStatement As String
Public Function StripNull(ByVal WhatStr As String) As String
On Error GoTo StripNullErr
    Dim pos As Integer
    pos = InStr(WhatStr, Chr$(0))
    If pos > 0 Then
        StripNull = Left$(WhatStr, pos - 1)
    Else
        StripNull = WhatStr
    End If
    Exit Function
StripNullErr:
    ErrMsgBox (Err.Description & " " & Err.Number & " in StripNull of modFindAllFiles")
End Function
         
Public Sub LoadFunction()
On Error GoTo LoadErr
 Dim file1
 file1 = FreeFile
 EndStatement = "End Sub"
 If InStr(1, StringToFind, "Function") <> 0 Then EndStatement = "End Function"
 frmProgress.ProgressBar1.Max = FileLen(FileFSToOpen)
 Open FileFSToOpen For Input As file1
  Dim textline
  Do While Not EOF(file1) ' Loop until end of file.
    If ProgressCancel Then Exit Do
    Line Input #file1, textline ' Read line into variable.
    If InStr(1, textline, StringToFind, vbTextCompare) <> 0 Then
      frmClipboard.txtClipboard.Text = frmClipboard.txtClipboard.Text & textline & Chr(13) & Chr(10)
      Do
        Line Input #file1, textline ' Read line into variable.
        frmClipboard.txtClipboard.Text = frmClipboard.txtClipboard.Text & textline & Chr(13) & Chr(10)
        If Trim(textline) = EndStatement Then Close file1: Exit Sub
      Loop
    frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + Len(textline)
    DoEvents
    End If
    frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + Len(textline)
    DoEvents
  Loop
 Close file1
 Exit Sub
LoadErr:
    Select Case Err.Number
        Case 62 ' Input past end of file
            Close file1
            ErrMsgBox (Err.Description & " " & Err.Number & " in LoadFunction of ModFindAllFiles")
            Resume Next
        Case 52
            ErrMsgBox ("The file being opened has an inproperly terminated end of file." & vbCrLf & "See LoadFunction of ModFindAllFiles")
            Exit Sub
        Case Else
            ErrMsgBox (Err.Description & " " & Err.Number & " in LoadFunction of ModFindAllFiles")
            Resume Next
    End Select
End Sub


