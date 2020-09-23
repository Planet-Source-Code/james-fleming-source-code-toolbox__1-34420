VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   0  'None
   Caption         =   "Progress"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProgCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3840
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1560
      Left            =   15
      Top             =   15
      Width           =   4335
   End
   Begin VB.Label lblProgress 
      Caption         =   "Exporting Functions and Subs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstt As Boolean
Dim fClipboard As frmClipboard ' Read
Public myIndex As Integer
Private Sub cmdProgCancel_Click()
    ProgressCancel = True
End Sub

Private Sub Form_Load()
'*******************************************************
' Purpose:  to offer feedback to the user that something is happening
' inputs :  myIndex.       Returns: none
' Comment:  This is based on a project example from planet-source-code.
'           The original used 3 different progress boxes, I chose to use one.
'           How I accomplished this was by passing a variable before the form loads
' Comment:  This is currently designed to input VB functions and subs. It can easily be expanded to input more functions of other types.
'           But since I do mostly VB, I will not be adding the input routines for the other types until a later release. I did however
'           illustrate how to set a select statement for checking against multiple values.
' Depends:  On having the value (myIndex) set by the calling routine.
'*******************************************************
On Error GoTo myErrHandler
    Call FormCenter(Me) ' center the form
    Select Case myIndex
        Case 0  ' Export
            lblProgress = "Exporting Functions and Subs"
            firstt = True
        Case 1  ' Search
            firstt = True
            lblProgress = "File Search in Progress"
        Case Else ' Read
            lblProgress = "Reading :"
            Set fClipboard = New frmClipboard
            ProgressBar1.Value = 0
            ProgressCancel = False
            Call StatusMsgDisplay("Reading :" & FileFSToOpen, 1)
    End Select
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)  ' call the error message box
    Resume Next                                     ' continue
End Sub

Private Sub Form_Paint()
'*******************************************************
' Purpose:  Fires the Timer event.
' inputs :  myIndex.       Returns: none
' Comment:  See Form_load for explanation of myIndex
'*******************************************************
 On Error GoTo myErrHandler
    Select Case myIndex
        Case 0  ' Export
            If firstt Then Timer1.Enabled = True
        Case 1  ' Search
            If firstt Then Timer1.Enabled = True
        Case Else  ' Read
            Timer1.Enabled = True
    End Select
     On Error GoTo myErrHandler
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)  ' call the error message box
    Resume Next                                     ' continue
End Sub

Private Sub Timer1_Timer()
'*******************************************************
' Purpose:  The timer event purpose is slightly different for each myIndex value.
'          It performs either a search, copy, read in or export.
' inputs :  myIndex.       Returns: none
' Comment:  See Form_load for explanation of myIndex
'*******************************************************
 On Error GoTo myErrHandler
    Select Case myIndex
        Case 0  ' Export
            Progress
            Timer1.Enabled = False
        Case 1  ' Search
            firstt = False
            fSearch.FindFile (fSearch.txtSearch.Text) & "\", fSearch.cmbExt.Text
            Timer1.Enabled = False
            Unload Me
        Case 2  ' Read
            Timer1.Enabled = False
            If TypeView = 0 Then
                Call LoadFunction
                Unload Me
                fClipboard.Show
            Else
                Call LoadFile
                Unload Me
                fClipboard.Show
            End If
        Case Else ' just copy the data without viewing it.
            Timer1.Enabled = False
            If TypeView = 0 Then
                Call LoadFunction
                Unload Me
                fClipboard.cmdClipboard_Click (1)
                fClipboard.cmdClipboard_Click (4)
            Else
                Call LoadFile
                Unload Me
                fClipboard.cmdClipboard_Click (1)
                fClipboard.cmdClipboard_Click (4)
            End If
    End Select
    Unload Me

    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in Timer of frmProgress.")  ' call the error message box
    Resume Next                                     ' continue
End Sub
'*************************
'  Export
'*************************
Private Sub Progress()
Dim i As Integer
Dim textline As String
    If NbFile < 2 Or fSearch.optFunct(1).Value = True Then
       Call ReturnOneFile
       Exit Sub
    End If
    ProgressCancel = False
    firstt = False
    ProgressBar1.Min = 1
    ProgressBar1.Max = NbFile
    ProgressBar1.Value = 1
    Dim file1, PathFile
    file1 = FreeFile
    For i = 1 To NbFile
        If ProgressCancel Then Unload Me: Exit Sub
        fSearch.msgSearch(0).Row = i
        fSearch.msgSearch(0).Col = 1
        PathFile = fSearch.msgSearch(0).Text
        fSearch.msgSearch(0).Col = 0
        PathFile = PathFile & fSearch.msgSearch(0).Text
        Call StatusMsgDisplay("File : " & PathFile, 1)
        Call StatusMsgDisplay(i & " " & UCase(Right(fSearch.cmbExt, Len(fSearch.cmbExt) - 2)) & " Procedures Found.", 2)
        Open PathFile For Input As file1
        Do While Not EOF(file1) ' Loop until end of file.
            Line Input #file1, textline ' Read line into variable.
            If fSearch.Check1.Value = 0 And InStr(1, textline, "_") <> 0 Then GoTo skip1
            If InStr(1, textline, "Public Sub") <> 0 Then
                modFindAllFiles.EndStatement = "End Sub"
                fSearch.AddFunctionSub Trim(textline), PathFile
            ElseIf InStr(1, textline, "Public Function") <> 0 Then
                modFindAllFiles.EndStatement = "End Function"
                fSearch.AddFunctionSub Trim(textline), PathFile
            ElseIf InStr(1, textline, "Private Function") <> 0 Then
                fSearch.AddFunctionSub Trim(textline), PathFile
                modFindAllFiles.EndStatement = "End Function"
            ElseIf InStr(1, textline, "Private Sub") <> 0 Then
                modFindAllFiles.EndStatement = "End Sub"
                fSearch.AddFunctionSub Trim(textline), PathFile
            End If
skip1:
        DoEvents
        Loop
    Close file1
    If ProgressBar1.Value < NbFile Then ProgressBar1.Value = ProgressBar1.Value + 1
    DoEvents
    Next
End Sub
Private Sub ReturnOneFile()
On Error GoTo ReturnOneFileErr
Dim i As Integer
    Dim oneNbFile As Integer
    oneNbFile = 1
    ProgressCancel = False
    firstt = False
    ProgressBar1.Min = 1
    ProgressBar1.Max = 2
    ProgressBar1.Value = 1
    Dim file1, PathFile
    file1 = FreeFile
    For i = 1 To oneNbFile
     If ProgressCancel Then Unload Me: Exit Sub
     fSearch.msgSearch(0).Row = i
     fSearch.msgSearch(0).Col = i
     If NbFile = 1 Then
        PathFile = fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 1) & fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 0)
    Else
        PathFile = fSearch.txtSearch.Text
     End If
     Call StatusMsgDisplay("File : " & PathFile, 1)
     Call StatusMsgDisplay(i & " " & UCase(Right(fSearch.cmbExt, Len(fSearch.cmbExt) - 2)) & " Files Found.", 2)
     Open PathFile For Input As file1
      Dim textline
      Do While Not EOF(file1) ' Loop until end of file.
        Line Input #file1, textline ' Read line into variable.
         If fSearch.Check1.Value = 0 And InStr(1, textline, "_") <> 0 Then GoTo skip1
           If InStr(1, textline, "Public Sub") <> 0 Then
             fSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Public Function") <> 0 Then
             fSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Private Function") <> 0 Then
             fSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Private Sub") <> 0 Then
             fSearch.AddFunctionSub Trim(textline), PathFile
           End If

skip1:
         DoEvents
       Loop
     
     Close file1
     If ProgressBar1.Value < oneNbFile Then ProgressBar1.Value = ProgressBar1.Value + 1
     DoEvents
    Next
    Unload Me
    Exit Sub
ReturnOneFileErr:
    Select Case Err.Number
        Case 52 ' bad file name
            InputErrBox (Err.Description)
            PathFile = fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 1) & fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 0)
            Exit Sub
        Case 53 ' file not found
            InputErrBox ("The file you specified is not found." & vbCrLf & "I will return the first of the files in the subset.")
            oneNbFile = NbFile
            PathFile = fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 1) & fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 0)
            Open PathFile For Input As file1
            Resume Next
        Case 380 ' invalid property value
        Resume Next
        Case Else
            Resume Next
    End Select
End Sub

Public Sub LoadFile()
 On Error GoTo loadFileErr
 Dim file1, textline As String, strTemp As String, pos As Long
 file1 = FreeFile
 ProgressBar1.Max = FileLen(FileFSToOpen)
 pos = 0
 Open FileFSToOpen For Input As file1
  Do While Not EOF(file1) ' Loop until end of file.
    If ProgressCancel Then Exit Do
    Line Input #file1, textline ' Read line into variable.
    textline = textline & Chr(13) & Chr(10)
    fClipboard.txtClipboard.Text = fClipboard.txtClipboard.Text & textline & Chr(13) & Chr(10)
        ProgressBar1.Value = ProgressBar1.Value + Len(textline)
    DoEvents
    pos = pos + Len(textline)
  Loop
Close #file1
Exit Sub
loadFileErr:
    Select Case Err.Number
        Case 380
            ' Invalid property value
            'ErrMsgBox ("I am unable to open the requested file. It is too large.")
            pos = pos + Len(textline)
            Close #file1
            Resume Next
        Case 52
            Exit Sub
        Case Else
            Resume Next
    End Select
End Sub

