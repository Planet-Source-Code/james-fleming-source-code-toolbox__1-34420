VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportProg 
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
   Begin VB.CommandButton Command1 
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
   Begin VB.Label Label1 
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
Attribute VB_Name = "frmExportProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstt As Boolean

Private Sub ExportProg_Click()
    ProgressCancel = True
End Sub

Private Sub Form_Load()
firstt = True
    Call FormCenter(Me)

End Sub

Private Sub Form_Paint()
    If firstt Then Timer1.Enabled = True
End Sub
Private Sub Progress()
Dim i As Integer
    If NbFile < 2 Or frmSearch.optFunct(1).Value = True Then
        Call ReturnOneFile
        Exit Sub
    End If
    ProgressCancel = False
    firstt = False
    frmExportProg.ProgressBar1.Min = 1
    frmExportProg.ProgressBar1.Max = NbFile
    frmExportProg.ProgressBar1.Value = 1
    Dim file1, PathFile
    file1 = FreeFile
    For i = 1 To NbFile
     If ProgressCancel Then Unload Me: Exit Sub
     frmSearch.msgSearch(0).Row = i
     frmSearch.msgSearch(0).Col = 1
     PathFile = frmSearch.msgSearch(0).Text
     frmSearch.msgSearch(0).Col = 0
     PathFile = PathFile & frmSearch.msgSearch(0).Text
     Call StatusMsgDisplay("File : " & PathFile, 1)
     Call StatusMsgDisplay(i & " " & UCase(Right(frmSearch.cmbExt, Len(frmSearch.cmbExt) - 2)) & " Procedures Found.", 2)
     Open PathFile For Input As file1
      Dim textline
      Do While Not EOF(file1) ' Loop until end of file.
        Line Input #file1, textline ' Read line into variable.
         If frmSearch.Check1.Value = 0 And InStr(1, textline, "_") <> 0 Then GoTo skip1
           If InStr(1, textline, "Public Sub") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Public Function") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Private Function") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Private Sub") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           End If

skip1:
         DoEvents
       Loop
     
     Close file1
     If ProgressBar1.Value < NbFile Then ProgressBar1.Value = ProgressBar1.Value + 1
     DoEvents
    Next
    Unload Me
End Sub
Private Sub ReturnOneFile()
On Error GoTo ReturnOneFileErr
Dim i As Integer
    Dim oneNbFile As Integer
    oneNbFile = 1
    ProgressCancel = False
    firstt = False
    frmExportProg.ProgressBar1.Min = 1
    frmExportProg.ProgressBar1.Max = 2
    frmExportProg.ProgressBar1.Value = 1
    Dim file1, PathFile
    file1 = FreeFile
    For i = 1 To oneNbFile
     If ProgressCancel Then Unload Me: Exit Sub
     frmSearch.msgSearch(0).Row = i
     frmSearch.msgSearch(0).Col = i
     If NbFile = 1 Then
        PathFile = frmSearch.msgSearch(0).TextMatrix(frmSearch.msgSearch(0).Row, 1) & frmSearch.msgSearch(0).TextMatrix(frmSearch.msgSearch(0).Row, 0)
    Else
        PathFile = frmSearch.txtSearch.Text
     End If
     Call StatusMsgDisplay("File : " & PathFile, 1)
     Call StatusMsgDisplay(i & " " & UCase(Right(frmSearch.cmbExt, Len(frmSearch.cmbExt) - 2)) & " Files Found.", 2)
     Open PathFile For Input As file1
      Dim textline
      Do While Not EOF(file1) ' Loop until end of file.
        Line Input #file1, textline ' Read line into variable.
         If frmSearch.Check1.Value = 0 And InStr(1, textline, "_") <> 0 Then GoTo skip1
           If InStr(1, textline, "Public Sub") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Public Function") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Private Function") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
           ElseIf InStr(1, textline, "Private Sub") <> 0 Then
             frmSearch.AddFunctionSub Trim(textline), PathFile
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
            PathFile = frmSearch.msgSearch(0).TextMatrix(frmSearch.msgSearch(0).Row, 1) & frmSearch.msgSearch(0).TextMatrix(frmSearch.msgSearch(0).Row, 0)
            Exit Sub
        Case 53 ' file not found
            InputErrBox ("The file you specified is not found." & vbCrLf & "I will return the first of the files in the subset.")
            oneNbFile = NbFile
            PathFile = frmSearch.msgSearch(0).TextMatrix(frmSearch.msgSearch(0).Row, 1) & frmSearch.msgSearch(0).TextMatrix(frmSearch.msgSearch(0).Row, 0)
            Open PathFile For Input As file1
            Resume Next
        Case 380 ' invalid property value
        Resume Next
        Case Else
            Resume Next
    End Select
End Sub
Private Sub Timer1_Timer()
Progress
Timer1.Enabled = False
End Sub
