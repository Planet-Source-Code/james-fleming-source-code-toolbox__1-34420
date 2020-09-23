VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReadProg 
   BorderStyle     =   0  'None
   Caption         =   "Reading File"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4365
   LinkTopic       =   "Form6"
   ScaleHeight     =   1830
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   3720
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1185
      TabIndex        =   0
      Top             =   1320
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Reading :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1800
      Left            =   10
      Top             =   10
      Width           =   4335
   End
End
Attribute VB_Name = "frmReadProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormTemp As Form5
Private Sub Command1_Click()
ProgressCancel = True
End Sub
Private Sub LoadFile()
 'On Error Resume Next
 Dim file1, textline As String, strtemp As String, pos As Long
 file1 = FreeFile
 ProgressBar1.Max = FileLen(FileFSToOpen)
 pos = 0
 Open FileFSToOpen For Input As file1
  Do While Not EOF(file1) ' Loop until end of file.
    If ProgressCancel Then Exit Do
    Line Input #file1, textline ' Read line into variable.
    textline = textline & Chr(13) & Chr(10)
    FormTemp.Text1.Text = FormTemp.Text1.Text & textline & Chr(13) & Chr(10)
        ProgressBar1.Value = ProgressBar1.Value + Len(textline)
    DoEvents
    pos = pos + Len(textline)
  Loop
Close #file1
End Sub
Private Sub LoadFunction()
 Dim EndStatement As String, file1
 file1 = FreeFile
 EndStatement = "End Sub"
 If InStr(1, StringToFind, "Function") <> 0 Then EndStatement = "End Function"
 ProgressBar1.Max = FileLen(FileFSToOpen)
 Open FileFSToOpen For Input As file1
  Dim textline
  Do While Not EOF(file1) ' Loop until end of file.
    If ProgressCancel Then Exit Do
    Line Input #file1, textline ' Read line into variable.
    If InStr(1, textline, StringToFind, vbTextCompare) <> 0 Then
      FormTemp.Text1.Text = FormTemp.Text1.Text & textline & Chr(13) & Chr(10)
      Do
        Line Input #file1, textline ' Read line into variable.
        FormTemp.Text1.Text = FormTemp.Text1.Text & textline & Chr(13) & Chr(10)
        If Trim(textline) = EndStatement Then Close file1: Exit Sub
      Loop
    ProgressBar1.Value = ProgressBar1.Value + Len(textline)
    DoEvents
    End If
    ProgressBar1.Value = ProgressBar1.Value + Len(textline)
    DoEvents
  Loop
 Close file1
End Sub

Private Sub Form_Load()
Set FormTemp = New Form5
ProgressBar1.Value = 0
ProgressCancel = False
Top = frmSearch.Top + ((frmSearch.Height / 2)) - (Height / 2)
Left = (frmSearch.Left + (frmSearch.Width / 2)) - (Width / 2)
lblSearch(0).Caption = "Reading :" & FileFSToOpen
End Sub

Private Sub Form_Paint()
Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
Timer1.Enabled = False
If TypeView = 0 Then
 LoadFunction
 Unload Me
 FormTemp.Show
Else
 LoadFile
 Unload Me
 FormTemp.Show
End If
End Sub
