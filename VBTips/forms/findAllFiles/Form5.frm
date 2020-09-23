VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Caption         =   "Description"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4635
   LinkTopic       =   "Form5"
   ScaleHeight     =   5805
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   5400
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
MsgBox "Put copy to clipboard code HERE!"
End Sub

Private Sub Command3_Click()
Dim file1
file:
CommonDialog1.Filter = "(*.txt)|*.txt|"
CommonDialog1.ShowSave
If CommonDialog1.filename = "" Then
  MsgBox "You did not enter any file name !"
  Exit Sub
Else
Dim Strtemp, rep2
rep2 = Dir(CommonDialog1.filename)
Strtemp = Mid(CommonDialog1.filename, (Len(CommonDialog1.filename) - Len(rep2)) + 1, Len(rep2))
 If rep2 = Strtemp Then
   Dim rep
   rep = MsgBox("The file already exist do you want to replace it ?", vbQuestion + vbYesNoCancel, "File Overwrite !!")
   If rep = vbNo Then
     GoTo file:
   ElseIf rep = vbYes Then
     Kill CommonDialog1.filename
   Else
     Exit Sub
   End If
 End If
 file1 = FreeFile
 EndStatement = "End Sub"
 If InStr(1, StringToFind, "Function") <> 0 Then EndStatement = "End Function"
 Open CommonDialog1.filename For Output As file1
 Write #file1, Text1.Text
 Close file1
  
End If
 

End Sub

Private Sub Form_Load()
 Me.Caption = StringToFind
End Sub


Private Sub Form_Resize()
If Me.WindowState <> 1 Then
Text1.Width = Width - 150
Text1.Height = Height - 1200
Command1.Top = Height - 915
Command2.Top = Height - 915
Command3.Top = Height - 915
End If
End Sub
