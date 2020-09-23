VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClipboard 
   Caption         =   "Description"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4635
   LinkTopic       =   "Form5"
   ScaleHeight     =   5805
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Add"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Save"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4680
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Copy"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Quit"
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   1
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtClipboard 
      Height          =   5295
      Left            =   0
      MaxLength       =   65500
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fClipboard As frmClipboard
Public Sub cmdClipboard_Click(Index As Integer)
    Select Case Index
        Case 0
            ftblTips.cmdAdd_Click
            ftblTips.txtFields(4).Text = txtClipboard.Text
            ftblTips.txtFields(0).SetFocus
        Case 1
            txtClipboard.SelStart = 0
            txtClipboard.SelLength = txtClipboard.MaxLength
            Clipboard.SetText txtClipboard.SelText
            txtClipboard.SelStart = 0
        Case 2
            FileSave
        Case 3
            Unload Me
    End Select

End Sub

Private Sub FileSave()

    Dim file1
file:
    CommonDialog1.Filter = "(*.txt)|*.txt|"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
      MsgBox "You did not enter any file name !"
      Exit Sub
    Else
    Dim Strtemp, rep2
    rep2 = Dir(CommonDialog1.FileName)
    Strtemp = Mid(CommonDialog1.FileName, (Len(CommonDialog1.FileName) - Len(rep2)) + 1, Len(rep2))
        If rep2 = Strtemp Then
          Dim rep
          rep = MsgBox("The file already exist do you want to replace it ?", vbQuestion + vbYesNoCancel, "File Overwrite !!")
            If rep = vbNo Then
              GoTo file:
            ElseIf rep = vbYes Then
              Kill CommonDialog1.FileName
            Else
              Exit Sub
            End If
        End If
     file1 = FreeFile
     EndStatement = "End Sub"
     If InStr(1, StringToFind, "Function") <> 0 Then EndStatement = "End Function"
     Open CommonDialog1.FileName For Output As file1
     Write #file1, txtClipboard.Text
     Close file1
      
    End If
 
End Sub

Private Sub Form_Load()
 Me.Caption = StringToFind
End Sub


Private Sub Form_Resize()
    Dim i As Integer
    Dim ileft As Integer
    
    If Me.WindowState <> 1 Then
        If Width < 2700 Then Width = 2700
        txtClipboard.Width = Me.ScaleWidth ' - 150
        txtClipboard.Height = Height - 1200
        ileft = txtClipboard.Left
        While i < 4
            With cmdClipboard(i)
                .Top = Height - 915
                .Width = txtClipboard.Width * 0.25
                .Left = ileft
            End With
            ileft = cmdClipboard(i).Left + cmdClipboard(i).Width
            i = i + 1
        Wend
    End If
End Sub
