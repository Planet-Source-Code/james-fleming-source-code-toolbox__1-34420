VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmClipboard 
   Caption         =   "Code Clipboard"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4635
   Icon            =   "frmClipboard.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   5805
   ScaleWidth      =   4635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Print"
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   5
      ToolTipText     =   "Send this code to the default printer."
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Add"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Add this code to database."
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Save"
      Height          =   375
      Index           =   2
      Left            =   1860
      TabIndex        =   3
      ToolTipText     =   "Save this code to a new file."
      Top             =   5400
      Width           =   735
   End
   Begin MSComDlg.CommonDialog dlgClipboard 
      Left            =   840
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontSize        =   12
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Copy"
      Height          =   375
      Index           =   1
      Left            =   990
      TabIndex        =   2
      ToolTipText     =   "Copy this code to clipboard."
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Quit"
      Height          =   375
      Index           =   4
      Left            =   3600
      TabIndex        =   1
      ToolTipText     =   "Close this form"
      Top             =   5400
      Width           =   735
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
Dim fLoad As Boolean
Dim fDirty As Boolean

Public Sub cmdClipboard_Click(Index As Integer)

    Select Case Index
        Case 0
            ftblTips.cmdTip_Click (2)
            ftblTips.txtFields(4).Text = txtClipboard.Text
            ftblTips.txtFields(0).SetFocus
        Case 1
            txtClipboard.SelStart = 0
            txtClipboard.SelLength = txtClipboard.MaxLength
            Clipboard.SetText txtClipboard.SelText
            txtClipboard.SelStart = 0
        Case 2
            Call FileSave(dlgClipboard, txtClipboard)
        Case 3
            Call PrintField(txtClipboard, dlgClipboard)
        Case 4
            Unload Me
    End Select

End Sub

Private Sub Form_Activate()
    fLoad = False
    Call StatusMsgDisplay("The Clipboard is active.", 2)
End Sub

Private Sub Form_Load()

    If StringToFind <> "" Then
        Caption = StringToFind
    Else
        Caption = "Code Clipboard"
    End If
    fMDI.mnuFilePrint.Enabled = False
    fMDI.tbToolbar.Buttons(4).Enabled = False
    fLoad = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If fDirty = True Then
     '    Dim intResponse As Integer
     '   intResponse = YesNo("The text has been changed. Save changes?")
     '   If intResponse = 1 Then Me.rtbFileView.SaveFile fMDI.dlgMDI.FileName, rtfText
    End If
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    Dim ileft As Integer
    
    If Me.WindowState <> 1 Then
        If Width < 2700 Then Width = 2700
        txtClipboard.Width = Me.ScaleWidth ' - 150
        txtClipboard.Height = Height - 1200
        ileft = txtClipboard.Left
         Do While i < 5
            With cmdClipboard(i)
                .Move ileft, (Height - 915), (txtClipboard.Width * 0.2)
            End With
            ileft = cmdClipboard(i).Left + cmdClipboard(i).Width
            i = i + 1
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fMDI.mnuFilePrint.Enabled = True
    fMDI.tbToolbar.Buttons(4).Enabled = True
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub

Private Sub txtClipboard_Change()
    If fLoad = False Then
        fDirty = True
    End If
End Sub
