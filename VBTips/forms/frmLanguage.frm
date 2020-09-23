VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLanguage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Language Category Maintenance"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmLanguage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Grid Example Using Bound Controls"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblLanguage"
      Top             =   5160
      Width           =   3855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmLanguage.frx":0442
      Height          =   3495
      Left            =   360
      OleObjectBlob   =   "frmLanguage.frx":0456
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   "Note! Each Record must be filled in completely."
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "From this form you may modify the entries in the Language Type drop-down combo box as well as the tip type option buttons. "
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
    Call StatusMsgDisplay("The Language Maintenance form is active.", 2)
End Sub

Private Sub Form_Load()
    With Data1
        .ToolTipText = "For most applications it is not robust enough to use, but in a situation like this where we only wish to modify a recordset and not add to it, it works just fine."
        .DatabaseName = App.path & "\databases\VBTips.mdb"
        .RecordSource = "tblLanguage"
    End With
    With Me
        .Top = (Forms.Count - 1) * 200
        .Left = (Forms.Count - 1) * 200
    End With
    If Forms.Count > 2 Then fMDI.Arrange vbCascade
End Sub

Private Sub Form_Resize()
Select Case Me.WindowState
    Case vbNormal
        With Me
            .Width = 6555
            .Height = 6315
        End With
        
        With DBGrid1
            .Columns(0).Width = (DBGrid1.Width * 0.15)
            .Columns(1).Width = (DBGrid1.Width * 0.3)
            .Columns(2).Width = (DBGrid1.Width * 0.5)
        End With
    Case vbMaximized
        With DBGrid1
            .Columns(0).Width = (DBGrid1.Width * 0.15)
            .Columns(1).Width = (DBGrid1.Width * 0.3)
            .Columns(2).Width = (DBGrid1.Width * 0.5)
        End With
    Case Else    ' do nothing (for now)
End Select
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    If g_blnUnload = True Then Exit Sub
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub
