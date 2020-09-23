VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmTipType 
   Caption         =   "Tip Type Maintenance"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   Icon            =   "frmTipType.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   6435
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Data Data1 
      Caption         =   "Grid Example Using Bound Controls"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\DevStudio\VB\VBTips\databases\VBTips.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1380
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblTipType"
      Top             =   5160
      Width           =   3855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmTipType.frx":0442
      Height          =   3495
      Left            =   360
      OleObjectBlob   =   "frmTipType.frx":0456
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Label Label2 
      Caption         =   $"frmTipType.frx":0FF5
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "From this form you may modify the entries in the Tip Type drop-down combo box as well as the tip type option buttons. "
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmTipType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
        Call StatusMsgDisplay("The Tip Type form is active.", 2)
        Data1.ToolTipText = "For most applications it is not robust enough to use, but in a situation like this where we only wish to modify a recordset and not add to it, it works just fine."
End Sub

Private Sub Form_Resize()
Select Case Me.WindowState
    Case vbNormal
        With Me
            .Width = 6555
            .Height = 6315
        End With

    Case Else    ' do nothing (for now)
        
    End Select
    DBGrid1.Columns(0).Width = (DBGrid1.Width * 0.15)
    DBGrid1.Columns(1).Width = (DBGrid1.Width * 0.3)
    DBGrid1.Columns(2).Width = (DBGrid1.Width * 0.5)
    Exit Sub
End Sub
