VERSION 5.00
Begin VB.Form frmtblTipType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "tblTipType"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Tag             =   "tblTipType"
   Begin VB.CommandButton cmdGrid 
      Caption         =   "&Grid"
      Height          =   300
      Left            =   4440
      TabIndex        =   8
      Tag             =   "&Grid"
      Top             =   700
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   300
      Left            =   3360
      TabIndex        =   7
      Tag             =   "&Update"
      Top             =   700
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   2280
      TabIndex        =   6
      Tag             =   "&Refresh"
      Top             =   700
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   300
      Left            =   1200
      TabIndex        =   5
      Tag             =   "&Delete"
      Top             =   700
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Tag             =   "&Add"
      Top             =   700
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "D:\BlackBox\VBTips.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblTipType"
      Top             =   1140
      Width           =   5550
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "TipType"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   1
      Top             =   40
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "ID:"
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "TipType:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "TipType:"
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmtblTipType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub


Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
End Sub


Private Sub cmdRefresh_Click()
    'this is really only needed for multi user apps
    Data1.Refresh
End Sub


Private Sub cmdUpdate_Click()
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub


Private Sub cmdGrid_Click()
    On Error GoTo cmdGrid_ClickErr


    Dim f As New frmDataGrid
    Set f.Data1.Recordset = Data1.Recordset
    f.Caption = Me.Caption & " Grid"
    f.Show


    Exit Sub
cmdGrid_ClickErr:
End Sub


Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    'This is where you would put error handling code
    'If you want to ignore errors, comment out the next line
    'If you want to trap them, add code here to handle them
    MsgBox "Data error event hit err:" & Error$(DataErr)
    Response = 0  'throw away the error
End Sub


Private Sub Data1_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'This will display the current record position
    'for dynasets and snapshots
    Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
    'for the table object you must set the index property when
    'the recordset gets created and use the following line
    'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub


Private Sub Data1_Validate(Action As Integer, Save As Integer)
    'This is where you put validation code
    'This event gets called when the following actions occur
    Select Case Action
        Case vbDataActionMoveFirst
        Case vbDataActionMovePrevious
        Case vbDataActionMoveNext
        Case vbDataActionMoveLast
        Case vbDataActionAddNew
        Case vbDataActionUpdate
        Case vbDataActionDelete
        Case vbDataActionFind
        Case vbDataActionBookmark
        Case vbDataActionClose
            Screen.MousePointer = vbDefault
    End Select
    Screen.MousePointer = vbHourglass
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub



