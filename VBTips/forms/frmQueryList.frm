VERSION 5.00
Begin VB.Form frmQueryList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Querys"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Tag             =   "Querys"
   Begin VB.ListBox lstQueryDefs 
      Height          =   1815
      Left            =   96
      TabIndex        =   0
      Top             =   274
      Width           =   3392
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   370
      Left            =   3570
      TabIndex        =   2
      Tag             =   "&Close"
      Top             =   775
      Width           =   1440
   End
   Begin VB.CommandButton cmdExecuteQuery 
      Caption         =   "&Execute"
      Enabled         =   0   'False
      Height          =   370
      Left            =   3570
      TabIndex        =   1
      Tag             =   "&Execute"
      Top             =   274
      Width           =   1440
   End
   Begin VB.Label lblSQL 
      Caption         =   "Saved Querys:"
      Height          =   251
      Index           =   0
      Left            =   108
      TabIndex        =   3
      Tag             =   "Saved Querys:"
      Top             =   24
      Width           =   2189
   End
End
Attribute VB_Name = "frmQueryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mdbDatabase As Database
Private Sub Form_Load()
    Set mdbDatabase = OpenDatabase(gsDatabase)
    RefreshQuerys
    Me.Left = GetSetting(App.Title, "Settings", "QueryLeft", 0)
    Me.Top = GetSetting(App.Title, "Settings", "QueryTop", 0)
End Sub





Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdExecuteQuery_Click()
    Dim rsTmp As Recordset
    Dim dbTmp As Database
    Dim qdfTmp As QueryDef
    Dim bSavedQDF As Boolean
    Dim sSQL As String
    

    Set dbTmp = OpenDatabase(gsDatabase)
    

    If lstQueryDefs.ListIndex < 0 Then Exit Sub
    

    sSQL = dbTmp.QueryDefs(lstQueryDefs.Text).SQL
    Set qdfTmp = dbTmp.QueryDefs(lstQueryDefs.Text)
    

    If Not SetQryParams(qdfTmp) Then Exit Sub
    

    Screen.MousePointer = vbHourglass
    

    If UCase(Mid(sSQL, 1, 6)) = "SELECT" And InStr(UCase(sSQL), " INTO ") = 0 Then
        On Error GoTo SQLErr
MakeDynaset:
        Dim f As New frmDataGrid
        Set rsTmp = qdfTmp.OpenRecordset()
        Set f.Data1.Recordset = rsTmp
        If bSavedQDF Then
            f.Caption = qdfTmp.Name
        Else
            f.Caption = Left(sSQL, 32) & "..."
        End If
        f.Show
    Else
        On Error GoTo SQLErr
        qdfTmp.Execute
    End If


    Screen.MousePointer = vbDefault
    Exit Sub


SQLErr:
    If Err = 3065 Or Err = 3078 Then 'row returning or name not found so try to create recordset
        Resume MakeDynaset
    End If
    MsgBox Err.Description


SQLEnd:


End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "QueryLeft", Me.Left
        SaveSetting App.Title, "Settings", "QueryTop", Me.Top
    End If
End Sub


Private Sub lstQueryDefs_Click()
    cmdExecuteQuery.Enabled = True
End Sub


Private Sub lstQueryDefs_DblClick()
    cmdExecuteQuery_Click
End Sub


Sub RefreshQuerys()
    Dim qdf As QueryDef
    

    lstQueryDefs.Clear
    

    For Each qdf In mdbDatabase.QueryDefs
        lstQueryDefs.AddItem qdf.Name
    Next
    

End Sub


Private Function SetQryParams(rqdf As QueryDef) As Boolean
    On Error GoTo SPErr
    

    Dim prm As Parameter
    Dim sTmp As String
    Dim i As Integer
    

    For Each prm In rqdf.Parameters
        'get the value from the user
        sTmp = InputBox("Enter Value for Parameter '" & prm.Name & "':")
        If Len(sTmp) = 0 Then
            'bail out if the user doesn't enter one of the params
            SetQryParams = False
            Exit Function
        End If
        'store the value
        prm.Value = CVar(sTmp)
    Next
    

    SetQryParams = True
    Exit Function
        

SPErr:
    MsgBox Err.Description
End Function

