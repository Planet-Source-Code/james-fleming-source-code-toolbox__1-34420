VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmWeb 
   Caption         =   "Programming Web Sites"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   5940
   Begin VB.TextBox txtWeb 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   75
      TabIndex        =   7
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox txtWeb 
      Height          =   285
      Index           =   0
      Left            =   960
      MaxLength       =   75
      TabIndex        =   3
      Top             =   135
      Width           =   3375
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Add"
      Height          =   315
      Index           =   0
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Help"
      Height          =   315
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Delete"
      Height          =   315
      Index           =   2
      Left            =   4560
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   3
      ForeColor       =   -2147483641
      BackColorFixed  =   -2147483638
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin VB.Label lblSubType 
      Caption         =   "URL"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblSubType 
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
' Purpose: this form is used to demonstrate MSFlexgrid
'           It is used for adding web sites to the database.
' Assumes: inclusion of the flexgrid control
' Author:   James R. Fleming    Date: 3/5/2000
'*******************************************************

'**********************************************************************************************************
'
' Purpose: Declare all module level variable
'
'**********************************************************************************************************

Dim blnFrmLoad As Boolean, blnGrdLoad As Boolean, blnGrdFocus As Boolean
Dim mGrdRS As Recordset


'**********************************************************************************************************
'
'Purpose:   All Form subs and events
'
'**********************************************************************************************************

Private Sub Form_Load()
'*******************************************************
' Purpose: Driver for the form load
' Comment: This will be called from frmTblTips most often, so it
'           checks to see if it is loaded.
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler

'    Dim oRS As Recordset
    blnFrmLoad = True
    Call FormOpen                               ' set the default open properties
    Call GridRoutine
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub FormOpen()

    Me.WindowState = vbNormal
    Move 0, 0, fMDI.ScaleWidth / 2, fMDI.ScaleHeight
    MSFlexGrid1.Move (ScaleWidth * 0.01), MSFlexGrid1.Top, (ScaleWidth * 0.98), ScaleHeight - (MSFlexGrid1.Top + 50)
    cmdGrid(0).Move (MSFlexGrid1.Left + MSFlexGrid1.Width) - cmdGrid(0).Width, 100
    cmdGrid(1).Move cmdGrid(0).Left, cmdGrid(0).Top + cmdGrid(0).Height + 50
    cmdGrid(2).Move cmdGrid(0).Left, cmdGrid(1).Top + cmdGrid(0).Height + 50
    cmdGrid(2).Enabled = False
End Sub
Private Sub Form_Activate()
'*****************************************************
' Purpose:  this sub displays a message in the status bar then changes it after 4 secs.
' Assume:   The inclusion of StatusMsgFlip in modFormUtilities
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myActivateerr

    Call StatusMsgFlip("The web form is active", mGrdRS.RecordCount)
    blnFrmLoad = False                      ' reset the flag
    Exit Sub
    
myActivateerr:
Select Case Err.Number
    Case 91
            Call StatusMsgFlip("The web form is active", 0)
            Resume Next
    Case Else
        ErrMsgBox (Err.Description & " " & Err.Number)
        Resume Next
    End Select
End Sub
Private Sub Form_DblClick()
'*******************************************************
' Purpose: Reposition the form
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler
    Me.WindowState = vbNormal               ' no min or max forms!
    Me.Height = fMDI.ScaleHeight            ' scale to the mdi container
    Move (fMDI.Left + 60), (fMDI.Top + 50)  ' tweak it
    Exit Sub                                ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*******************************************************
' Purpose: Close the record set on the unload event
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler
    On Error Resume Next            ' no reason to stop now
    mGrdRS.Close                    ' close the recordset
    If g_blnUnload = True Then Exit Sub
    Call StatusMsgDisplay("There are no active forms", 2)
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub Form_Resize()
'*******************************************************
' Purpose: Handle resizing events, scale the grid.
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler
If blnFrmLoad = True Then Exit Sub
Select Case Me.WindowState
    Case vbMaximized
        MSFlexGrid1.Move MSFlexGrid1.Left, MSFlexGrid1.Top, ScaleWidth * 0.95, Height = ScaleHeight - (MSFlexGrid1.Top + 50)
    Case vbMinimized    ' do nothing (for now)
    Case Else
        Width = MSFlexGrid1.Left + MSFlexGrid1.Width + MSFlexGrid1.Left
        MSFlexGrid1.Height = ScaleHeight - (MSFlexGrid1.Top + 50)
    End Select
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
'*******************************************************
' Purpose: Update the status bar
' Assume:   The inclusion of StatusMsgDisplay in modFormUtilities
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler
    If g_blnUnload = True Then Exit Sub
    Call StatusMsgDisplay("There are no active forms", 2)
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
'*****************************************************
'
' All button and control events
'
'*****************************************************
Private Sub MSFlexGrid1_GotFocus()
'*******************************************************
' Purpose: Use the focus event to activate buttons
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler
     blnGrdFocus = True
    If blnGrdLoad = False Then Call MSFlexGrid1_Click
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub MSFlexGrid1_RowColChange()
'*******************************************************
' Purpose: If the grid is loading won't fire  the click event
' Inputs: None                      Returns: None
'*******************************************************
On Error GoTo myErrHandler
    If blnGrdLoad = False Then Call MSFlexGrid1_Click
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub MSFlexGrid1_Click()
On Error GoTo myErrHandler
    Dim itemp As Integer

    If MSFlexGrid1.Row <> itemp Then itemp = (MSFlexGrid1.Row - 1)
    With mGrdRS
        .AbsolutePosition = itemp
        txtWeb(0).Text = !strSiteName
        txtWeb(1).Text = !strURL
    End With
    cmdGrid(2).Enabled = True
    MSFlexGrid1.TabStop = True
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub cmdGrid_Click(Index As Integer)
'*****************************************************
' Purpose:  Call the events based on the index, caption and a flag set on (0).tag
' Assumes:  inclusion of modControls
' Inputs:   the Index of the current control      ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    Select Case Index
        Case 0                                  ' the Add/Cancel Button
            If cmdGrid(0).Caption = "Add" Then
                Call RowAdd                     ' visual cues of modification started
            Else                                ' the action was cancelled
                Call RowAddCancel               ' visual cues of modification canceled
                txtWeb(0).SetFocus              ' set focus back to the text box
            End If                              ' end if
        Case 1                                  ' the Edit/Update Button
            If cmdGrid(0).Tag = "Add" Then
                Call RowAddCommit               ' add a new record
            ElseIf cmdGrid(0).Tag = "Edit" Then
                Call RowUpdateCommit            ' update an existing record
            Else                                ' nothing to update yet!
                Call InputInfoBox("To add a web site:" & vbCrLf & "Press the Add button." & vbCrLf & "Add the Name and URL. Press update." & vbCrLf & vbLf & "To edit a web site:" & vbCrLf & "Select the row you wish to edit." & vbCrLf & "Modify the text. Press update." & vbCrLf & vbLf & "To delete a web site:" & vbCrLf & "Select the row you wish to delete." & vbCrLf & "Press Delete.")
            End If
        Case 2                                  ' the Delete Button
            Call RowDelete                      ' delete the record
            Call RowAddCancel                   ' visual cues of modification canceled
            txtWeb(0).SetFocus                  ' set focus back to the text box
    End Select
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

'**********************************************************************************************************
'
'Purpose:  All of the general subs and functions
'
'**********************************************************************************************************

Private Sub RowAdd()
On Error GoTo myErrHandler
    cmdGrid(0).Tag = "Add"           ' set the cmd tags & captions
    cmdGrid(0).Caption = "Cancel"
    cmdGrid(1).Caption = "Update"
    cmdGrid(2).Enabled = False
    
    txtWeb(0).Tag = txtWeb(0).Text   ' set any current values temporarily
    cmdGrid(0).Tag = "Add"
    With MSFlexGrid1
        .Enabled = False             ' toggle the grid
        .BackColorSel = vbWhite
        .ForeColorSel = &H80000011
        .ForeColor = &H80000011
        .ForeColorFixed = &H80000011
    End With
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowAddCommit()
On Error GoTo myErrHandler
    If textValidate = True Then
        Call RecordAdd              ' call the update subroutine
        Call GrdTglEnabled
        cmdGrid(2).Enabled = False
    End If
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub GrdTglEnabled()
On Error GoTo myErrHandler
    With MSFlexGrid1
        .ForeColor = &H80000007      ' toggle the grid
        .ForeColorFixed = &H80000012
        .BackColorSel = &H8000000D
        .ForeColorSel = &H8000000E
        .TabStop = True
        .Enabled = True
    End With
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowUpdateCommit()
On Error GoTo myErrHandler
    If textValidate = True Then
        Call RecordUpdate               ' call the update subroutine
        txtWeb(0) = ""                  ' clear the text boxes
        txtWeb(1) = ""                  ' clear the text boxes
        cmdGrid(0).Tag = ""             ' set the cmd tags & captions
        cmdGrid(0).Caption = "Add"
        cmdGrid(1).Caption = "Help"
        cmdGrid(2).Enabled = False
        Call GrdTglEnabled              ' toggle the grid
    End If
    Exit Sub                            ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowAddCancel()
On Error GoTo myErrHandler
    cmdGrid(0).Caption = "Add"      ' set the button properties
    cmdGrid(1).Caption = "Help"
    cmdGrid(2).Enabled = False
    cmdGrid(0).Tag = ""
    txtWeb(0).Tag = ""              ' clear the text boxes & tags
    txtWeb(0) = ""
    txtWeb(1).Tag = ""
    txtWeb(1) = ""
    Call GrdTglEnabled              ' toggle the grid
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowDelete()
On Error GoTo myErrHandler
    Dim intResponse As Integer
    intResponse = MsgBox("Delete Current Tip", vbYesNo + vbQuestion, "Delete")
    If intResponse = vbYes Then
        txtWeb(0) = ""
        Call RecordDelete
    Else
        intResponse = 0             ' skip it!
    End If                          ' end the nested if
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub GridMaint(flxgrd As MSFlexGrid)

On Error GoTo myErrHandler
    With flxgrd
        .Enabled = False
        .ColWidth(0) = .Width * 0.05
        .ColWidth(1) = .Width * 0.4
        .ColWidth(2) = .Width * 0.55
        .Clear
        .Cols = 3
        .Rows = 0
    End With
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Public Sub FlexGridLoad(flxgrd As MSFlexGrid, strSql As String)

' Assumes:  modConstants
On Error GoTo myErrHandler
    blnGrdLoad = True
    Dim strEntry As String
    Set mGrdRS = gdb.OpenRecordset(strSql, dbOpenDynaset)     'open the recordset
    
    With mGrdRS
        strEntry = "ID" & Chr(9) & "Web Site Title" & Chr(9) & "URL"
        flxgrd.AddItem strEntry
        .MoveFirst

        Do Until .EOF
            strEntry = !lngWebID & Chr(9) & !strSiteName & Chr(9) & !strURL & Chr(9)
            flxgrd.AddItem strEntry
            flxgrd.Row = .AbsolutePosition
            .MoveNext
        Loop
        .MoveFirst
        With flxgrd
            .FixedRows = 1
            .BackColorFixed = &H8000000A
            .BackColor = vbWhite
            .Enabled = True
        End With
        blnGrdLoad = False
        Call StatusMsgDisplay(TipCount(mGrdRS.RecordCount), 2)
    End With
    Exit Sub
myErrHandler:
    Select Case Err.Number
        Case 3021   ' there are no current records
            txtWeb(0) = ""
            InputInfoBox ("There are no current records.")
            txtWeb(0).SetFocus
            Exit Sub
        Case 30016
            Resume Next
        Case Else
            ErrMsgBox (Err.Description & " " & Err.Number)
            Resume Next
    End Select
End Sub

Private Sub RecordAdd()
On Error GoTo myErrHandler
    With mGrdRS
        .AddNew
        !strSiteName = txtWeb(0).Text
        !strURL = txtWeb(1).Text
        .Update
    End With
    mGrdRS.Close                         ' close the recordset
    GridRoutine
    Exit Sub
myErrHandler:
    Select Case Err.Number
        Case 3022
            ErrMsgBox ("The changes you requested to the table were not successful because that URL is already in the database.  Change the data in the URL field and try again.")
            txtWeb(1).SetFocus
            txtWeb(1).SelStart = 0
            txtWeb(1).SelLength = txtWeb(1).MaxLength
        Case Else
            ErrMsgBox (Err.Description & " " & Err.Number)
            Resume Next
    End Select
End Sub
Private Sub RecordUpdate()

On Error GoTo myErrHandler
    
    With mGrdRS
        .Edit
        !strSiteName = txtWeb(0).Text
        !strURL = txtWeb(1).Text
        .Update
    End With
    
    mGrdRS.Close                         ' close the recordset
    Call GridRoutine
    Exit Sub
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RecordDelete()
On Error GoTo myErrHandler

    With mGrdRS
        .Edit
        .Delete
    End With
    mGrdRS.Close                         ' close the recordset
    Call GridRoutine
    Exit Sub
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Function textValidate() As Boolean
On Error GoTo myErrHandler
    If Len(txtWeb(0)) Then
        If Len(txtWeb(1)) > 15 Then
            If Left(txtWeb(1), 11) = "http://www." Then
                textValidate = True
            ElseIf Left(txtWeb(1), 4) = "www." Then
                txtWeb(1).Text = "http://" & txtWeb(1).Text
                txtWeb(1).SetFocus
                txtWeb(1).SelStart = 11
                txtWeb(1).SelLength = txtWeb(1).MaxLength
            End If
        End If
    End If
    Exit Function                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Function

Private Sub GridRoutine()
On Error GoTo myErrHandler

    Call GridMaint(MSFlexGrid1)                 ' initialize the grid
    Call FlexGridLoad(MSFlexGrid1, qryWebLoad)  ' load it

    Exit Sub                                    ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub txtWeb_Change(Index As Integer)
On Error GoTo myErrHandler
'*****************************************************
' Purpose:  Tests to see if the input has actually changed.
' Inputs:   the Index of the current control    ' Returns:  None
' Assumes:  inclusion of modControls
'*****************************************************
    If txtWeb(0).Tag <> txtWeb(0).Text And cmdGrid(0).Tag <> "Add" Then
        cmdGrid(0).Caption = "Cancel"
        cmdGrid(0).Tag = "Edit"
        cmdGrid(1).Caption = "Update"
        MSFlexGrid1.TabStop = False
    End If
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub txtWeb_GotFocus(Index As Integer)
'*****************************************************
' Purpose:  Call the label and textbox onfocus routines
' Inputs:   the Index of the current control    ' Returns:  None
' Assumes:  inclusion of modControls
'*****************************************************
On Error GoTo myErrHandler
    txtWeb(Index).Tag = txtWeb(Index).Text
    Call ControlGotFocus(lblSubType(Index))
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub txtWeb_LostFocus(Index As Integer)
'*****************************************************
' Purpose:  Call the label and textbox lostfocus routines
' Inputs:   the Index of the current control    ' Returns:  None
' Assumes:  inclusion of modControls
'*****************************************************
On Error GoTo myErrHandler
    If txtWeb(Index).Text <> txtWeb(Index).Tag And cmdGrid(0).Tag <> "Add" Then
        txtWeb(Index).Tag = txtWeb(Index).Text
        cmdGrid(0).Tag = "Edit"
    End If
    Call ControlLostFocus(lblSubType(Index))
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
