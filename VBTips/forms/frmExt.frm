VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Extension Maintenance"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5775
   Icon            =   "frmExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Delete"
      Height          =   315
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Help"
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrid 
      Caption         =   "Add"
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtSubType 
      Height          =   285
      Left            =   1320
      MaxLength       =   75
      TabIndex        =   1
      Top             =   495
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5175
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   9128
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
   Begin VB.ComboBox cmbTypeName 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Language"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblSubType 
      Caption         =   "Extension"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   495
      Width           =   735
   End
   Begin VB.Label lblSubType 
      Caption         =   "Type"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
' Purpose: this form is used to demonstrate MSFlexgrid
'           It is used for adding Tip subcategories to the mdb
' Assumes: inclusion of the flexgrid control
' Comment: This is fairly robust, but not flawless.
'          I worked on this to pass time one snowy weekend
' Author:   James R. Fleming    Date: 1/25/2000
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
    Dim bflag As Boolean
    blnFrmLoad = True
    Call ComboboxLoad(cmbTypeName, qryCombo)

    ' test to see if the form is loaded. If so load based on the active type
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = "frmSearch" Then ' use the form without calling a new instance of it
            bflag = True
            Call SelectStr(1, fSearch.cmbTipType.ListIndex)
            Exit For
         End If
    Next i
    
    If bflag = False Then Call SelectStr(0) ' if frmtblTips wasn't loaded then...
    blnFrmLoad = False                      ' reset remaining flags
    bflag = False
    cmdGrid(2).Enabled = False
    Call Form_DblClick
    
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub Form_Activate()
'*****************************************************
' Purpose:  this sub displays a message in the status bar then changes it after 4 secs.
' Assume:   The inclusion of StatusMsgFlip in modFormUtilities
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myActivateerr

    Call StatusMsgFlip("The subtype form is active", mGrdRS.RecordCount)
    Exit Sub
    
myActivateerr:
Select Case Err.Number
    Case 91
            Call StatusMsgFlip("The subtype form is active", 0)
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
If bflag = True Then Exit Sub
Select Case Me.WindowState
    Case vbMaximized
        MSFlexGrid1.Height = ScaleHeight - (MSFlexGrid1.Top + 50)
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
    Call StatusMsgDisplay("", 2)
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
        txtSubType.Tag = !strLang
        txtSubType = !strLang
    End With
    cmdGrid(2).Enabled = True
    MSFlexGrid1.TabStop = True
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub cmbTypeName_GotFocus()
'*****************************************************
' Purpose:  Call the label onfocus routines
' Assumes:  inclusion of modControls
' Inputs:   none   ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    cmbTypeName.Tag = cmbTypeName.Text
    Call ControlGotFocus(lblSubType(1))
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub cmbTypeName_Click()
'*****************************************************
' Purpose:  Call the SelectStr subroutine
' Inputs:   none   ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
If blnFrmLoad = True Then Exit Sub               ' click event gets fired on form_load
    If cmbTypeName.Tag <> cmbTypeName.Text Then  ' they actually did make a change and not just click
        cmbTypeName.Tag = cmbTypeName.Text       ' reset the flag
        txtSubType = ""                          ' clear the text boxes
        Call SelectStr(1, cmbTypeName.ListIndex) ' assign the proper sqlString
    End If
    Exit Sub                                     ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub cmbTypeName_LostFocus()
'*****************************************************
' Purpose:  Call the label and textbox lostfocus routines
' Inputs:  None    ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    Call ControlLostFocus(lblSubType(1))
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
' Comment:  Also relys on a flag set on cmdGrid(0).Caption & cmdGrid(0).tag
'*****************************************************
On Error GoTo myErrHandler
    Select Case Index
        Case 0                                  ' the Add/Cancel Button
            If cmdGrid(0).Caption = "Add" Then
                Call RowAdd                     ' visual cues of modification started
            Else                                ' the action was cancelled
                Call RowAddCancel               ' visual cues of modification canceled
                cmbTypeName.SetFocus            ' set focus back to the combo
            End If                              ' end if
        Case 1                                  ' the Edit/Update Button
            If cmdGrid(0).Tag = "Add" Then
                Call RowAddCommit               ' add a new record
            ElseIf cmdGrid(0).Tag = "Edit" Then
                Call RowUpdateCommit            ' update an existing record
            Else                                ' nothing to update yet!
                Call InputInfoBox("To add a file extension:" & vbCrLf & "Press the Add button." & vbCrLf & "Add the extension. Press update." & vbCrLf & vbLf & "To edit an extension:" & vbCrLf & "Select the row you wish to edit." & vbCrLf & "Modify the text. Press update." & vbCrLf & vbLf & "To delete an extension:" & vbCrLf & "Select the row you wish to delete." & vbCrLf & "Press Delete.")
            End If
        Case 2                                  ' the Delete Button
            Call RowDelete                  ' delete the record
            Call RowAddCancel               ' visual cues of modification canceled
            cmbTypeName.SetFocus            ' set focus back to the combo
    End Select
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub txtSubType_Change()
On Error GoTo myErrHandler
'*****************************************************
' Purpose:  Tests to see if the input has actually changed.
' Inputs:   the Index of the current control    ' Returns:  None
'*****************************************************
    If txtSubType.Tag <> txtSubType.Text And cmdGrid(0).Tag <> "Add" Then
        cmdGrid(0).Caption = "Cancel"
        cmdGrid(1).Caption = "Update"
        MSFlexGrid1.TabStop = False
    End If
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub txtSubType_GotFocus()
'*****************************************************
' Purpose:  Call the label and textbox onfocus routines
' Inputs:   the Index of the current control    ' Returns:  None
' Assumes:  inclusion of modControls
'*****************************************************
On Error GoTo myErrHandler
    txtSubType.Tag = txtSubType.Text
    Call ControlGotFocus(lblSubType(0))
    If Me.txtSubType.Tag <> "*." Then Call TextboxSelect(txtSubType)
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub txtSubType_LostFocus()
'*****************************************************
' Purpose:  Call the label and textbox lostfocus routines
' Inputs:   the Index of the current control    ' Returns:  None
' Assumes:  inclusion of modControls
'*****************************************************
On Error GoTo myErrHandler
    If txtSubType.Text <> txtSubType.Tag And cmdGrid(0).Tag <> "Add" Then
        txtSubType.Tag = txtSubType.Text
        cmdGrid(0).Tag = "Edit"
    End If
    Call ControlLostFocus(lblSubType(Index))
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
'*****************************************************
' Purpose:  Handle all of the visual feedback of adding
' Inputs:   None   ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    cmdGrid(0).Tag = "Add"                      ' set the cmd tags & captions
    cmdGrid(0).Caption = "Cancel"
    cmdGrid(1).Caption = "Update"
    cmdGrid(2).Enabled = False
    
    txtSubType.Tag = txtSubType.Text            ' set any current values temporarily
    cmdGrid(0).Tag = "Add"
    txtSubType.Text = "*."
    txtSubType.SetFocus                         ' set focus
    txtSubType.SelStart = 2
    MSFlexGrid1.Enabled = False                 ' toggle the grid
    MSFlexGrid1.BackColorSel = vbWhite
    MSFlexGrid1.ForeColorSel = &H80000011
    MSFlexGrid1.ForeColor = &H80000011
    MSFlexGrid1.ForeColorFixed = &H80000011

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
'*****************************************************
' Purpose:  Set the visual properties of the enabled grid
' Inputs:   none   ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    MSFlexGrid1.ForeColor = &H80000007      ' toggle the grid
    MSFlexGrid1.ForeColorFixed = &H80000012
    MSFlexGrid1.BackColorSel = &H8000000D
    MSFlexGrid1.ForeColorSel = &H8000000E
    MSFlexGrid1.TabStop = True
    MSFlexGrid1.Enabled = True
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowUpdateCommit()
'*****************************************************
' Purpose:  Handle all of the Update events
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    If textValidate = True Then         ' the input is good.
        Call RecordUpdate               ' call the update subroutine
        txtSubType = ""                 ' clear the text boxes
        cmdGrid(0).Tag = ""             ' set the cmd tags & captions
        cmdGrid(0).Caption = "Add"
        cmdGrid(1).Caption = "Help"
        cmdGrid(2).Enabled = False
        Call GrdTglEnabled              ' toggle the grid
    End If                              ' end if
    Exit Sub                            ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowAddCancel()
'*****************************************************
' Purpose:  Handle all of the Row Cancel events
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrHandler
    cmdGrid(0).Caption = "Add"
    cmdGrid(1).Caption = "Help"
    cmdGrid(2).Enabled = False
    cmdGrid(0).Tag = ""
    txtSubType.Tag = ""             ' clear the text boxes
    txtSubType = ""
    Call GrdTglEnabled              ' toggle the grid
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RowDelete()
On Error GoTo myErrHandler
    Dim intResponse As Integer
    intResponse = YesNo("Delete Current Tip")
    If intResponse = vbYes Then
        txtSubType = ""
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
'*****************************************************
' Purpose:  Handle all of the grid initializing
' Inputs:   the FlexGrid    ' Returns:  None
' Assumes:  inclusion of modControls
'*****************************************************
On Error GoTo myErrHandler
    With flxgrd
        .ColWidth(0) = flxgrd.Width * 0.1     ' set the col width
        .ColWidth(1) = flxgrd.Width * 0.87
        .Clear                    ' clear the grid
        .Cols = 2
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
    Dim strEntry As String, col0 As String, col1 As String
    Set mGrdRS = gdb.OpenRecordset(strSql, dbOpenDynaset)     'open the recordset
    
    With mGrdRS
        strEntry = "" & Chr(9) & "SubType Title"
        flxgrd.AddItem strEntry
        .MoveFirst
        Do Until .EOF
            col0 = ""
            col1 = !strLang
            strEntry = col0 & Chr(9) & col1 & Chr(9)
            flxgrd.AddItem strEntry
            flxgrd.Row = .AbsolutePosition
            .MoveNext
        Loop
        .MoveFirst
        Call StatusMsgDisplay(TipCount(mGrdRS.RecordCount), 2)
    End With
        flxgrd.FixedRows = 1
        flxgrd.BackColorFixed = &H8000000A
        flxgrd.BackColor = vbWhite
        blnGrdLoad = False
    Exit Sub
myErrHandler:
    Select Case Err.Number
        Case 3021   ' there are no current records
            txtSubType = ""
            InputInfoBox ("There are no current records.")
            txtSubType.SetFocus
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
        !lngTable_PK = cmbTypeName.ListIndex
        !strLang = txtSubType.Text
        .Update
    End With
 
    mGrdRS.Close                         ' close the recordset
    Call SelectStr(1, cmbTypeName.ListIndex)
    Exit Sub
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Sub RecordUpdate()

On Error GoTo myErrHandler
    
    With mGrdRS
        .Edit
        !lngTable_PK = cmbTypeName.ListIndex
        !strLang = txtSubType
        .Update
    End With
    
    mGrdRS.Close                         ' close the recordset
    Call SelectStr(2, cmbTypeName.ListIndex)
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
    Call SelectStr(2, cmbTypeName.ListIndex)
    Exit Sub
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Private Function textValidate() As Boolean
On Error GoTo myErrHandler
    If txtSubType = "" Or txtSubType = "*." Then
        InputErrBox ("The Extension Type appears to be incorrect. " & vbCrLf & "A valid file extension is required.")
        txtSubType.SetFocus
        txtSubType = "*."
        txtSubType.SelStart = 2
        textValidate = False
    ElseIf Left(txtSubType, 2) <> "*." Then InputErrBox ("The Extension Type appears to be incorrect." & vbCrLf & "A wildcard and period (*.) is required before the extension.")
        txtSubType.SetFocus
        txtSubType = "*." & txtSubType.Text
        txtSubType.SelStart = 2
        textValidate = False
    Else
        textValidate = True
    End If
    Exit Function                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next

End Function

Private Sub SelectStr(iCase As Integer, Optional Index As Integer)
On Error GoTo myErrHandler
Dim strSql As String

    Select Case iCase
        Case 0   ' show none.
            Call GridMaint(MSFlexGrid1)
            With mGrdRS
                entry = "   " & Chr(9) & "Extension Title"
                MSFlexGrid1.AddItem entry
            End With
            cmbTypeName.Text = "Select Tip Type"
            txtSubType = ""
            txtSubType.Enabled = False
            cmdGrid(0).Enabled = False
            MSFlexGrid1.Enabled = False
            Exit Sub
        Case 1  ' load tip based on ftblTip active type
            cmbTypeName.ListIndex = Index
            txtSubType.Enabled = True
            cmdGrid(0).Enabled = True
            MSFlexGrid1.Enabled = True
            strSql = ExtLoad(cmbTypeName.ListIndex)
            
        Case 2
            cmbTypeName.ListIndex = Index
            txtSubType.Enabled = True
            cmdGrid(0).Enabled = True
            MSFlexGrid1.Enabled = True
            strSql = ExtLoad(cmbTypeName.ListIndex)
        Case Else
            ErrMsgBox ("Something unforseen happened in frmExt selectStr")
    End Select
    Call GridMaint(MSFlexGrid1)
    Call FlexGridLoad(MSFlexGrid1, strSql)
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

