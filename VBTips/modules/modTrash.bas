Attribute VB_Name = "modTrash"

' Here is where I stick  the stuff I no longer believe I need.
Public Function SubListPopulate(strSql As String, myListbox As Control) As Integer
'*****************************************************
' Purpose: creates a sublist of data based on a search
' Assumes: modConstants
' Inputs:  the sql statement
' Returns: an integer value for the sublist.
' Comment: The return value is required to keep moving through the record set working properly.
' Insert the double commented out ('') code below into the form that has myListbox (remember to remove the '')
''    Dim miListCount As Integer    ' for creating a temp index when a partial list of titles is made. (Put in general declarations)
''    Dim bSubList As Boolean     ' flag for if a partial list of tips if being viewed
''
''    If bSubList = False Then    ' it must be the whole enchelada
''       lstTitle.ListIndex = moRS.AbsolutePosition 'triggers the lstTitle_click event
''    Else
''       lstTitle.ListIndex = miListCount - 1 ' it is the sublist and
''    End If

'*************************************************
    MsgBox "I thought you didn't need this"
    
    On Error GoTo SubListPopErr
    Dim oRS As Recordset        'open the recordset
    Dim iCount As Integer ' dimension a new index for a sublist
    
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot) ' open the recordset
    myListbox.Clear           ' clear the list box
    'loop to add the titles to the listbox
    Do Until oRS.EOF
        With myListbox
            .AddItem oRS!strTitle & ""
            .ItemData(.NewIndex) = oRS!lngTblTipsID
        End With
        oRS.MoveNext
        iCount = iCount + 1         ' increment my sublist index
        
    Loop
    iCount = oRS.RecordCount
  '  Call StatusMsgDisplay(TipCount(oRS.RecordCount), 2) ' count the tips and display the result
    oRS.Close                       ' close the record set

    SubListPopulate = iCount        ' set the value of the returning set.
Exit Function
SubListPopErr:
Select Case Err
    Case Else
    ErrMsgBox (Err.Description & " " & Err.Number & ". In Function SubListPopulate of modControls.")
    Resume Next
End Select
End Function


'Private Sub mnuHelpContents_Click()
'    On Error GoTo myErrorHandler
'
'    Dim nRet As Integer
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
'    Else
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
'    End If
'    Exit Sub
'myErrorHandler:
'    ErrMsgBox (Err.Description & " " & Err.Number)
'    Resume Next
'End Sub

'Private Sub mnuHelpSearch_Click()
'
'On Error GoTo myErrorHandler
'    Dim nRet As Integer
'    'if there is no helpfile for this project display a message to the user
'    'you can set the HelpFile for your application in the
'    'Project Properties dialog
'    If Len(App.HelpFile) = 0 Then
'        ErrMsgBox ("Unable to display Help Contents. There is no Help associated with this project.")
'    Else
'        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
'    End If
'    Exit Sub
'myErrorHandler:
'    ErrMsgBox (Err.Description & " " & Err.Number)
'    Resume Next
'End Sub
'Private Sub cmdNav_Click(index As Integer)
''*****************************************************
'' Purpose: Navigate through the records of the list
'' Comment:  We have to test to make sure we're not on the first or last record
''           We have to handle moving through the sublists differently than
''           the entire collection
'' Inputs:   the index of the control.   Returns:  None
'' assume:   All the nav buttons are an array.
''*****************************************************
' On Error GoTo cmdNavErr
' Dim miListCount As Integer    ' for creating a temp index when a partial list of titles is made
'    miListCount = lstTitle.ListCount
'    Select Case index
'        Case 0     ' move to the first record in the recordset
'           moRS.MoveFirst
'           lstTitle.ListIndex = moRS.AbsolutePosition 'triggers the lstTitle_click event
'        Case 1     ' Move to the previous code sample
'           moRS.MovePrevious               ' back up one record
'           If Not moRS.BOF Then            ' it is safe to proceed
'               If bSubList = False Then    ' we have the entire list
'                   lstTitle.ListIndex = moRS.AbsolutePosition 'triggers the lstTitle_click event
'               Else                        ' we have the sublist
'                   If lstTitle.ListIndex > 0 Then ' we're not at the beginning of the file
'                       lstTitle.ListIndex = lstTitle.ListIndex - 1 ' move back 1
'                   End If
'               End If
'           Else    ' we are at the beginning
'               Call cmdNav_Click(0) ' call the routine that handles the first record
'           End If
'        Case 2     ' Move to the next code sample
'           moRS.MoveNext                   ' go to the next record in the recordset
'           If Not moRS.EOF Then            ' It's safe to move forward
'               If bSubList = False Then    ' it's the entire list
'                   lstTitle.ListIndex = moRS.AbsolutePosition 'triggers the lstTitle_click event
'               Else
'                    lstTitle.ListIndex = lstTitle.ListIndex + 1 ' move forward
'               End If
'           Else    ' we're at the end
'               Call cmdNav_Click(3) ' call the routine that handles the last record
'           End If
'        Case 3 ' Purpose:  Navigate to the end.
'            moRS.MoveLast   ' go to the end of the line
'            If bSubList = False Then    ' it must be the whole enchelada
'               lstTitle.ListIndex = moRS.AbsolutePosition 'triggers the lstTitle_click event
'            Else
'               lstTitle.ListIndex = miListCount - 1   ' it is the sub list
'            End If
'         Case Else
'            lstTitle.ListIndex = index ' it is the sub list
'    End Select
'    If lblTitleWarn.Caption <> "" Then lblTitleWarn.Caption = "" ' hide the title warning
'    Exit Sub
'cmdNavErr:
'    Beep
'    On Error Resume Next
'End Sub

'Public Sub ComboboxLoad(cmb As Control, strSql As String)
''*****************************************************
'' Purpose:  loads the combo box with the titles
'' Inputs:   The combo being filled, the sql statement being used in the recordset
'' Assumes:  modError
'' Returns:  None
''*****************************************************
'On Error GoTo myErrorHandler
'
'    Dim oRS As Recordset
'    cmb.Clear
'    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)    'open the recordset
'
'    Do Until oRS.EOF        'loop to add the titles to the listbox
'        With cmb                                    ' with statement
'            .AddItem oRS!strTitle & ""            ' add the tip Title
'            .ItemData(.NewIndex) = oRS!intLangID    ' add the tip index
'            .ListIndex = oRS!intLangID
'        End With                                    ' end
'        oRS.MoveNext                                ' go to next record
'    Loop                                            ' do the loop
'    oRS.Close                                       ' close the recordset
'    Exit Sub
'myErrorHandler:
'    ErrMsgBox (Err.Description & " " & Err.Number)
'    Resume Next
'End Sub

''Private Sub TipSourceLoad() Commented out 2/12/00
'''*****************************************************
''' Purpose:  loads the combo box with the titles of the tip source
''' Assumes:  modConstants
''' Inputs:   None    ' Returns:  None
'''*****************************************************
''On Error GoTo myErrorHandler
''
''    Dim oRS As Recordset
''    Dim strSQL As String
''
''    strSQL = "SELECT tblAuthor.lngAuthorID, tblAuthor.strAuthor FROM tblAuthor ORDER BY tblAuthor.strAuthor;"
''
''    Set oRS = gdb.OpenRecordset(strSQL, dbOpenSnapshot)    'open the recordset
''
''    Do Until oRS.EOF    'loop to add the titles to the listbox
''        With cmbTipSubType
''            .AddItem oRS!Author & ""
''            .ItemData(.NewIndex) = oRS!id
''        End With
''        oRS.MoveNext
''    Loop
''    oRS.Close
''    Exit Sub
''myErrorHandler:
''    ErrMsgBox ("Error has occured during Private sub 'TipSourceLoad' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
''    Resume Next
''End Sub

''
''Private Sub LoadByType(lstBox As ListBox) commented out 02/12/00
'''*****************************************************
''' Purpose:  loads the list box with the titles
''' Assumes: modConstants
''' Inputs:   None    ' Returns:  None
'''*****************************************************
''    Dim oRS As Recordset
''    Dim strSQL As String
''
''    strSQL = "SELECT tblTips.lngTblTipsID, tblTips.strTitle FROM tblTips ORDER BY tblTips.Type;"
''    Set oRS = gdb.OpenRecordset(strSQL, dbOpenSnapshot) 'open the recordset
''
''    Do Until oRS.EOF    'loop to add the titles to the listbox
''        With lstBox
''            .AddItem oRS!Title & ""
''            .ItemData(.NewIndex) = oRS!id
''        End With
''        oRS.MoveNext
''    Loop
''
''    oRS.Close   ' close the record set
''
''    Exit Sub
''myErrorHandler:
''    ErrMsgBox ("Error has occured during Private sub LoadByType " & Err.Description & " " & Err.Number)
''    Resume Next
''End Sub

'
'Private Sub cmdFind_Click()
''*****************************************************
'' Purpose:  Create the SQL string based on the content of
''           the txtFields(2).text. Passes the string to the appropriate sub
'' Inputs:   None  ' Returns:  None
''*****************************************************
'On Error GoTo myErrorHandler
'
'    Call qrySearch(cmdFind, txtFields(2), lblDetail(2))
''    Dim qryString As String
''    If cmdFind.Caption = "Find" Then    ' we wish to search for a tip
''        If txtFields(2).Text = "" Then  ' test for a value. If blank reset the listbox
''            Call LoadAllRecords         ' reset the tips to start up status.
''            txtFields(2).SetFocus       ' search text box
''            Exit Sub                    ' exit the routine
''        End If                          ' close inner if statement
''        qryString = """*"                ' add wild card to the front
''        qryString = qryString & Me.txtFields(2).Text ' add the criteria
''        qryString = qryString & "*"""     ' add wild card to the back
''
''        ' here's where we decide to search the title or the whole thing
''        If lblDetail(2).Caption = "Find &Quick" Then    'test for search type
''             ' Display Status of Tip Count,populate list w/Title Search on string
''            Call StatusMsgDisplay(TipCount(ListPopulate(SearchTitle(qryString), lstTitle)), 2)
''        Else    ' Display Status of Tip Count,populate list w/Title Search on string (Title, & Memo fields)
''             Call StatusMsgDisplay(TipCount(ListPopulate(SearchAll(qryString), lstTitle)), 2)    ' creates a sublist of an access datatable
''        End If
''        cmdFind.Caption = "Clear"       ' set the button to toggle
''    Else                                ' must have been set Clear
''        txtFields(2).Text = ""          ' empty text box
''        txtFields(2).SetFocus           ' search text box
''        cmdFind.Caption = "Find"        ' reset the button
''    End If                              ' end if statement
'    FirstRecordShow                     ' populate the list
'    Exit Sub                            ' exit the routine
'myErrorHandler:
'    ErrMsgBox (Err.Description & " " & Err.Number)
'    Resume Next
'End Sub
'
'Public Sub SizeFormByPixels(frm As Form, nWidth As Integer, nHeight As Integer)
'  ' receives an incoming width and height of an internal control
'  ' sets the form's size to just a few pixels larger than the internal control's size
'
'  Dim OffsetX As Integer
'  Dim OffsetY As Integer
'  Dim NewSizeX As Integer
'  Dim NewSizeY As Integer
'  Dim OldScaleMode As Integer
'
'  On Error Resume Next
'  With frm
'    OldScaleMode = .ScaleMode
'    .ScaleMode = vbPixels
'    OffsetX = .ScaleWidth - nWidth - 2
'    OffsetY = .ScaleHeight - nHeight - 2
'    NewSizeX = .Width - (Screen.TwipsPerPixelX * OffsetX)
'    NewSizeY = .Height - (Screen.TwipsPerPixelY * OffsetY)
'    .Move .Left, .Top, NewSizeX, NewSizeY
'    .ScaleMode = OldScaleMode
'  End With
'End Sub

