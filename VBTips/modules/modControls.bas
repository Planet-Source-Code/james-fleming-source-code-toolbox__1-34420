Attribute VB_Name = "modControls"
Option Explicit

Public Sub ComboboxLoad(cmb As ComboBox, strSql As String)
'*****************************************************
' Purpose:  loads the combo box with the titles
' Inputs:   The combo being filled, the sql statement being used in the recordset
' Assumes:  modError
' Returns:  None
' comment:  As the fields are not named explicitly, these row in the table may not be moved.
'*****************************************************
On Error GoTo myErrorHandler
    Dim oRS As Recordset
    cmb.Clear
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)    'open the recordset

    Do Until oRS.EOF        'loop to add the titles to the listbox
        With cmb                                    ' with statement
            .AddItem oRS.Fields(1) & ""             ' add the tip Title
            .ItemData(.NewIndex) = oRS.Fields(0)    ' add the tip index
            .ListIndex = oRS.Fields(0)
        End With                                    ' end
        oRS.MoveNext                                ' go to next record
    Loop                                            ' do the loop
    oRS.Close                                       ' close the recordset
    Exit Sub

myErrorHandler:
    Select Case Err.Number
        Case 380 ' Invalid property value (one of the properties of cmb isn't being used)
            Resume Next
        Case Else
        ErrMsgBox (Err.Description & " " & Err.Number)
        Resume Next
    End Select
End Sub

Public Sub ComboboxDualLoad(cmb1 As Control, cmb2 As Control, strSql As String)
'*****************************************************
' Purpose:  loads the combo box with the titles
' Inputs:   The combo being filled, the sql statement being used in the recordset
' Assumes:  modError
' Returns:  None
' comment:  As the fields are not named explicitly, these row in the table may not be moved.
'*****************************************************
On Error GoTo myErrorHandler
    
    Dim oRS As Recordset
    cmb1.Clear
    cmb2.Clear
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)    'open the recordset

    Do Until oRS.EOF        'loop to add the titles to the listbox
        With cmb1                                    ' with statement
            .AddItem oRS.Fields(1) & ""             ' add the tip Title
            .ItemData(.NewIndex) = oRS.Fields(0)    ' add the tip index
            .ListIndex = oRS.Fields(0)
        End With                                    ' end
        With cmb2                                   ' 2nd with statement
            .AddItem oRS.Fields(1) & ""             ' add the tip Title
            .ItemData(.NewIndex) = oRS.Fields(0)    ' add the tip index
            .ListIndex = oRS.Fields(0)
        End With                                    ' end
        oRS.MoveNext                                ' go to next record
    Loop                                            ' do the loop
    oRS.Close                                       ' close the recordset
    Exit Sub

myErrorHandler:
    Select Case Err.Number
        Case 380 ' Invalid property value (one of the properties of cmb isn't being used)
            Resume Next
        Case Else
        ErrMsgBox (Err.Description & " " & Err.Number)
        Resume Next
    End Select
End Sub
Public Sub SubTypeLoad(strSql As String, cmbSubType As Control)
'*****************************************************
' Purpose:  loads the SubType ComboBox based on the values from the TipType Combobox.
' Inputs:   the index of the combobox with the master values, the combo control receiving the return values
' Assumes:  modError
' Returns:  None
' comment:  Unlike ComboboxLoad above which I wanted to be more generic, this routine
'           doesn't pass in the sql statement, just the index of the control.
'*****************************************************
On Error GoTo myErrorHandler

    Dim oRS As Recordset                ' Open up a record set
    Dim strCaption As String            ' dim strings
    
    strCaption = cmbSubType.Text        ' assign the cmb text to temp var because the list box is about to be cleared
    
    cmbSubType.Clear                    ' clear the combo box
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)    'open the recordset

    Do Until oRS.EOF        'loop to add the titles to the listbox
        With cmbSubType     ' set combo properties
            .AddItem oRS!strSTTitle & ""
            .ItemData(.NewIndex) = oRS!lngSubTypeID
        End With
        oRS.MoveNext                    ' advance the recordset
    Loop                                ' next
    oRS.Close                           ' close the recordset
    cmbSubType.Text = strCaption        ' reset the combo text
    Exit Sub                            ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & ". In Function SubTypeLoad of modControls.")
    Resume Next
End Sub
Public Sub DualSubTypeLoad(strSql As String, cmbSubType1 As Control, cmbSubType2 As Control)
'*****************************************************
' Purpose:  loads the SubType ComboBox based on the values from the TipType Combobox.
' Inputs:   the index of the combobox with the master values, the combo control receiving the return values
' Assumes:  modError
' Returns:  None
' comment:  Unlike ComboboxLoad above which I wanted to be more generic, this routine
'           doesn't pass in the sql statement, just the index of the control.
'*****************************************************
On Error GoTo myErrorHandler

    Dim oRS As Recordset                ' Open up a record set
    Dim strCaption1 As String, strCaption2 As String            ' dim strings
    strCaption1 = cmbSubType1.Text        ' assign the cmb text to temp var because the list box is about to be cleared
    strCaption2 = cmbSubType2.Text        ' assign the cmb text to temp var because the list box is about to be cleared
    cmbSubType1.Clear                    ' clear the combo box
    cmbSubType2.Clear                    ' clear the combo box
    
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)    'open the recordset

    Do Until oRS.EOF                    'loop to add the titles to the listbox
        With cmbSubType1                ' set combo properties
            .AddItem oRS!strSTTitle & ""
            .ItemData(.NewIndex) = oRS!lngSubTypeID
        End With
        With cmbSubType2                ' set combo properties
            .AddItem oRS!strSTTitle & ""
            .ItemData(.NewIndex) = oRS!lngSubTypeID
        End With
        oRS.MoveNext                    ' advance the recordset
    Loop                                ' next
    oRS.Close                           ' close the recordset
    cmbSubType1.Text = strCaption1       ' reset the combo text
    cmbSubType2.Text = strCaption2       ' reset the combo text
    Exit Sub                            ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & ". In Function DualSubTypeLoad of modControls.")
    Resume Next
End Sub

Public Sub ControlGotFocus(myCaption As Control)
'*****************************************************
' Purpose:  The purpose of this Sub is to Highlight the label or other
'           control that is passed in
' Inputs:   You must pass in the control being 'lit up'
' Returns:  nothing but an illuminated control forecolor!
' Comments: I like to use this feature to light up labels to better show focus
'*****************************************************
Dim lngHighlight As Long
    lngHighlight = &HFFFF&       ' changes the constant to yellow
    myCaption.ForeColor = lngHighlight  ' light me up
End Sub

Public Sub ControlLostFocus(myCaption As Control)
'*****************************************************
' Purpose:  The purpose of this Sub is to return to the default mode
'           the label or other control that is passed in
' Inputs:   You must pass in the control that is 'lit up'
' Returns:  none
'*****************************************************
    myCaption.ForeColor = vbDefault ' reset the default color
End Sub

Public Function TextboxSelect(myControl As TextBox)
'*****************************************************
' Purpose:  Highlight any text in the text box to make
'           editing easier
' Inputs:   the text control being selected
' Returns: none
' Comments: This is to function just like in Access.
'           You must set the max length property in each control
'***************************************************
    myControl.SelStart = 0                      ' the start position
    myControl.SelLength = myControl.MaxLength   ' the end position
End Function

Public Function ListPopulate(strSql As String, myListbox As Control) As Integer
'*****************************************************
' Purpose:  this sub populates the list box based on a
'           record set connection to the database
' Inputs:   the sql statement, the listbox being filled
' Assumes: modError
' Returns:  The record count
'*****************************************************
On Error GoTo myErrorHandler
    Dim i As Integer
    Dim oRS As Recordset    'open the recordset
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)

    myListbox.Clear
    Do Until oRS.EOF                'loop to add the titles to the listbox
        With myListbox
            .AddItem oRS!strTitle & ""
            .ItemData(.NewIndex) = oRS!lngTblTipsID ' note the difference from sublist populate
        End With
        i = i + 1                   ' increment my sublist index
        oRS.MoveNext                ' advance the recordset
    Loop                            ' continue

    ListPopulate = oRS.RecordCount  ' count the tips and return that number
    If i <> ListPopulate Then
        MsgBox "James, there are situations where ListPopulates oRS.RecordCount <> the # of records. You must replace listPop in all functions that call the sublistPop."
    End If
    oRS.Close                       ' close the record set
    Exit Function
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & ". In Function ListPopulate of modControls.")
    Resume Next
End Function
Public Function ListNavigate(cmdIndex As Integer, lstBox As ListBox, myRS As Recordset, bSubList As Boolean)

'*****************************************************
' Purpose: Navigate through the records of the list
' Comment: We have to test to make sure we're not on the first or last record
'          We have to handle moving through the sublists differently than
'          the entire collection.
' Comment: ALso if we are looking at a sublist of an active recordset, then the problems are handled by a module level flag.
'   For example in a sub list the indexes can get all screwed up. That is handled by setting the bSubList=True.
' Inputs:  the index of the control calling the event: 0=First; 1=Previous 2=Next; 3=Last
'          the listbox being navigated, the recordset being traversed, a boolean flag if the list being navigated is a subset of the recordset.
' Returns: None
' Assume:  All the nav buttons are an array. There is a module level recordset that is open to the database with a live connection
'*****************************************************
 On Error GoTo cmdNavErr
 Dim miListCount As Integer    ' for creating a temp index when a partial list of titles is made
    miListCount = lstBox.ListCount
    
    Select Case cmdIndex
        Case 0                              ' move to the first record in the recordset
           myRS.MoveFirst
           lstBox.ListIndex = myRS.AbsolutePosition 'triggers the lstBox_click event
        Case 1                              ' Move to the previous code sample (Back)
           myRS.MovePrevious                ' back up one record
           If Not myRS.BOF Then             ' it is safe to proceed
               If bSubList = False Then     ' we have the entire list
                   lstBox.ListIndex = myRS.AbsolutePosition 'triggers the lstBox_click event
               Else                        ' we have the sublist
                   If lstBox.ListIndex > 0 Then ' we're not at the beginning of the file
                       lstBox.ListIndex = lstBox.ListIndex - 1 ' move back 1
                   End If
               End If
           Else    ' we are at the beginning
               Call ListNavigate(0, lstBox, myRS, bSubList)      ' call the routine that handles the first record
           End If
        Case 2                             ' Move to the next code sample (Next)
           myRS.MoveNext                   ' go to the next record in the recordset
           If Not myRS.EOF Then            ' It's safe to move forward
               If bSubList = False Then    ' it's the entire list
                   lstBox.ListIndex = myRS.AbsolutePosition 'triggers the lstBox_click event
               Else
                    lstBox.ListIndex = lstBox.ListIndex + 1 ' move forward
               End If
           Else                             ' we're at the end
               Call ListNavigate(3, lstBox, myRS, bSubList)        ' call the routine that handles the last record
           End If
        Case 3 ' Purpose:  Navigate to the end.
            myRS.MoveLast   ' go to the end of the line
            If bSubList = False Then    ' it must be the whole enchelada
               lstBox.ListIndex = myRS.AbsolutePosition 'triggers the lstBox_click event
            Else
               lstBox.ListIndex = miListCount - 1   ' it is the sub list
            End If
         Case Else
            ErrMsgBox ("It seems we have more buttons than required for naviation of the list box. The error is in ListNavigate")
    End Select
    
    Exit Function
cmdNavErr:
    Beep
    On Error Resume Next
End Function

Public Sub qrySearch(cmdSearch As CommandButton, txtSearch As TextBox, lblSearch As Label, lstTitle As ListBox)
'*****************************************************
' Purpose:  Create the SQL string based on the content of
'           the txtSearch.text. Passes the string to the appropriate sub
' Inputs:   None  ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
    Dim qryString As String
    If cmdSearch.Caption = "Find" Then          ' we wish to search for a tip
        qryString = """*"                       ' add wild card to the front
        qryString = qryString & txtSearch.Text  ' add the criteria
        qryString = qryString & "*"""           ' add wild card to the back
        ' here's where we decide to search the title or the whole thing
        If lblSearch.Caption = "Find &Quick" Then    'test for search type
            ' Display Status of Tip Count,populate list w/Title Search on string
            Call StatusMsgDisplay(TipCount(ListPopulate(SearchTitle(qryString), lstTitle)), 2)
        Else
            ' Display Status of Tip Count,populate list w/Title Search on string (Title, & Memo fields)
            Call StatusMsgDisplay(TipCount(ListPopulate(SearchAll(qryString), lstTitle)), 2)    ' creates a sublist of an access datatable
        End If                          '
        cmdSearch.Caption = "Clear"     ' set the button to toggle
    Else                                ' must have been set Clear
        txtSearch.Text = ""             ' empty text box
        txtSearch.SetFocus              ' search text box
        cmdSearch.Caption = "Find"      ' reset the button
    End If                           ' end if statement
    Exit Sub                         ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
