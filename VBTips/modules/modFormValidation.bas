Attribute VB_Name = "modFormValidation"
Public Function TitleValidate(strTitle As String, Optional lblWarn As Label) As String
'*****************************************************
' Purpose:  Test for illegal characters (',") and spaces in the title'
' Inputs:  the string being checked, the warning label
' Returns: a new name for txtField.text if the current one contains illegal characters
'*****************************************************
Dim i As Integer, myLen As Integer
Dim bRetest As Boolean
Dim strTest As String * 1, strNewTitle As String, strOldTitle As String
    
    strOldTitle = strTitle                      ' stuff the input into a temp var
    myLen = Len(strTitle)                       ' find the length of the string
    
     Do While myLen > 0
        strTest = Left$(strOldTitle, 1)         ' take the left most character
        If strTest = "'" Or strTest = "''" Then ' test the character to see if is illegal
            lblWarn.Visible = True              ' warn against long names
            lblWarn.Caption = "Your Title contained illegal characters which have been removed."
            bRetest = True                      ' set flag for retest
        ElseIf strTest = " " Then               ' test for a proper use of space
            If i <> myLen Then                  ' next
                strNewTitle = strNewTitle & strTest 'concatinate the string
                i = myLen - 1                   ' decrement
            Else
                i = myLen - 1                   ' this takes care of multiple spaces
            End If                              ' closer inner if statement
        Else                                    ' legal character & spacing
            strNewTitle = strNewTitle & strTest ' if it is legal add to the string.
        End If                                  ' close outer if
        myLen = myLen - 1                       ' decriment the string by one character.
        strOldTitle = Right$(strOldTitle, myLen) ' redimension the string
    Loop
    If bRetest = True Then
        bRetest = False
        TitleValidate = Trim$(strNewTitle)          ' trim and return the new value
        Call TitleValidate(TitleValidate, lblWarn)  ' call recursively to double check the new name is not also taken.
    End If
    TitleValidate = Trim$(strNewTitle)              ' trim and return the new value
    Exit Function
myErrorHandler:
    ErrMsgBox ("Error has occured during Private sub 'TitleValidate' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
    Resume Next
End Function

Public Function TestTitle(myTipTitle As String) As String
'*****************************************************
' Purpose:  The purpose of this Sub is to test to make sure that I have an
'           original title. If it is not, it increments the title and retests to make sure
'           I don't have a title out of sequence
' Inputs:   txtFields(0).text
' Assumes:  modConstants
' Returns:  a new name for txtFields(0).text if the current one is taken.
'*****************************************************
    Dim oRS As Recordset
    Dim strSql As String, strRipQuote As String
    Dim i As Integer
        
    strRipQuote = Left$(myTipTitle, Len(myTipTitle) - 1) ' pull the right most quote off of the string to add #
 '   strSql = "SELECT tblTips.strTitle FROM tblTips WHERE (((tblTips.strTitle) Like " & myTipTitle & " Or (tblTips.strTitle) Like " & strRipQuote & " #" & """  Or (tblTips.strTitle) Like " & strRipQuote & " ##""" & "));"
    Set oRS = TitleTestRS(myTipTitle, strRipQuote)
    Do Until oRS.EOF    ' loop through all possible titles
        i = i + 1       ' increment the counter
        oRS.MoveNext    ' move to the next record
    Loop                ' do the loopdy loop
    oRS.Close           ' close it
    If i > 0 Then       ' test for the number of possible matches (there can be one at the most)
        myTipTitle = strRipQuote & " " & i  ' adjust the title
        myTipTitle = Right(myTipTitle, Len(myTipTitle) - 1)
           ' just in case the assigned title is taken
        TestTitle = TestTitle(TitleReTitle(myTipTitle)) ' test the new title recursively
    Else
        myTipTitle = Right(myTipTitle, Len(myTipTitle) - 1) ' pull of extra (")
        myTipTitle = Left(myTipTitle, Len(myTipTitle) - 1)  ' pull of extra (")
        TestTitle = myTipTitle
    End If                  ' end if
    Exit Function           ' end function
myErrorHandler:
    ErrMsgBox ("Error has occured during Private sub 'TestTitle' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
    Resume Next
End Function
Public Function TitleReTitle(strTitle As String) As String
'*****************************************************
' Purpose:  This function prepares the string for testing the retitled Tip's Title.
' Inputs:   the Title of the Tip     Returns:  The title bundled in quotes ("")
'*****************************************************
Dim myTipTitle As String

    myTipTitle = """" ' I need to pass an actual string, not just the value contained within
    myTipTitle = myTipTitle & Left$(strTitle, Len(strTitle))
    TitleReTitle = myTipTitle & """"
   
End Function

Public Sub FormClear(txtBox As Object)
'*****************************************************
' Purpose:  clear the form when adding a new record.
'           this will trigger the txtFields_change event.
' Inputs:   None    ' Returns:  None
'*****************************************************
   
    Dim i As Integer
    Do While i < txtBox.Count
        txtBox(i) = "" ' wipe out text
        i = i + 1
    Loop
    Exit Sub
myErrorHandler:
    ErrMsgBox ("Error has occured during Private sub 'FormClear' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
    Resume Next
End Sub

Public Function TxtLenWarn(txtBox As TextBox, intLen As Integer) As String
'*****************************************************
' Purpose:  Warn that a textbox may be too many characters, but still within the allowable length
' Inputs:   the textbox and it's length  Returns:  a string which may be put into a label or status bar, whatever
'*****************************************************
On Error GoTo myErrorHandler
    Dim itemp As Integer
    itemp = Len(txtBox) - intLen
    If itemp > 0 Then
        TxtLenWarn = "It is best to use less than " & intLen & " characters. You're over by: " & itemp
    Else                                ' there are not too many characters
        TxtLenWarn = ""                 ' hide the warning
    End If                              ' end if statement
    Exit Function                       ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Function

Public Function FormCheck(frm As Form) As Boolean
'*****************************************************
' Purpose:  Check the form at the form level before saving
' Inputs:   none
' Returns:  a value of "true" if their are required fields left blank
'*****************************************************
    If Len(frm.txtFields(0)) = 0 Then   ' The tip title is missing
        InputErrBox ("The Title Must Not be Left Blank")
        frm.txtFields(0).SetFocus
        FormCheck = True    ' set the flag
    ElseIf Len(frm.cmbTipType) = 0 Then   ' the tip type is missing
        InputErrBox ("The Tip Type Must Not be Left Blank")
        frm.cmbTipType.SetFocus
        FormCheck = True    ' set the flag
    ElseIf Len(frm.cmbTipSubType.Text) = 0 Then  ' the tip source is missing
        InputErrBox ("The Tip Source Should Not be Left Blank")
        frm.cmbTipSubType.SetFocus
        FormCheck = True    ' set the flag
    ElseIf Len(frm.txtFields(3)) = 0 And Len(frm.txtFields(4)) = 0 Then ' the tip type is missing
        InputErrBox ("You Have Not Yet Added a Tip.")
        frm.txtFields(3).SetFocus
        FormCheck = True    ' set the flag
    End If
End Function
