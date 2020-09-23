Attribute VB_Name = "modFormUtilities"
Option Explicit

Public Sub FormCenter(frmAny As Form)
'*******************************************************
' Purpose:  For centering a from within the MDI Parent Screen
' Effects:  Centers form within the screen area
' Inputs:   The name of the form being centered.    Returns: None
'*******************************************************

    frmAny.Top = (Screen.Height / 2) - (frmAny.Height / 2)
    frmAny.Left = (Screen.Width / 2) - (frmAny.Width / 2)
End Sub

Public Sub UnloadAllForms()
'*****************************************************
' Purpose:  This is the sub that unloads all forms
' Inputs:   None ' Returns:  None
' Comments: This is used to make certain all forms are released from memory
'*****************************************************
    Dim frm As Form
    On Error Resume Next
    For Each frm In Forms ' using the forms collection
        Unload frm
    Next frm
End Sub

Public Sub UnloadChildForms()
'*****************************************************
' Purpose: Close all child forms
' Assumes:  modConstants
' Inputs:   None ' Returns:  None
'*****************************************************

' Turn on error handling

    On Error Resume Next
    ' Reset the Err object
    Dim i As Integer
    Err = 0
    ' Close all the child forms
    Do Until Err
        Unload fMDI.ActiveForm
    Loop

End Sub
Public Sub FormCaption(frm As Form, strCaption As String)
'*****************************************************
' Purpose: puts the current tip info into the form caption
' Inputs:   None    ' Returns:  None
'*****************************************************
    frm.Caption = strCaption
End Sub
