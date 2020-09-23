Attribute VB_Name = "modStatusBar"
Option Explicit

Public Sub StatusMsgDisplay(ByVal myString As String, iPanel As Integer)

'*******************************************************
' Purpose:  Put the count of tips into the status box
' Depends:  The status bar has as many panels as the highest integer
' Assumes:  modConstants
' Inputs:   the string being displayed  Returns: None
' Author:   James R. Fleming
'*******************************************************

    fMDI.sbStatusBar.Panels(iPanel).Text = myString
End Sub
Public Function TipCount(Optional iTipCount As Integer, Optional staticCount As Boolean) As String
'*******************************************************
' Purpose:  Put the count of tips into the status box
' Assumes:  The inclusion of StatusMsgDisplay
' Inputs:   the number of tips being displayed or a flag (when set to true)
'           to redisplay the last number of tips
' Returns:  a string that displays the number of tips
' Author:   James R. Fleming    Date:1/25/2000
'*******************************************************
    Static iTips As Integer
    Dim strMessage
    If staticCount = True Then
        iTipCount = iTips
    End If
    Select Case iTipCount
        Case 0
            strMessage = "There are no tips."
        Case 1
            strMessage = "There is one tip."
        Case Else
            strMessage = "There are " & iTipCount & " tips."
    End Select
    iTips = iTipCount
    TipCount = strMessage

End Function
Public Sub StatusMsgFlip(strFirstMsg As String, Optional inSeconds As Integer = 1, Optional intTipCount As Integer, Optional blnTipCount As Boolean)
'*****************************************************
' Purpose:  this sub displays a message in the status bar
'           then changes it after 1 secs by calling TipCount.
' Inputs:   The string being displayed first, and EITHER the number of tips OR the flag set to true.
'           if both optional inputs are not entered then the number of tips will be 0.
' Returns:  a call to StatusMsgDisplay and TipCount
'*****************************************************
    Dim strTime As String                       ' dim a string
    strTime = Time                              ' set the string to the current time
    Call StatusMsgDisplay(strFirstMsg, 2)
    Do Until DateDiff("s", strTime, Time) > inSeconds   ' display for x seconds
      DoEvents                                  ' releases the processor to continue application
    Loop
    ' it is possible for the form to hold up space in memory if unloaded while messages threads are pending in a loop
    ' there for we test to see if fMDI has been unloaded. It it is we bust out of here.
    If g_blnUnload = True Then Exit Sub
    If blnTipCount = True Then
        Call StatusMsgDisplay(TipCount(, blnTipCount), 2)
    Else
        Call StatusMsgDisplay(TipCount(intTipCount), 2)
    End If
End Sub
Public Sub StatusFlip(strFirstMsg As String, strSecondMsg As String, PanelIndex As Integer, Optional inSeconds As Integer = 1)
'*****************************************************
' Purpose:  this sub displays a message in the status bar
'           then changes it after 1 secs by calling TipCount.
' Inputs:   The string being displayed first, and EITHER the number of tips OR the flag set to true.
'           if both optional inputs are not entered then the number of tips will be 0.
' Returns:  a call to StatusMsgDisplay and TipCount
'*****************************************************
    Dim strTime As String                       ' dim a string
    strTime = Time                              ' set the string to the current time
    Call StatusMsgDisplay(strFirstMsg, PanelIndex)
    Do Until DateDiff("s", strTime, Time) > inSeconds   ' display for specificed seconds
      DoEvents                                  ' releases the processor to continue application
    Loop
    ' it is possible for the form to hold up space in memory if unloaded while messages threads are pending in a loop
    ' there for we test to see if fMDI has been unloaded. It it is we bust out of here.
    If g_blnUnload = True Then Exit Sub
    Call StatusMsgDisplay(strSecondMsg, PanelIndex)

End Sub
