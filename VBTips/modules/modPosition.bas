Attribute VB_Name = "modPosition"
Option Explicit

Public Sub AlignOpt(opt As Object)
'*****************************************************
' Purpose:  Set the alignment of the options into an m x n array
' Inputs:   None     Returns:  None
'*****************************************************
    opt(0).Left = 100                           ' first row
    opt(1).Left = opt(0).Left
    opt(2).Left = opt(0).Left
    opt(3).Left = opt(0).Left
    opt(4).Left = opt(0).Left + opt(0).Width    ' second row
    opt(5).Left = opt(4).Left
    opt(6).Left = opt(4).Left
    opt(7).Left = opt(4).Left
    opt(8).Left = opt(5).Left + opt(5).Width    ' third row
    opt(9).Left = opt(8).Left
    opt(10).Left = opt(8).Left
    opt(11).Left = opt(8).Left
End Sub

Public Sub ControlSize(myWidth As Long, lContainer As Long, txtBox As Object, cmbType As ComboBox, cmbSubType As ComboBox, labels As Object, lblWarn As Label, cmdButton As CommandButton)
'*****************************************************
' Purpose:  This is the movement handling specifically for the Details tab
' Inputs:   None    ' Returns:  None
'*****************************************************
    txtBox(0).Width = myWidth
    txtBox(1).Width = myWidth
    txtBox(2).Width = (myWidth - (cmdButton.Width + 305)) ' search text box
    cmbSubType.Width = (myWidth * 0.4)
    cmbType.Width = cmbSubType.Width
    lblWarn.Width = myWidth

    labels(0).Left = lContainer * 0.07823960880196    ' align the label controls
    labels(1).Left = labels(0).Left
    labels(2).Left = labels(0).Left
    labels(3).Left = labels(0).Left
    ' labels(4).Left depends on the placement of its control

    txtBox(0).Left = labels(0).Left + labels(0).Width + 150    ' align the text controls
    cmdButton.Left = ((txtBox(0).Left + txtBox(0).Width) - (cmdButton.Width))
    txtBox(1).Left = txtBox(0).Left
    txtBox(2).Left = txtBox(0).Left   ' search text box
    
    cmbType.Left = txtBox(0).Left
    cmbSubType.Left = (txtBox(0).Left + txtBox(0).Width) - (cmbSubType.Width)
    lblWarn.Left = txtBox(0).Left
    labels(4).Left = cmbSubType.Left - labels(4).Width - ((cmbSubType.Left - (cmbType.Left + cmbType.Width + labels(4).Width)) / 2)
        
End Sub

Public Sub FrameSize(myFrame As Frame, frmWidth As Long, frmLeft As Long, opt As Object)
'*****************************************************
' Purpose:  Set the size of the frame that contains the tip types.
'           Also sizes the option controls with the frame
' Inputs:   None     Returns:  None
'*****************************************************
    Dim i As Integer, j As Integer
    Dim temp As Integer
    
    myFrame.Width = frmWidth    ' ((cmdFind.Left + cmdFind.Width) - txtFields(0).Left)
    myFrame.Left = frmLeft      ' txtFields(0).Left
    ' find the length of the longest name
    For i = 0 To opt.Count - 1
        If Len(opt(i).Caption) > temp Then
            temp = Len(opt(i).Caption)
        End If
    Next i
    ' now that we know the width determine the divisor based on that width
    Select Case temp
        Case 0 To 6     ' divide width by 6
            temp = (myFrame.Width / 6)
        Case 7 To 10  ' divide width by 5
            temp = (myFrame.Width / 5)
        Case 11 To 15 ' divide width by 4
            temp = (myFrame.Width / 3) - 70
        Case 16 To 20 ' divide width by 3
            temp = (myFrame.Width / 3) - 100
        Case Else     ' divide width by 1
            temp = (myFrame.Width / 2)
    End Select
    ' set the width of each option based on that width.
    For i = 0 To opt.Count - 1
        opt(i).Width = temp
    Next i
   
End Sub

Public Sub SSTabResize(frm As Form)
'*****************************************************
' Purpose:  The purpose of this Sub is to manipulate
'           the tab control during resizing
' Inputs:   None    ' Returns:  None
' Comments: For dynamically resizing the tab control
'*****************************************************
On Error GoTo TabResizeErr
If frm.WindowState = vbMinimized Then Exit Sub
Dim iListWidth As Integer, iTabWidth As Integer

    'resizes the form and resizes controls, if necessary
    iListWidth = (frm.ScaleWidth / 3)    'set variables for the list and tab widths based on the form width
    iTabWidth = frm.ScaleWidth - iListWidth - 400
    If iTabWidth < 0 Then Exit Sub ' prevent illegal values
    'move and resize the list box and tab control
    frm.lstTitle.Move 100, 100, iListWidth, frm.ScaleHeight - 550
    frm.SSTab1.Move iListWidth + 200, 100, iTabWidth, frm.ScaleHeight - 550
    Exit Sub                ' that's all
TabResizeErr:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Public Sub NaviButtonsMove(cmd As Object, iTop As Integer, iWidth As Integer)
'*****************************************************
' Purpose:  move these buttons with the bottom of the form during resizing
' Inputs:   None    ' Returns:  None
'*****************************************************
' move = left , top, width, height
    cmd(0).Move cmd(0).Left, iTop, iWidth                 ' first
    cmd(1).Move (cmd(0).Left + iWidth), iTop, iWidth      ' previous
    cmd(2).Move (cmd(1).Left + iWidth), iTop, iWidth      ' next
    cmd(3).Move (cmd(2).Left + iWidth), iTop, iWidth      ' last
End Sub

Public Sub SortButtonsMove(iTop As Integer, ileft As Integer, iWidth As Integer, cmdButton As Object)
'*****************************************************
' Purpose:  move these buttons with the bottom of the form during resizing
' Inputs:   None    ' Returns:  None
'*****************************************************
' move = left , top, width, height
Dim i As Integer
    For i = 0 To cmdButton.Count - 1
        cmdButton(i).Move ileft + (i * iWidth), iTop, iWidth     ' view type
    Next

End Sub


