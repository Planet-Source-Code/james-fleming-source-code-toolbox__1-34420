Attribute VB_Name = "modTabControl"
Option Explicit
Public Sub TabControlsMove(frm As Form)
'*****************************************************
' Purpose:  Handle the moving and resizing of tab control
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
Dim lngWidth As Long

    If frm.SSTab1.Caption = "&Notes" Then                   ' if the Notes tab is forward, resize its text field
        frm.txtFields(3).Move 100, 600, (frm.SSTab1.Width - 200), (frm.SSTab1.Height - 900)
    ElseIf frm.SSTab1.Caption = "&Code" Then                ' if the Code tab is forward resize its text field
        frm.txtFields(4).Move 100, 600, (frm.SSTab1.Width - 200), (frm.SSTab1.Height - 900)
    Else                                                    ' SSTab1.Caption = "&Details"
        If frm.Width < 8565 Then frm.Width = 8565           ' the minimum size must be larger on the Details tab
        If frm.SSTab1.Width < 6015 Then                     ' If WindowState = vbNormal Or WindowState = vbMaximized Then
            lngWidth = frm.SSTab1.Width * 0.65              ' for a narrower form, a smaller multiple is in order
            Call ControlSize(lngWidth, frm.SSTab1.Width, frm.txtFields(), frm.cmbTipType, frm.cmbTipSubType, frm.lblDetail, frm.lblTitleWarn, frm.cmdFind)
            If frm.Frame1.Visible = True Then
                Call FrameSize(frm.Frame1, ((frm.cmdFind.Left + frm.cmdFind.Width) - frm.txtFields(0).Left), frm.txtFields(0).Left, frm.optViewType)        ' resize the frame based on the forms width
                Call AlignOpt(frm.optViewType())
            End If
        Else
            lngWidth = frm.SSTab1.Width * 0.745721271393643 'set the width of the controls
            Call ControlSize(lngWidth, frm.SSTab1.Width, frm.txtFields(), frm.cmbTipType, frm.cmbTipSubType, frm.lblDetail, frm.lblTitleWarn, frm.cmdFind)
            If frm.Frame1.Visible = True Then
                Call FrameSize(frm.Frame1, ((frm.cmdFind.Left + frm.cmdFind.Width) - frm.txtFields(0).Left), frm.txtFields(0).Left, frm.optViewType)         ' resize the frame based on the forms width
                Call AlignOpt(frm.optViewType())
            End If
        End If
    End If
    Exit Sub
myErrorHandler:
Select Case Err.Number
    Case 384 ' A form can't be moved or sized while minimized or maximized
        Call SSTabResize(frm)
        Resume Next
    Case Else
        ErrMsgBox ("Error has occured during Private sub 'TabControlsMove' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
        Resume Next
    End Select
End Sub

Public Sub SSTab1Focus(sst As SSTab, txtBox As Object)
'*****************************************************
' Purpose:  Position focus on the proper text box
' Inputs:   None    ' Returns:  None
'*****************************************************
    If sst.Caption = "&Notes" Then
        txtBox(3).SetFocus
    ElseIf sst.Caption = "&Code" Then
        txtBox(4).SetFocus
    Else
        txtBox(0).SetFocus
    End If
End Sub

