Attribute VB_Name = "modScrollBar"
Option Explicit
'*******************************************************
' Purpose:  This module is for manipulating a picture box with a status bar
'
' Assumes: Calling forms KeyPreview is set to true and
'           the 2 PictureBoxes are set AutoResize=True
' Comments: This control could easily be modified for moving anyother type of control
'           The first 2 subs (HScrollChange & VScrollChange could actually be placed in the change events of the scroll bars.
' Author:   Based on code by raulopez@hotmail.com taken from PlanetSourceCode. I made it far robust I must say.
'*******************************************************

Public Sub HScrollChange(inPic As PictureBox, hsb As Control)
    inPic.Left = -hsb.Value
End Sub

Public Sub VScrollChange(picInner As PictureBox, vsb As Control)
    picInner.Top = -vsb.Value
End Sub

Public Sub PicBoxResize(frm As Form, pic As PictureBox, vsb As Control, hsb As Control)
    On Error GoTo PicBoxResizeErr
    Dim iWidth As Integer
    iWidth = frm.ScaleWidth - (vsb.Width)
    
    ' Move Left, Top, Width, Height
    With vsb
        .Move (frm.ScaleWidth - vsb.Width), 0, 250, (frm.ScaleHeight)
        .Max = pic.Height - frm.ScaleHeight        ' Set VScrollBar Max
        .Visible = True
    End With
    With hsb
        .Move 0, (vsb.Height - 250), (frm.ScaleWidth - vsb.Width), 250
        .Max = pic.Width - frm.ScaleWidth          ' Set HScrollBar Max
        .Visible = True
    End With
    Call PicLoad(frm, pic, vsb, hsb)
    Exit Sub
PicBoxResizeErr:
    Select Case Err.Number
        Case 5     ' Invalid procedure call or argument
            Resume Next
        Case Else
            ErrMsgBox (Err.Description & " " & Err.Number & "  in PicBoxResize in modScrollBar.")
            Resume Next
        End Select
End Sub

Public Sub PicLoad(frm As Form, pic As PictureBox, vsb As Control, hsb As Control)
On Error GoTo PicLoadErr
'    'VERY IMPORTANT
    frm.KeyPreview = True
    If vsb.Max < 20 Then vsb.Max = 20
    
    'Set VScrollBar LargeChange and SmallChange
    vsb.LargeChange = vsb.Max \ 10
    If vsb.LargeChange < 10 Then
        vsb.SmallChange = 1
        vsb.Min = vsb.Max
    Else: vsb.SmallChange = vsb.LargeChange \ 5
    End If
    'You can set it to any value
    If hsb.Max < 20 Then hsb.Max = 20
    'Set HScrollBar LargeChange and SmallChange
    hsb.LargeChange = hsb.Max \ 10
    If hsb.LargeChange < 10 Then
        hsb.SmallChange = 1
        hsb.Min = hsb.Max
    Else: hsb.SmallChange = hsb.LargeChange \ 5
    End If
    If pic.Width < frm.Width Then hsb.Min = hsb.Max
    If pic.Height < frm.Height Then vsb.Min = vsb.Max
    Exit Sub
PicLoadErr:
    Select Case Err.Number
        Case 1
        Case 380

        Case Else
            Call ErrMsgBox(Err.Description & " " & Err.Number & " in PicLoad of modScrollbar")
        Resume Next
        
    End Select
End Sub

Public Sub PicScroll(KeyCode As Integer, vsb As Control, hsb As Control)
    Select Case KeyCode
        Case vbKeyUp
            If vsb.Value - vsb.SmallChange < 0 Then
                vsb.Value = 0
                'This will prevent overscrolling (Error 380)
            Else
                vsb.Value = vsb.Value - vsb.SmallChange
            End If
        Case vbKeyDown
            If vsb.Value + vsb.SmallChange > vsb.Max Then
                vsb.Value = vsb.Max
                'This will prevent overscrolling (Error 380)
            Else
                vsb.Value = vsb.Value + vsb.SmallChange
            End If
        Case vbKeyLeft
            If hsb.Value - hsb.SmallChange < 0 Then
                hsb.Value = 0
                'This will prevent overscrolling (Error 380)
            Else
                hsb.Value = hsb.Value - hsb.SmallChange
            End If
        Case vbKeyRight
                If hsb.Value + hsb.SmallChange > hsb.Max Then
                    hsb.Value = hsb.Max
                'This will prevent overscrolling (Error 380)
                Else
                    hsb.Value = hsb.Value + hsb.SmallChange
                End If
    End Select
End Sub
