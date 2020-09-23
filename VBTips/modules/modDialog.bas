Attribute VB_Name = "modDialog"
Option Explicit
Option Base 1
Option Compare Text
Public gFileView() As New frmFileView
Public lDocumentCount As Long
Public iImgCount As Integer
Public gCurrentInstance As Long         ' Points to the current active form
Public gFileName() As String
Public Loading As Boolean
Public iControl As Integer
Dim strFileName As String                    'The title of the form's return value
Public Function FileOpen(dlgControl As CommonDialog, frm As Form, Optional strFile As String, Optional strFilter As String) As String
'*******************************************************
' Purpose:  For opening files of various types using the common dialog control.
'
' Assumes:  modError
' Inputs:   the form and the commonDialog Control    Returns:None
' Comments: The case statement calls subs containing the function (for readability)
' Dependant: The calling form has a dialog contol
' Author:   James R. Fleming    Date:
'*******************************************************
On Error GoTo FileOpenErr

Dim HoldExtension As String, myFilter As String
    If strFilter <> "" Then
        strFilter = strFilter & "|" & strFilter & "|" & "Graphics (GIF, JPE, JPG, JIF, TIFF)|*.bmp;*.gif;*.jpe;*.jpg;*.jif;*.tiff;*.tif |C & C++ (C, C++, H)|*.cpp;*.c;*.h|Web Pages (HTM, HTML)|*.htm; *.html|SQL|*.sql|Text (RTF, TXT)|*.rtf;*.txt|Visual Basic (BAS, CLS, FRM)|*.bas;*.cls; *.frm; *.frx"
    Else: strFilter = "Text (RTF, TXT)|*.rtf;*.txt|Graphics (GIF, JPE, JPG, JIF, TIFF)|*.bmp;*.gif;*.jpe;*.jpg;*.jif;*.tiff;*.tif |C & C++ (C, C++, H)|*.cpp;*.c;*.h|Web Pages (HTM, HTML)|*.htm; *.html|SQL|*.sql|Visual Basic (BAS, CLS, FRM)|*.bas;*.cls; *.frm; *.frx"
    End If
    With dlgControl
        .DialogTitle = App.Title & " Open Files"
        .FileName = strFile
        .Filter = strFilter
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Function
        End If
        strFileName = .FileName
    End With
    
    Screen.MousePointer = vbHourglass

    HoldExtension = UCase(ExtReturn(strFileName))
    Select Case HoldExtension
    Case "RTF"  ' rich text
        RichTextLoad (strFileName)
    Case "TXT", "BAS", "C", "CLS", "CPP", "FRM", "FRX", "H", "HTML", "HTM", "SQL"
        TextLoad (strFileName)
    Case "BMP", "JPG", "JPE", "GIF", "JIF"
        PictureLoad (strFileName)
    Case "TIFF", "TIF"
        ImgLoad (strFileName)
    Case Else
        InputInfoBox ("I am currently unable to open this type of file." & vbCrLf & "You may open  BAS, BMP, C, CPP, FRM, FRX, GIF, H, HTM, HTML, JPE, JPG, RTF, SQL, TIF, TIFF, TXT")
End Select
    Screen.MousePointer = vbDefault
    FileOpen = HoldExtension
    Exit Function
FileOpenErr:
    Select Case Err.Number
        Case 481
            InputInfoBox ("Unable to open files of this type. You may open RTF, TXT, BMP, JPG, JPE, GIF, TIFF, TIF")
            Exit Function
        Case 32755  ' user cancels the action
            Exit Function
        Case Else
            Call ErrMsgBox(Err.Description & " in FileOpen of ModDialog.")
        Resume Next
    End Select
End Function
Public Sub RTFTXTSave(dlgControl As CommonDialog, rtb As RichTextBox, strFilter As String, frm As Form)
'**************************************************************
'Purpose: To print text out to a file
' inputs: The textbox being saved, the common dialog being used and the filters for the dialog
' return: the new title for the active form. (Change the title to match the saved file name.
'**************************************************************
On Error GoTo myErrorHandler
    Dim varFile As Variant
    Dim strTemp As String
    Dim iResponse As Integer
file:
    dlgControl.Filter = strFilter               ' set the filter
    dlgControl.ShowSave                         ' show the common dialog
    If dlgControl.FileName = "" Then            ' test for null titles
        InputErrBox ("You did not enter any file name!")
        Exit Sub                                 ' end
    Else
        strFileName = Dir(dlgControl.FileName)
        strTemp = Mid(dlgControl.FileName, (Len(dlgControl.FileName) - Len(strFileName)) + 1, Len(strFileName))
            If strFileName = strTemp And strTemp <> "" Then
              iResponse = MsgBox("The file already exist do you want to replace it ?", vbQuestion + vbYesNoCancel, "File Overwrite !!")
                If iResponse = vbNo Then        ' start again
                  GoTo file:
                ElseIf iResponse = vbYes Then   ' replace it
                  Kill dlgControl.FileName      ' destroy the existing file
                Else
                  Exit Sub                      ' user canceled the operation
                End If
            End If
            If strFilter = "RTF" Then
                rtb.SaveFile dlgControl.FileName, rtfRTF
            Else
                rtb.SaveFile dlgControl.FileName, rtfText
            End If
            frm.Caption = dlgControl.FileName
    End If
    Exit Sub
myErrorHandler:
 Select Case Err.Number
 Case 32755 ' Cancel was selected
   Exit Sub
 Case Else
    Call ErrMsgBox(Err.Description & " in FileSave of ModDialog.")
    Resume Next
 End Select
End Sub
Public Function TextSave(txtBox As String, dlgControl As CommonDialog, Optional strFilter As String = "Text|*.txt|All|*.*") As String
'**************************************************************
'Purpose: To print text out to a file
' inputs: The textbox being saved, the common dialog being used and the filters for the dialog
' return: the new title for the active form. (Change the title to match the saved file name.
'**************************************************************
On Error GoTo myErrorHandler                    'Handle an error.
    Dim intFile As Integer, iResponse As Integer
    Dim strTemp As String
file:                                           ' goto label
    With dlgControl                             ' set multiple dialog properties
        .Filter = strFilter                     ' set the filter
        .DialogTitle = App.Title & " Save As"   ' Set the title
        .ShowSave                               ' show it
    End With                                    ' end with the common dialog
    If dlgControl.FileName = "" Then            ' test for null titles
        InputErrBox ("You did not enter any file name!")
        Exit Function                            ' end
    Else
        strFileName = Dir(dlgControl.FileName)
        strTemp = Mid(dlgControl.FileName, (Len(dlgControl.FileName) - Len(strFileName)) + 1, Len(strFileName))
        If strFileName = strTemp And strTemp <> "" Then
            iResponse = MsgBox("The file already exist do you want to replace it ?", vbQuestion + vbYesNoCancel, "File Overwrite !!")
            If iResponse = vbNo Then        ' start again
                GoTo file:                    ' be very careful with this sort of thing!
            ElseIf iResponse = vbYes Then   ' replace it
                Kill dlgControl.FileName      ' destroy the existing file
            Else
                Exit Function                 ' user canceled the operation
            End If
        End If
        intFile = FreeFile
        Open dlgControl.FileName For Append As #intFile ' append doesn't stick me with additional " at the beginning and end of my file!
        Print #intFile, txtBox
        Close #intFile
        TextSave = FileNameReturn(dlgControl.FileName)  ' return a value to be the name of the active form
        Exit Function
    End If
    Exit Function
myErrorHandler:
    Select Case Err.Number
        Case 32755 ' Cancel was selected
          Exit Function
        Case Else
           Call ErrMsgBox(Err.Description & " in TextSave of ModDialog.")
           Resume Next
    End Select
End Function
Public Sub TextSaveChanges(txtBox As String)
'**************************************************************
'Purpose: To print text out to a file
' inputs: The textbox being saved, the common dialog being used and the filters for the dialog
' return: the new title for the active form. (Change the title to match the saved file name.
'**************************************************************
On Error GoTo myErrorHandler                    'Handle an error.
    Dim intFile As Integer, iResponse As Integer
    Dim strTemp As String
    Kill strFileName      ' destroy the existing file
    intFile = FreeFile
    Open strFileName For Append As #intFile ' append doesn't stick me with additional " at the beginning and end of my file!
    Print #intFile, txtBox
    Close #intFile
    Exit Sub
myErrorHandler:
    Select Case Err.Number
        Case 32755 ' Cancel was selected
          Exit Sub
        Case Else
           Call ErrMsgBox(Err.Description & " in TextSaveChanges of ModDialog.")
           Resume Next
    End Select
End Sub
Public Function FileSave(dlgControl As CommonDialog, ctrAny As Control, Optional strFilter As String = "Text|*.txt|Any|*.*") As String

'*****************************************************
' Purpose: This uses the common dialog control to save a file to a folder
'
' Inputs:  The CommonDialog, the textbox with the file name and the filter.
' Assumes: modError
' Returns: None
' comment: The filter can be any list of file types ie: "Text|*.txt|Modules|*.bas|Class|*.cls|Forms|*.frm"
'*****************************************************
On Error GoTo myErrorHandler
    Dim varFile As Variant
    Dim strTemp As String, strFileName As String
    Dim iResponse As Integer
file:
    dlgControl.Filter = strFilter               ' set the filter
    dlgControl.ShowSave                         ' show the common dialog
    If dlgControl.FileName = "" Then            ' test for null titles
        modMsgBox.InputInfoBox ("You did not enter a file name or save operation was canceled.")
        Exit Function                           ' end
    Else
        strFileName = Dir(dlgControl.FileName)
        strTemp = Mid(dlgControl.FileName, (Len(dlgControl.FileName) - Len(strFileName)) + 1, Len(strFileName))
            If strFileName = strTemp And strTemp <> "" Then
              iResponse = MsgBox("The file already exist do you want to replace it ?", vbQuestion + vbYesNoCancel, "File Overwrite !!")
                If iResponse = vbNo Then        ' start again
                  GoTo file:
                ElseIf iResponse = vbYes Then   ' replace it
                  Kill dlgControl.FileName      ' destroy the existing file
                Else
                  Exit Function                 ' user canceled the operation
                End If
            End If
         varFile = FreeFile
         EndStatement = "End Sub"
         If InStr(1, StringToFind, "Function") <> 0 Then EndStatement = "End Function"
         Open dlgControl.FileName For Append As varFile
         Write #varFile, Right(Left(ctrAny, Len(ctrAny) - 1), Len(ctrAny) - 1)
         Close varFile
    End If
    FileSave = strFileName  ' return a value to be the name of the active form
    Exit Function
myErrorHandler:
 Select Case Err.Number
 Case 32755 ' Cancel was selected
   Exit Function
 Case Else
    Call ErrMsgBox(Err.Description & " in FileSave of ModDialog.")
    Resume Next
 End Select
End Function

Private Sub RichTextLoad(strFileName As String)
    lDocumentCount = lDocumentCount + 1         ' increment the name of the doc
    ReDim gFileView(lDocumentCount)
    ReDim Preserve gFileView(lDocumentCount)
    Loading = True
    With gFileView(lDocumentCount)
        .rtbFileView.Move 0, 0, .ScaleWidth, .ScaleHeight
        .rtbFileView.LoadFile strFileName, rtfRTF
        .Caption = strFileName
        .Show
        .WindowState = vbMaximized
    End With
    gFileView(lDocumentCount).Picture1.Move -74760, 0
    gFileView(lDocumentCount).ImgEdit1.Move -74760, 0
End Sub

Private Sub TextLoad(strFileName As String)
    lDocumentCount = lDocumentCount + 1         ' increment the name of the doc
    ReDim gFileView(lDocumentCount)
    ReDim Preserve gFileView(lDocumentCount)
    Loading = True
    With gFileView(lDocumentCount)
        .WindowState = vbMaximized
        .rtbFileView.Move 0, 0, .ScaleWidth, .ScaleHeight
        .rtbFileView.LoadFile strFileName, rtfText
        .Caption = strFileName
        .VScroll1.Visible = False
        .rtbFileView.Visible = True
        .Show
    End With
    gFileView(lDocumentCount).Picture1.Move -74760, 0
    gFileView(lDocumentCount).ImgEdit1.Move -74760, 0
    Exit Sub
PicLoadErr:
Select Case Err.Number
    Case 50003 ' unexpected error
        ErrMsgBox ("An unexpected error occured and I am unable to open this file.")
        Exit Sub
 Case Else
    Call ErrMsgBox(Err.Description & " in TextLoad of modDialog.")
    Resume Next
 End Select
End Sub

Private Sub PictureLoad(strFileName As String)
On Error GoTo PicLoadErr
    Dim iTop As Integer
    Dim sLeft As Single
    iImgCount = iImgCount + 1         ' increment the name of the doc
    ReDim gFileView(iImgCount)
    ReDim Preserve gFileView(iImgCount)
    Loading = True
    With gFileView(iImgCount)
         If sLeft > (Screen.Width - .Picture1.Width) / 3 Then
            sLeft = sLeft
         Else: sLeft = ((Screen.Width - .Picture1.Width) / 3)
         End If
         If iTop > (Screen.Height - .Picture1.Height) / 3 Then
            iTop = iTop
         Else: iTop = ((Screen.Height - .Picture1.Height) / 3)
         End If
        .Hide
        .Picture1.Visible = True
        .VScroll1.Visible = False
        .rtbFileView.Move -74760    ' Hide the other controls.
        .ImgEdit1.Move -74760       ' Hide the other controls.
        .Picture1.Move 0, 1
        .Picture1.AutoSize = True
        .Picture1.Picture = LoadPicture(strFileName)
        .Move 100, 100, .Picture1.Width, .Picture1.Height
        .Caption = strFileName
        .Show
    '    .fLoad = False
    End With
    Exit Sub
PicLoadErr:
Select Case Err.Number
    Case 50003 ' unexpected error
        ErrMsgBox ("An unexpected error occured and I am unable to open this file.")
        Exit Sub
 Case Else
    Call ErrMsgBox(Err.Description & " in PictureLoad of modDialog.")
    Resume Next
 End Select
End Sub

Private Sub ImgLoad(strFileName As String)
On Error GoTo ImgLoadErr
    iImgCount = iImgCount + 1         ' increment the name of the doc
    ReDim gFileView(iImgCount)
    ReDim Preserve gFileView(iImgCount)
    Loading = True
    With gFileView(iImgCount)
        .WindowState = vbMaximized
        .VScroll1.Visible = False
        .HScroll1.Visible = False
        .rtbFileView.Move -74760    ' Hide the other controls.
        .Picture1.Move -74760       ' Hide the other controls.
        .ImgEdit1.Width = .ScaleWidth
        .ImgEdit1.Height = .ScaleHeight
        .ImgEdit1.Move 0, 0 ' the left = 0 is a flag being used by the frmFileView
        .ImgEdit1.Image = strFileName
        .Caption = strFileName
        .ImgEdit1.Display
        .Show
        .ImgEdit1.Visible = True
        .SetFocus
    End With
    Exit Sub
ImgLoadErr:
    Select Case Err.Number
        Case 50003 ' unexpected error
                ErrMsgBox ("An unexpected error occured and I am unable to open this file.")
                Exit Sub
         Case Else
            Call ErrMsgBox(Err.Description & " in ImgLoad of modDialog.")
            Resume Next
    End Select
End Sub
