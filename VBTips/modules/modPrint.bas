Attribute VB_Name = "modPrint"
Option Explicit

Public Sub PrintField(strPrint As String, dlg As CommonDialog)
  ' Print the selected file
  ' declare printer object variable
  Dim prtViewer As Printer
  dlg.DialogTitle = App.Title
  dlg.Flags = cdlPDReturnDC + cdlPDNoPageNums
  dlg.PrinterDefault = True
  dlg.CancelError = True
  If Len(strPrint) = 0 Then
    dlg.Flags = dlg.Flags + cdlPDAllPages + cdlPDNoSelection
  Else
    dlg.Flags = dlg.Flags + cdlPDSelection
  End If
  On Error Resume Next
  dlg.ShowPrinter
  If Err.Number <> 0 Then Exit Sub
  For Each prtViewer In Printers
    If prtViewer.DeviceName = Printer.DeviceName Then
      Set Printer = prtViewer
      Exit For
    End If
  Next
  Screen.MousePointer = vbHourglass
  Printer.Print strPrint
 ' Printer.hdc
  Printer.EndDoc
  Screen.MousePointer = vbNormal
End Sub
