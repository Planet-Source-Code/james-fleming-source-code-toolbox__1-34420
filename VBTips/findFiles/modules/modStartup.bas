Attribute VB_Name = "modStartup"
Option Explicit
  
Sub Main()
    Dim blnDatabaseOK As Boolean
    gsDatabase = App.path & "\databases\VBTips.mdb" ' note the path is hard coded here and only here
    blnDatabaseOK = DataInit                        'open database
    If blnDatabaseOK Then
        blnDatabaseOK = DataOpen(gsDatabase)
    Else
        ErrMsgBox ("Database open didn't work") ' oops
        Resume Next
    End If

    Load frmSearch              ' load instance of MDIMain

End Sub
