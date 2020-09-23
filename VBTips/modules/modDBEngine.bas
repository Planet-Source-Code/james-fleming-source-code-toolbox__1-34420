Attribute VB_Name = "modDBEngine"
Option Explicit
'*****************************************************
' Purpose:   The purpose of this module is to work with the Jet Engine.
'           The heirarch is:
'               DBEngine
'                    Workspace
'                        Database
'                            Recordset
' Inputs:   You need to include Microsoft DAO 3.5 Object libray from
'           the Project/References menu.
' Comments: Although this app was built around using DAO, If I have time, I'll include examples using ADO for comparison.
'*****************************************************

Public Function DataInit() As Boolean
'*****************************************************
' Purpose:   The purpose of this Function is to create a connection
'        to the database. DataInit will initialize the public Workspace Object
'
' Inputs:   You need to include Microsoft DAO 3.5 Object libray from
'           the Project/References menu.
' Assumes:  modConstants
' Returns:  A boolean value of true if opened successfully.
' Comments: You can pass in a user id and password string if the app is to be secured.
'
'*****************************************************

    On Error GoTo DataInit_EH           ' set the error handler
    DataInit = True                     ' set the flag to true
    Set gws = DBEngine.Workspaces(0)    ' open the first workspace
    Exit Function                       ' exit the function
DataInit_EH:
Select Case Err.Number
    Case 3356 ' database that is already opened exclusively
        ErrMsgBox (Err.Description & vbNewLine & "In DataInit() in modDBEngine.bas")
        End
    Case Else
    DataInit = False
        ErrMsgBox ("An error has occured at the Function DataInit() in modDBEngine.bas")
        Exit Function
    End Select
End Function
Public Function DataOpen(ByVal strDBName As String) As Boolean
'*****************************************************
' Purpose:   The purpose of this Function is to open the
'        database. DataOpen will open an Access database
'
' Assumes:  You need to include Microsoft DAO 3.5 Object libray from
'           the Project/References menu.
' Assumes:  modConstants
' Inputs: the gsDatabase variable which contains the database path & name
'           from Sub Main() in modStartup.
' Returns: a boolean value of true if opened successfully.
' Comments: DataOpen can only open an access database.
'       Error handling must be done because so much can go awry.
'*****************************************************
    On Error GoTo DataOpen_EH           ' set the error handler
   ' Call DataInit                      ' initialize the workspace
    Screen.MousePointer = vbHourglass   ' disable the mouse
    DataOpen = True                     ' change flag to true
    Set gdb = gws.OpenDatabase(strDBName) ' attempt to open the database
    Screen.MousePointer = vbDefault     ' reset the pointer
    Exit Function
DataOpen_EH:
    Select Case Err.Number
        Case 3356 ' database that is already opened exclusively
            ErrMsgBox (Err.Description & vbNewLine & "Error occured in Function DataInit() in modDBEngine.bas")
            End
        Case Else
            DataOpen = False
            Screen.MousePointer = vbDefault
            ErrMsgBox ("An error has occured at the Function DataOpen() in modDBEngine.bas")
            End
    End Select
End Function

Public Function DatabaseRepair(ComDial As CommonDialog)

'*****************************************************
' Purpose:   You can use the RepairDatabase method to fix
'   Corrupted database files.l The default syntax to invoke this method is
'   dbEngine.RepairDatabase databasename
' Assumes:  modConstants
' Inputs: You must pass in the Common Dialog control from the calling form.
'   It also uses the system constants for the project properties.
' Returns:   None
'*****************************************************

    Dim strDBName As String
    
    ComDial.FileName = App.path & gsMyDBase
    ComDial.DialogTitle = App.Title & " " & App.Major & ". " & App.Minor
    ComDial.ShowOpen
    strDBName = ComDial.FileName
    If Len(strDBName) Then
        DBEngine.RepairDatabase strDBName
        MsgBox App.path & vbCrLf & gsMyDBase & " Has been Repaired"
    End If
End Function

Public Sub DatabaseCompact()
'*****************************************************
' Purpose:   The Compact method clean out empy space in Jet databases
'       and performs general optimization chores that improve access speed.
'       you can also use compact to convert old versions of Jet to newer versions.
' Inputs: None   Returns:   None
'*****************************************************
    Dim strOldDBName As String
    Dim strNewDBName As String
    Dim strVersion As String
    Dim strHeader As String
    Dim intEncrypt, intVersion, intRet As Integer

DBCompactStart:
    '
    ' int vars
    strOldDBName = ""
    strNewDBName = ""
    strVersion = ""
    strHeader = "Compact Database Example"
    '
    ' get db to read
    ftblTips.CommonDialog1.DialogTitle = "Open Database to Write"
    ftblTips.CommonDialog1.Filter = "MS Jet | *.mdb"
    ftblTips.CommonDialog1.FileName = "VBTips.mdb"
    ftblTips.CommonDialog1.ShowOpen
    strNewDBName = ftblTips.CommonDialog1.FileName
    '
    If Trim(strNewDBName) = "" Then GoTo DBCompactStart
    '
    ' get target vesion (must be same or higher)
dbVersion:
    intRet = 0
    intVersion = 0
    strVersion = InputBox("Enter target version" & vbCrLf & "1.1, 2.0,2.5, 3.0, 3.5", strHeader)
    Select Case Trim(strVersion)
        Case "1.1"
            intVersion = dbVersion11
        Case "2.0"
            intVersion = dbVersion20
        Case "2.5"
            intVersion = dbVersion20
        Case "3.0"
            intVersion = dbVersion30
        Case "3.5"
            intVersion = dbVersion30
        Case ""
            Exit Sub
        Case Else
            ErrMsgBox ("Invalid version!  Version Error")
            Exit Sub
    End Select
    '
    ' encryption check
    intEncrypt = MsgBox("Encrypt this Database?", vbInformation + vbYesNo, strHeader)
    If intEncrypt = vbYes Then
        intEncrypt = dbEncrypt
    Else
        intEncrypt = dbDecrypt
    End If
    '
    ' now try to do it!
    DBEngine.CompactDatabase strOldDBName, strNewDBName, dbLangGeneral, intVersion + intEncrypt
        MsgBox "Process Completed"
    '
End Sub
