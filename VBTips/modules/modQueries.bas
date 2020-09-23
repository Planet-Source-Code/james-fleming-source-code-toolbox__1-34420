Attribute VB_Name = "modQueries"
Option Explicit
'**********************************************************************************************************
'
'   The queries to the access database
'
' Note:     All query constants are currently in modConstants.
'
' Comment: All of the queries are stored here ideally would be stored proceedures in the database, or a separate file.
'          By at least putting them all here it brings me one step closer to making them stored proceedures.
'          What is best is to store them outside the compiled program and if the file is not found, then the can use these queries.
'          By using a stored proceedure over a compiled one it saves from having to recompile and redistribute the exe if a change is made in a query.
'**********************************************************************************************************

Public Function BuildRecordSet() As Recordset
'*****************************************************
' Purpose:  opens an instance of the record set
' Assumes:  modConstants
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
Dim strSql As String
    
    strSql = "SELECT tblTips.lngTblTipsID, tblTips.strTitle, tblTips.intTypeID, tblLanguage.strLang, tblLanguage.strToolTip, tblTips.strIndex, tblTips.datTipDate, tblNotes.lngNoteTipsFK, tblNotes.memNotes, tblCode.lngCodeTipsFK, tblCode.memCode, tblTips.lngSubTypeID, tblSubType.strSTTitle, tblSubType.strSTToolTip " & _
            "FROM tblLanguage INNER JOIN (((tblTips INNER JOIN tblSubType ON tblTips.lngSubTypeID = tblSubType.lngSubTypeID) LEFT JOIN tblCode ON tblTips.lngTblTipsID = tblCode.lngCodeTipsFK) LEFT JOIN tblNotes ON tblTips.lngTblTipsID = tblNotes.lngNoteTipsFK) ON tblLanguage.intTable_PK = tblTips.intTypeID " & _
            "ORDER BY tblTips.strTitle;"
    
    Set BuildRecordSet = gdb.OpenRecordset(strSql, dbOpenDynaset)
    Exit Function
myErrorHandler:
    ErrMsgBox ("Error has occured during Private sub 'BuildRecordSet' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
    Resume Next
End Function

Public Function ViewOneTypeTip(myOption As Integer) As String
'*****************************************************
' Purpose:  creates an sql statement
' Inputs:   optViewType.index, the listbox being populated is passed in to pass through to the next routine
' Returns:  a sql query def string to the calling function
'*****************************************************
    Dim strSql As String
    
    strSql = "SELECT tblTips.lngTblTipsID, tblTips.strTitle, tblTips.intTypeID, tblLanguage.strLang, tblLanguage.strToolTip, tblTips.strIndex, tblTips.datTipDate, tblNotes.memNotes, tblCode.memCode, tblTips.lngSubTypeID, tblSubType.strSTTitle, tblSubType.strSTToolTip " & _
            "FROM tblLanguage INNER JOIN (((tblTips INNER JOIN tblSubType ON tblTips.lngSubTypeID = tblSubType.lngSubTypeID) LEFT JOIN tblCode ON tblTips.lngTblTipsID = tblCode.lngCodeTipsFK) LEFT JOIN tblNotes ON tblTips.lngTblTipsID = tblNotes.lngNoteTipsFK) ON tblLanguage.intTable_PK = tblTips.intTypeID " & _
            "Where (((tblTips.intTypeID) = " & myOption & "))ORDER BY tblTips.strTitle;" ' create the SQL string
    ViewOneTypeTip = strSql
End Function

Public Function ViewBySubType(myOption As Integer) As String
'*****************************************************
' Purpose:  creates an sql statement
' Inputs:   optViewType.index
' Returns:  a sql query def string
'*****************************************************

    ViewBySubType = "SELECT tblTips.strTitle, tblSubType.strSTTitle, tblNotes.memNotes, tblCode.memCode, tblTips.lngTblTipsID, tblTips.datTipDate, tblTips.lngSubTypeID, tblLanguage.intTable_PK, tblLanguage.strLang, tblLanguage.strToolTip, tblSubType.strSTToolTip " & _
            "FROM ((tblSubType INNER JOIN (tblLanguage INNER JOIN tblTips ON tblLanguage.intTable_PK = tblTips.intTypeID) ON tblSubType.lngSubTypeID = tblTips.lngSubTypeID) LEFT JOIN tblCode ON tblTips.lngTblTipsID = tblCode.lngCodeTipsFK) LEFT JOIN tblNotes ON tblTips.lngTblTipsID = tblNotes.lngNoteTipsFK " & _
            "Where (((tblTips.intTypeID) = " & myOption & ")) ORDER BY tblSubType.strSTTitle;"

End Function
Public Function SubTypeQry(iTypeID As Integer) As String
'*****************************************************
' Purpose:  Create a SQL query def to return a type's subtype.
' Inputs:   The type ID
' Returns:  a SQL query def string to the calling function
'*****************************************************

    SubTypeQry = "SELECT tblSubType.strSTTitle, tblSubType.lngSubTypeID, tblSubType.intTypeID, tblSubType.strSTToolTip From tblSubType Where (((tblSubType.intTypeID) = " & iTypeID & ")) ORDER BY tblSubType.strSTTitle;"

End Function

Public Function SearchAll(strCriteria As String) As String
'*****************************************************
' Purpose:  find tips based on a database wide search (title, code & notes) for key word
' Inputs:   the search criteria   ' search text box
' Returns:  a SQL query def string
'*****************************************************

    SearchAll = "SELECT tblTips.strTitle, tblNotes.memNotes, tblCode.memCode, tblTips.lngTblTipsID, tblTips.datTipDate, tblTips.lngSubTypeID, tblSubType.strSTTitle, tblLanguage.intTable_PK, tblLanguage.strLang, tblLanguage.strToolTip, tblSubType.strSTToolTip " & _
            "FROM ((tblLanguage INNER JOIN (tblTips LEFT JOIN tblSubType ON tblTips.lngSubTypeID = tblSubType.lngSubTypeID) ON tblLanguage.intTable_PK = tblTips.intTypeID) LEFT JOIN tblCode ON tblTips.lngTblTipsID = tblCode.lngCodeTipsFK) LEFT JOIN tblNotes ON tblTips.lngTblTipsID = tblNotes.lngNoteTipsFK " & _
            "Where (((tblTips.strTitle) Like " & strCriteria & ")) Or (((tblTips.strIndex) Like " & strCriteria & ")) Or (((tblNotes.memNotes) Like " & strCriteria & ")) Or (((tblCode.memCode) Like " & strCriteria & ")) " & _
            "ORDER BY tblTips.strTitle;"
End Function
Public Function SearchTitle(strCriteria As String) As String
'*****************************************************
' Purpose:  find tips based on a database title only search for key word
' Inputs:   txtFields(2).text
' Returns:  a query def string to the calling function
'*****************************************************
    SearchTitle = "SELECT tblTips.lngTblTipsID, tblTips.strTitle FROM tblTips " & _
    "Where (((tblTips.strTitle) Like " & strCriteria & ")) Or (((tblTips.strIndex) Like " & strCriteria & "))ORDER BY tblTips.strTitle;" ' create the SQL string

End Function

Public Sub optTypeName(strSql As String, optArray As Object)
'*****************************************************
' Purpose:  loads the Frame option buttons with the titles
' Assumes:  modConstants
' Inputs:   the query string, the option buttons receiving their labels from that query
' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
    Dim i As Integer
    Dim oRS As Recordset
    
    Set oRS = gdb.OpenRecordset(strSql, dbOpenSnapshot)     'open the recordset

    Do Until oRS.EOF                                        'loop to add the titles to the listbox
        With optArray(i)
            .Caption = oRS!strLang
            .ToolTipText = oRS!strToolTip
        End With
        i = i + 1
        oRS.MoveNext
    Loop
    oRS.Close
    Exit Sub
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub
Public Function ExtLoad(comboIndex As Integer) As String

ExtLoad = "SELECT tblExtension.lngTable_PK, tblExtension.strLang, tblExtension.lngExtID " & _
    "From tblExtension Where (((tblExtension.lngTable_PK) = " & comboIndex & ")) " & _
    "ORDER BY tblExtension.lngExtID;"
    
End Function

Public Function qryLoadSubType(Index As Integer) As String

    qryLoadSubType = "SELECT tblSubType.intTypeID, tblSubType.strSTTitle, tblSubType.strSTToolTip, tblSubType.lngSubTypeID " & _
                 "FROM tblLanguage INNER JOIN tblSubType ON tblLanguage.intTable_PK = tblSubType.intTypeID " & _
                 "Where (((tblSubType.intTypeID) = " & Index & " )) " & _
                 "ORDER BY tblSubType.strSTTitle;"
End Function

Public Function qryUpdateType(Index As Integer) As String

    qryUpdateType = "SELECT tblSubType.intTypeID, tblSubType.strSTTitle, tblSubType.strSTToolTip, tblSubType.lngSubTypeID " & _
            "FROM tblLanguage INNER JOIN tblSubType ON tblLanguage.intTable_PK = tblSubType.intTypeID " & _
            "Where (((tblSubType.intTypeID) = " & Index & " )) " & _
            "ORDER BY tblSubType.strSTTitle;"
End Function

Public Function qryToolBarSTcmb(iTTListIndex As Integer, strSTText As String) As String
Dim strSql As String
        strSql = "SELECT tblTips.lngTblTipsID, tblTips.strTitle, tblTips.intTypeID, tblLanguage.strLang, tblLanguage.strToolTip, tblTips.strIndex, tblTips.datTipDate, tblNotes.memNotes, tblCode.memCode, tblTips.lngSubTypeID, tblSubType.strSTTitle, tblSubType.strSTToolTip " & _
                "FROM tblLanguage INNER JOIN (((tblTips INNER JOIN tblSubType ON tblTips.lngSubTypeID = tblSubType.lngSubTypeID) LEFT JOIN tblCode ON tblTips.lngTblTipsID = tblCode.lngCodeTipsFK) LEFT JOIN tblNotes ON tblTips.lngTblTipsID = tblNotes.lngNoteTipsFK) ON tblLanguage.intTable_PK = tblTips.intTypeID " & _
                "WHERE (((tblTips.intTypeID)=" & iTTListIndex & ") AND ((tblSubType.strSTTitle)="
        qryToolBarSTcmb = strSql & " """ & strSTText & """ " & "))" & _
                "ORDER BY tblSubType.strSTTitle;" ' it's tricky business passing quotes to SQl

End Function

Public Sub KillNote(iTipFK As Integer)
'*******************************************************
' Purpose:  To handle the removal of records in a child table without killing the parent record.
' Inputs:   The foreign key of the tip being altered Returns:None
' Comments: This allows me to keep my main table small
' Author:   James R. Fleming    Date: 3/7/200
'*******************************************************

Dim strSql As String
Dim rsKill As Recordset
On Error GoTo KillNoteErr
    strSql = "SELECT tblNotes.lngtblNotesID, tblNotes.lngNoteTipsFK, tblNotes.memNotes " & _
    "From tblNotes WHERE (((tblNotes.lngNoteTipsFK)=" & iTipFK & "));"
    
    Set rsKill = gdb.OpenRecordset(strSql, dbOpenDynaset)
    rsKill.Delete
    Exit Sub
KillNoteErr:
    Resume Next
End Sub

Public Sub KillCode(iTipFK As Integer)
'*******************************************************
' Purpose:  To handle the removal of records in a child table without killing the parent record.
' Inputs:   The foreign key of the tip being altered Returns: None
' Comments: This allows me to keep my main table small
' Author:   James R. Fleming    Date: 3/7/200
'*******************************************************
On Error GoTo KillCodeErr
Dim strSql As String
Dim rsKill As Recordset

    strSql = "Select tblCode.lngTblCodeID, tblCode.lngCodeTipsFK, tblCode.memCode " & _
    "From tblCode WHERE (((tblCode.lngCodeTipsFK)=" & iTipFK & "));"
    
    Set rsKill = gdb.OpenRecordset(strSql, dbOpenDynaset)
    rsKill.Delete
    Exit Sub
KillCodeErr:
    Resume Next
End Sub

Public Function TitleTestRS(strTitle As String, strWildCard As String) As Recordset
' myTipTitle ' strRipQuote
    Dim strSql As String
    strSql = "SELECT tblTips.strTitle FROM tblTips WHERE (((tblTips.strTitle) Like " & strTitle & " Or (tblTips.strTitle) Like " & strWildCard & " #" & """  Or (tblTips.strTitle) Like " & strWildCard & " ##""" & "));"
    Set TitleTestRS = gdb.OpenRecordset(strSql, dbOpenSnapshot) 'open the recordset
End Function
