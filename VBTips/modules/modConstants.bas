Attribute VB_Name = "modConstants"
Option Explicit
'*****************************************************
' Purpose:  Here we declare all global and public variables.
' Comment:  These are all the global vars used in the project
'*****************************************************
' Database vars
Global gsDatabase As String
Global gsConnect As String
'Global gsRecordsource As String
Public Const gsMyDBase As String = "\vbTips.mdb"
Public gws As Workspace
Public gdb As Database
Public g_blnUnload As Boolean
' instanciate forms
Public fAuthor As frmAuthor
Public fFileView As frmFileView
Public fLanguage As New frmLanguage
Public fMDI As MDIMain
Public fSearch As New frmSearch
Public fSplash As frmSplash
Public fSubType As New frmSubType
Public ftblTips As frmtblTips

' Query constants.
Public Const qryList As String = "SELECT tblTips.lngTblTipsID, tblTips.strTitle FROM tblTips ORDER BY tblTips.strTitle;"
Public Const qryOptName As String = "SELECT tblLanguage.intTable_PK, tblLanguage.strLang, tblLanguage.strToolTip FROM tblLanguage ORDER BY tblLanguage.intTable_PK;"
Public Const qryCombo As String = "SELECT tblLanguage.intTable_PK, tblLanguage.strLang, tblLanguage.strToolTip FROM tblLanguage ORDER BY tblLanguage.strLang;"
Public Const qryDateSort As String = "SELECT tblTips.* From tblTips ORDER BY tblTips.datTipDate DESC;"
Public Const qryWebLoad = "SELECT tblWebSites.lngWebID, tblWebSites.strSiteName, tblWebSites.strURL From tblWebSites ORDER BY tblWebSites.strSiteName;"
    
' vbcrlf is carriage return + line feed

'*******************************************************
' Purpose: Create instances all forms
' Inputs:  None                Returns: None
' Assumes:  modConstants
' Comments: Use these names not the actual form names
'*******************************************************
Public Sub Instanciate()
    Set fMDI = New MDIMain          'create a new instance of MDIMain
    Set ftblTips = New frmtblTips   'create a new instance of frmtblTips
    Set fSplash = New frmSplash     'create a new instance of frmtblTips
    Set fAuthor = New frmAuthor     'create a new instance of frmAuthor
End Sub

