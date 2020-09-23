Attribute VB_Name = "modConst"
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

' instanciate forms
'Public fMDI As MDIMain
'Public ftblTips As frmtblTips
'Public fSplash As frmSplash
'Public fAuthor As frmAuthor
'Public fSubType As New frmSubType
' vbcrlf is carriage return + line feed


'*******************************************************
' Purpose: Instanciate all forms
' Inputs:  None                Returns: None
' Assumes:  modConstants
' Comments: Use these names not the actual form names
'*******************************************************
Public Sub Instanciate()
'    Set fMDI = New MDIMain     'create a new instance of MDIMain
'    Set ftblTips = New frmtblTips   'create a new instance of frmtblTips
'    Set fSplash = New frmSplash     'create a new instance of frmtblTips
'    Set fAuthor = New frmAuthor     'create a new instance of frmAuthor
End Sub
