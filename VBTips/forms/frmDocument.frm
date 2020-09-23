VERSION 5.00
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdDoc 
      Caption         =   "HTML Template"
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtText 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdDoc 
      Caption         =   "VB Module Template"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fLoad As Boolean
Public fDirty As Boolean
Private Sub cmdDoc_Click(Index As Integer)
'*******************************************************
' Purpose: adds a comments header block to the doc
' inputs:   None Returns: None.
'*******************************************************
    Dim strHeader As String
    Dim strTime As String                       ' dim a string
    strTime = Date                              ' set the string to the current time
    Select Case Index
        Case 0  ' VB Template
            strHeader = "'*******************************************************" & vbCrLf & _
            "' Purpose:  (Req)" & vbCrLf & _
            "'" & vbCrLf & _
            "' Title:" & vbCrLf & _
            "' Keywords:" & vbCrLf & _
            "' Language: Visual Basic               SubType: " & vbCrLf & _
            "' Assumes:" & vbCrLf & _
            "' Effects:" & vbCrLf & _
            "' Inputs:   (Req)                Returns: (Req)" & vbCrLf & _
            "' Comments:" & vbCrLf & _
            "' Depends:" & vbCrLf & _
            "' Author:   James R. Fleming    Date: " & strTime & vbCrLf & _
            "'*******************************************************" & vbCrLf & _
            "On error goto myErrorHandler" & vbCrLf
            txtText = strHeader & txtText ' append any existing data to the bottom of comments
            txtText = txtText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                "    Exit sub" & vbCrLf & _
                "myErrorHandler:" & vbCrLf & _
                "    Select case err.number" & vbCrLf & _
                "        Case Else" & vbCrLf & _
                "            ErrMsgBox (Err.Description & Err.Number)" & vbCrLf & _
                "            Resume Next" & vbCrLf & _
                "    End Select" & vbCrLf & _
                "End Sub"
        Case 1 ' HTML template
        strHeader = "<!DOCTYPE HTML PUBLIC ''-//W3C//DTD HTML 4.0 Transitional//EN'' ''http://www.w3.org/TR/REC-html40/loose.dtd''>" & vbCrLf & _
            "<!-- This entire site and all contents (C) Copyright 2000 -->" & vbCrLf & _
            "<!-- All U.S. Copyright laws apply in full effect.  -->" & vbCrLf & _
            "<HTML>" & vbCrLf & _
            "<HEAD>" & vbCrLf & _
            "<title></title>" & vbCrLf & "<!-- Begin Meta Data-->" & vbCrLf & _
            vbTab & "<meta http-equiv=''Content-Type'' content=''text/html; charset=iso-8859-1''>" & vbCrLf & _
            vbTab & "<META NAME=''Author'' CONTENT='' ''>" & vbCrLf & _
            vbTab & "<META NAME=''KeyWords'' CONTENT=''''>" & vbCrLf & _
            vbTab & "<META NAME=''Description'' CONTENT=''''>" & vbCrLf & _
            vbTab & "<META NAME=''Summary'' CONTENT=''''>" & vbCrLf & _
            vbTab & "<BASE href=''''>" & vbCrLf & _
            "<!-- Insert style sheet-->" & vbCrLf & _
            "<link rel=''stylesheet'' type=''text/css'' href=[path goes here] >" & vbCrLf & vbCrLf & _
            "<!-------------------- BEGIN THE JAVASCRIPT SECTION HERE -------------------->" & vbCrLf & _
            "</HEAD>" & vbCrLf & "<BODY>" & vbCrLf
            txtText = strHeader & txtText ' append any existing data to the bottom of comments
            txtText = txtText & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
            "</BODY>" & vbCrLf & _
            "</HTML>"
    End Select
End Sub

Private Sub Form_Activate()
' Assumes: modFormUtilities
    Call StatusMsgDisplay("The document form is active.", 2)
    fLoad = False

End Sub

Private Sub Form_Click()
   Me.Caption = TextSave(txtText, fMDI.dlgMDI)
End Sub

Private Sub Form_Load()
    Move 0, 0, fMDI.ScaleWidth, fMDI.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If fDirty = True Then
        Dim intResponse As Integer
        intResponse = YesNo("The text has been changed. Save changes?")
        If intResponse = 1 Then Call FileSave(fMDI.dlgMDI, txtText, "Text|*.txt|All|*.*")
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 615
    cmdDoc(0).Move txtText.Left, txtText.Top + txtText.Height + 100
    cmdDoc(1).Move (cmdDoc(0).Left + cmdDoc(0).Width + 120), txtText.Top + txtText.Height + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub

Private Sub txtText_Change()
    If fLoad = False Then
        fDirty = True
    End If
End Sub
