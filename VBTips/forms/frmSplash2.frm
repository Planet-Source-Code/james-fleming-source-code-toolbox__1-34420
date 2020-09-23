VERSION 5.00
Begin VB.Form frmSplash2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash2.frx":000C
   ScaleHeight     =   7320
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSplash2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()

   ' lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   ' lblProductName.Caption = App.Title
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
