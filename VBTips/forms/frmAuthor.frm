VERSION 5.00
Begin VB.Form frmAuthor 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Author: James Robert Fleming"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Symbol"
      Size            =   9.75
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmAuthor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3795
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3900
      Left            =   480
      Picture         =   "frmAuthor.frx":0442
      Top             =   360
      Width           =   4500
   End
End
Attribute VB_Name = "frmAuthor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call StatusFlip("Thank you for using my application.", "Email me at umbcifsm@netscape.net.", 2, 1)
End Sub

Private Sub Form_Load()
    Label1.Caption = "James Robert Fleming holds a Bachelor of Science in Information Systems from UMBC and may be reached at umbcifsm@netscape.net." & vbCrLf & _
    "He acknowleges that this application would not be possible without the support of his wife and the many contributing authors at PlanetSourceCode.com."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub
