VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************

' Name: Explode & Implode Your forms great for openings!

' Description:This code implodes & explodes forms

' By: Greg Henderson

'

'

' Inputs:none

'

' Returns:none

'

'Assumes:None

'

'Side Effects:none

'

'***************************************************************


Private Sub Command1_Click()
'In Command1

Call ImplodeForm(Me, 2, 500, 1)

'Set Form1 = Nothing
End Sub

Private Sub Form_Load()
'In form load

Call ExplodeForm(Me, 500)
End Sub

'in query unload


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call ImplodeForm(Me, 2, 500, 1)
End Sub

'There you have it e-mail me if you have any problems

 

