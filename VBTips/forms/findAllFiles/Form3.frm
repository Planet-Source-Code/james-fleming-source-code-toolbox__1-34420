VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "About"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4425
   LinkTopic       =   "Form3"
   ScaleHeight     =   2175
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "elterrorista@videotron.ca"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "VB Mania"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Harvey      (VB Mania)"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

