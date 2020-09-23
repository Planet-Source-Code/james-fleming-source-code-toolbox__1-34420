VERSION 5.00
Begin VB.Form frmSplash 
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   6630
   ScaleWidth      =   9510
   Visible         =   0   'False
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Microsoft Microborgs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Tag             =   "Product"
      Top             =   600
      Width           =   2880
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: 1999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Tag             =   "Copyright"
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   600
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Warning: Resistance is Futile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Tag             =   "Warning"
      Top             =   6000
      Width           =   2130
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Windoze"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   7440
      TabIndex        =   2
      Tag             =   "Platform"
      Top             =   240
      Width           =   1920
   End
   Begin VB.Label lblCompany 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "James Robert Fleming"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6960
      TabIndex        =   1
      Tag             =   "Company"
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Code from CyberSpace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Tag             =   "Product"
      Top             =   240
      Width           =   4680
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
' Assumes: modFormEffects
    Call ImplodeForm(Me, 2, 500, 1) ' close the form on demand
End Sub

Private Sub Form_Load()
' Assumes: modFormEffects
    Call FormCenter(Me)         ' center the form in formUtilities
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.ProductName
    Call ExplodeForm(Me, 500)   ' display effect in modFormEffects
End Sub



