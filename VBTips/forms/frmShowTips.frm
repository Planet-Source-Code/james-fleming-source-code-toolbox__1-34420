VERSION 5.00
Begin VB.Form frmShowTips 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Tips "
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frmShowTips.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5745
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Press here to exit."
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Tips"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Press here to open the Tip of the day file. "
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Show tips next time at Startup"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Click Here to Show the Tips dialog box when the application starts."
      Top             =   1800
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.Label lblShowTips 
      BackStyle       =   0  'Transparent
      Caption         =   "lblShowTips"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmShowTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
On Error Resume Next
    If Check1.Value = 1 Then
     ' save whether or not this form should be displayed at startup
        SaveSetting App.Title, "Settings", "Show Tips at Startup", 1
    Else
        SaveSetting App.Title, "Settings", "Show Tips at Startup", 0
    End If
End Sub

Private Sub cmdOpen_Click()
Dim sFile As String

On Error GoTo tipserr
    sFile = App.path & "\files\" & "TIPOFDAY.txt"
    Dim RetVal
    
    RetVal = Shell("C:\WINNT\system32\Notepad.EXE  " & sFile, 3)     ' Run Notepad and load TIPOFDAY.txt.

tipserr:
    Select Case Err.Number
        Case 53
            RetVal = Shell("C:\Windows\Notepad.EXE  " & sFile, 3)     ' Run Notepad and load TIPOFDAY.txt.
        Case Else
            Resume Next
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Click()
Call StatusMsgDisplay("The Tip of the Day maintenance form is active.", 2)
End Sub

Private Sub Form_Load()
Dim ShowAtStartup As Long
    
    lblShowTips.Caption = "1. Press the Open Tips command button to open the tips file" & vbCrLf & _
    "2. You can add or remove tips" & vbCrLf & _
    "3. Put one tip per line of text" & vbCrLf & _
    "4. If the Tip of the Day form doesn't appear at start up" & vbCrLf & _
    "     " & "Click the check box below."

    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.Title, "Settings", "Show Tips at Startup", 1)
    Check1.Value = ShowAtStartup
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub
