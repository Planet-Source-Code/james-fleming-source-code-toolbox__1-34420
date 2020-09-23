VERSION 5.00
Begin VB.Form frmTipOfDay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip of the Day"
   ClientHeight    =   2550
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5475
   Icon            =   "frmTipOfDay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   120
      Picture         =   "frmTipOfDay.frx":0442
      ScaleHeight     =   1935
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   120
      Width           =   3735
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   6
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "Show Tips at Startup"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdRandTip 
      Caption         =   "&Random Tips"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTipOfDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tips As New Collection      ' The in-memory database of tips.
Const TIP_FILE = "TIPOFDAY.TXT" ' Name of tips file
Dim CurrentTip As Long          ' Index in collection of tip currently being displayed.

Private Sub DoNextTip()
Dim iRnd As Integer
Static lastTip As Integer

    iRnd = RandomInterval(0, Tips.Count)     ' create a random number within a range
    CurrentTip = iRnd                        ' Select a tip at random.
    If CurrentTip = lastTip Then DoNextTip   ' don't display the same one twice
    If Tips.Count < CurrentTip Then          ' if we somehow went beyond the end
        CurrentTip = 1                       ' start again
    End If
    lastTip = CurrentTip
    frmTipOfDay.DisplayCurrentTip     ' Show it.
    
End Sub
Private Function RandomInterval(ByVal Min As Long, ByVal Max As Long) As Long
    'Returns a random integer Min <= N <= Max
    Randomize   ' Required function to reseed Rnd which is the multiplier
    RandomInterval = Int((Max - Min + 1) * Rnd + Min)
End Function
Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    Do Until EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Loop
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveSetting App.Title, "Settings", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdNext_Click()
    ' Or, you could cycle through the Tips in order

    CurrentTip = CurrentTip + 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    frmTipOfDay.DisplayCurrentTip    ' Show it.
End Sub

Private Sub cmdRandTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
    ftblTips.Show
End Sub

Private Sub Form_Activate()
    Call StatusMsgDisplay("The TipofDay form is active.", 2)
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title & " Tip of the Day"
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbChecked
       
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.path & "\files\" & TIP_FILE) = False Then
        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
        "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
        "Then place it in the same directory as the application. "
    End If

End Sub

Public Sub DisplayCurrentTip()
    If CurrentTip > 0 And CurrentTip < Tips.Count - 1 Then
        lblTipText.Caption = Tips.item(CurrentTip)
    End If
End Sub
