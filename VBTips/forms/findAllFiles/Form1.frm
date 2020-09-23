VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Programming Utilities"
   ClientHeight    =   5655
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   5160
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8916
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Search File"
      TabPicture(0)   =   "Form1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblSearch(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSearch(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSearch(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSearch(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Grille1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Combo1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Drive1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Search Function"
      TabPicture(1)   =   "Form1.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblSearch(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblSearch(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Grille2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Check1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "Form1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   -72720
         TabIndex        =   15
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search Function and Sub"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include ''_"" character"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -68880
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   2055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -71760
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Form1.frx":0054
         Left            =   -73440
         List            =   "Form1.frx":005E
         TabIndex        =   4
         Text            =   "*.FRM"
         Top             =   480
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid Grille1 
         Height          =   2895
         Left            =   -74640
         TabIndex        =   1
         Top             =   1320
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5106
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid Grille2 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
      End
      Begin VB.Label lblSearch 
         Caption         =   "VB Functions and Subs"
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   14
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblSearch 
         Caption         =   "Functions/Subs Found"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblSearch 
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -69840
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblSearch 
         Caption         =   "Drive :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -72480
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblSearch 
         Caption         =   "File Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblSearch 
         Caption         =   "Files"
         Height          =   375
         Index           =   5
         Left            =   -74520
         TabIndex        =   2
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Label lblSearch 
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   4080
      TabIndex        =   19
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblSearch 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label lblSearch 
      Caption         =   "Label0"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContent 
         Caption         =   "Help Content"
      End
      Begin VB.Menu mnuline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuSPMenu 
      Caption         =   "SPMenu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rightmouse As Boolean
Dim okMNU As Boolean
Dim posx As Single
Dim posy As Single

Private Sub Combo1_LostFocus()
    If Combo1.Text = "*.FRM" Then
        Command1.Caption = "Find FRM files"
    ElseIf Combo1.Text = "*.BAS" Then
        Command1.Caption = "Find BAS files"
    End If
End Sub

Private Sub Command1_Click()
    Grille2.Clear
    frmExportProg.Show 1
    okMNU = True
End Sub
Public Sub AddFunctionSub(ByVal item As String, ByVal fpath As String)
With Grille2
     .AddItem item
     .Row = .Rows - 1
     .Col = 1
     .Text = fpath
End With
End Sub
Private Sub Command2_Click()
    ProgressCancel = False
    Grille1.Clear
    Grille2.Clear
    Grille1.Rows = 1
    Grille2.Rows = 1
    InitGrille1
    InitGrille2
    NbFile = 0
    okMNU = False
    Command1.Enabled = False
    frmSearchProg.Show 1
    Command1.Enabled = True
    okMNU = True
End Sub


Private Sub Command3_Click()
End
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Form_Load()
    Command1.Enabled = False
    ProgressCancel = False
    
    Drive1.Drive = "C:\"
    InitGrille1
    InitGrille2
End Sub
Private Sub InitGrille2()
With Grille2
.Cols = 2
.Rows = 1
.Row = 0
.Col = 0
.Text = "Function/Sub"
.Col = 1
.Text = "File"
.ColWidth(0) = 4550
.ColWidth(1) = 4550
.Width = 4550 + 4550 + 350
End With
End Sub
Private Sub InitGrille1()
With Grille1
.Cols = 3
.Rows = 1
.Row = 0
.Col = 0
.Text = "File Name"
.Col = 1
.Text = "Path"
.Col = 2
.Text = "Size"
.ColWidth(0) = 2550
.ColWidth(1) = 4950
.ColWidth(2) = 1600
.Width = 1600 + 4950 + 2550 + 350
End With
End Sub
'**********************************************
'* Function FindFile is From Planet-Source-Code
'* Strongly modified by Carlos 09-10-99
'***********************************************
Public Sub FindFile(ByVal path As String, ByVal ftype As String)
       Dim hFile As Long, ts As String, WFD As WIN32_FIND_DATA
       Dim result As Long, sAttempt As String, szPath As String
       Dim strtemp
       If ProgressCancel Then Exit Sub
       frmSearchProg.ProgressBar1.Value = 1
       szPath = path & "*.*" & Chr$(0)
       'Start asking windows for files.
       putfileinpath path, ftype
       hFile = FindFirstFile(szPath, WFD)
       Do
         If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
          'Hey look, we've got a directory!
             ts = StripNull(WFD.cFileName)
             If Not (ts = "." Or ts = "..") Then
                 'Don't look for hidden or system directories
                 If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
                     'Search directory recursively
                     FindFile path & ts & "\", ftype
                 End If
             End If
           End If
           WFD.cFileName = ""
           result = FindNextFile(hFile, WFD)
           lblSearch(0).Caption = "Searching in: " & path
           DoEvents
          If frmSearchProg.ProgressBar1.Value = frmSearchProg.ProgressBar1.Max Then frmSearchProg.ProgressBar1.Value = 1
          frmSearchProg.ProgressBar1.Value = frmSearchProg.ProgressBar1.Value + 1
        Loop Until result = 0
       FindClose hFile
End Sub
'**********************************************
'* Function putfileinpath is From Planet-Source-Code
'* Modified by Carlos 09-10-99
'***********************************************
Private Sub putfileinpath(ByVal zpath As String, ByVal FileType As String)
       Dim hFile As Long, result As Long, szPath As String
       Dim WFD As WIN32_FIND_DATA
       szPath = zpath & FileType & Chr$(0)
       'Start asking windows for files.
       hFile = FindFirstFile(szPath, WFD)
       Dim pos1
       Do
           pos1 = InStr(1, WFD.cFileName, Chr$(0), vbBinaryCompare)
           If Trim(Mid(WFD.cFileName, 1, pos1 - 1)) <> "" Then
              AddAfile WFD, zpath
           End If
           WFD.cFileName = ""
           result = FindNextFile(hFile, WFD)
          ' DoEvents
       Loop Until result = 0
       FindClose hFile
End Sub

Private Sub AddAfile(WFDP As WIN32_FIND_DATA, ByVal path As String)
          NbFile = NbFile + 1
          With Grille1
             .AddItem Trim(WFDP.cFileName)
             .Row = NbFile
             .Col = 1
             .Text = path
             .Col = 2
             .Text = WFDP.nFileSizeLow / 1000 & " Kb   "
          End With
          lblSearch(1).Caption = NbFile & " Files Found."
End Sub


Private Sub Grille2_Click()
If okMNU Then
  If Not rightmouse Then
    'MsgBox Grille1.Row
    PopupMenu mnuSPMenu
    
    rightmouse = False
  End If
End If
End Sub

Private Sub Grille2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
 rightmouse = True
End If
End Sub

Private Sub mnuAbout_Click()
Form3.Show 1
End Sub

Private Sub mnuCopy_Click()
MsgBox "Put copy to clipboard code HERE!"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuView_Click()
    If Grille2.Col = 1 Then
      TypeView = 1
    Else
      TypeView = 0
    End If
    StringToFind = Grille2.Text
    Grille2.Col = 1
    FileFSToOpen = Grille2.Text
    frmReadProg.Show 1
End Sub


Private Sub TabStrip1_Click()
    If TabStrip1.SelectedItem = "Search File" Then
     Picture1.ZOrder
    Else
    Picture2.ZOrder
    End If
End Sub
