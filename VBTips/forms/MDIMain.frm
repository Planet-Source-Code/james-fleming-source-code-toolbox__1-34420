VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   BackColor       =   &H00000000&
   Caption         =   "SourceCode from CyberSpace"
   ClientHeight    =   6510
   ClientLeft      =   -75
   ClientTop       =   165
   ClientWidth     =   7080
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIMain.frx":0442
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6255
      Visible         =   0   'False
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7937
            MinWidth        =   7937
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "There are no active forms."
            TextSave        =   "There are no active forms."
            Key             =   "TipNum"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "9:28 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPictures 
      Left            =   3000
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":84DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":85F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8702
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8926
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8C5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8D6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":90A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":91B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlPictures"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.Tag             =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.Tag             =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Description     =   "Left"
            Object.ToolTipText     =   "Left"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Center"
            Object.ToolTipText     =   "Center"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Description     =   "Right"
            Object.ToolTipText     =   "Right"
            ImageIndex      =   13
         EndProperty
      EndProperty
      Begin VB.ComboBox cmbTipType 
         Height          =   315
         ItemData        =   "MDIMain.frx":92C8
         Left            =   4560
         List            =   "MDIMain.frx":92CA
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Language"
         ToolTipText     =   "Select the source code type here."
         Top             =   52
         Width           =   1215
      End
      Begin VB.CommandButton cmdTlbrST 
         Caption         =   "Search"
         Height          =   315
         Left            =   8280
         TabIndex        =   0
         ToolTipText     =   "Press to view only the selected subtype."
         Top             =   52
         Width           =   855
      End
      Begin VB.ComboBox cmbTipSubType 
         Height          =   315
         Left            =   5880
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "Tip Category"
         ToolTipText     =   "Select the source code subtype here."
         Top             =   52
         Width           =   2295
      End
      Begin VB.Label lblPrintBuffer 
         Caption         =   "Label1"
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin MSComDlg.CommonDialog dlgMDI 
      Left            =   3600
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      FontSize        =   12
      Min             =   10
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuViewData 
      Caption         =   "Fo&rms"
      Begin VB.Menu mnuViewDatatblTips 
         Caption         =   "Code"
      End
      Begin VB.Menu mnuFError 
         Caption         =   "Error"
      End
      Begin VB.Menu mnuFSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuFWeb 
         Caption         =   "Web Sites"
      End
      Begin VB.Menu mnuAppMaint 
         Caption         =   "App Maintainence"
         Begin VB.Menu mnuFSubType 
            Caption         =   "SubType"
         End
         Begin VB.Menu mnuFExt 
            Caption         =   "Extension"
         End
         Begin VB.Menu mnuViewDatatblTipType 
            Caption         =   "Language"
         End
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuViewDatatblAuthor 
         Caption         =   "Author"
      End
      Begin VB.Menu mnuFClip 
         Caption         =   "Clipboard"
      End
      Begin VB.Menu mnuFReadMe 
         Caption         =   "Read Me"
      End
      Begin VB.Menu mnuV_DFSplash 
         Caption         =   "Splash Screen"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHTips 
         Caption         =   "Start Up Tips"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About VBTips..."
      End
   End
   Begin VB.Menu mnuImport 
      Caption         =   "Search"
      Visible         =   0   'False
      Begin VB.Menu mnuImpView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuImpCopy 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu mnuSearchFiles 
      Caption         =   "Search Files"
      Visible         =   0   'False
      Begin VB.Menu mnuSFOpen 
         Caption         =   "Open This File"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim strFileExt As String ' holds fileType of last file opened
Dim strPrintBuffer As String

'*******************************************************
' Purpose:  This is the MDI parent for this project
' Assumes:  modConstants, modControls, modFormEffects, modFormUtilities, modstartup, modDBEngine
' Effects:  Container for all mdi child forms
' Project:  Source_Code_Tool_Box
' Depends:  Is called from sub Main in modstartup
' Comment:  It is called by fMDI in modConstants
' Comment:  If you are setting your resolution to 1024 * 768 you will want to change your mdiMain.Picture to the AlienMoon1024 stored in the images directory.
' Authors:   James R. Fleming
'*******************************************************

Private Sub MDIForm_Initialize()
'*******************************************************
' Purpose: Call the saved setting from the registry for the toolbar & menu bar.
' Assumes: modFormUtilities
' Inputs: None                      Returns: None
' Comments: these changes are written to the windows registry
' you can open the resistry by: pressing the Start Key then pressing the Run menu option.
' ie.:  Start/Run/ in the Run command line window type regedit.exe
' you will find the setting saved under HKEY_CURRENT_USER/software/VB and VBA Program Settings
' (at least that's where it is on my machine.)
' Author:   James R. Fleming
'*******************************************************
    tbToolbar.Visible = GetSetting(App.Title, "Settings", "tbToolbarVisible", True)
    mnuViewToolbar.Checked = GetSetting(App.Title, "Settings", "mnuViewToolbarChecked", True)
    sbStatusBar.Visible = GetSetting(App.Title, "Settings", "sbStatusBarVisible", True)
    mnuViewStatusBar.Checked = GetSetting(App.Title, "Settings", "mnuViewStatusBarChecked", True)
End Sub

Private Sub MDIForm_Load()
'*******************************************************
' Purpose: Call the saved setting from the registry for the height & position.
' Inputs: None                      Returns: None
' Comments: these changes are written to the windows registry (see previous note in activate event.)
' Author:   James R. Fleming
'*******************************************************
On Error GoTo mdiMainFormErr
    If g_blnUnload = True Then Exit Sub
    ' recall the saved left property                            ' recall the saved Top property                         ' recall the saved width property                   ' recall the saved height property
    Move (GetSetting(App.Title, "Settings", "MainLeft", 1000)), (GetSetting(App.Title, "Settings", "MainTop", 1000)), (GetSetting(App.Title, "Settings", "MainWidth", 6500)), (GetSetting(App.Title, "Settings", "MainHeight", 6500))
    WindowState = (GetSetting(App.Title, "Settings", "WindowState", vbMaximized))
    Exit Sub
mdiMainFormErr:
    ErrMsgBox (Err.Description & " " & " Error # " & Err.Number & " has occured during the MDIMain form load event.")
    Resume Next
End Sub

Private Sub LoadNewDoc()
'*******************************************************
' Purpose: Add a Wordpad like document.
'
' Inputs:  None      Returns: None
' Comments:   Most of the requiste functionality hasn't been added.
' Author:   James R. Fleming
'*******************************************************
    Dim frmD As frmDocument                     ' dimension an instance of the form
    lDocumentCount = lDocumentCount + 1         ' increment the name of the doc
    Set frmD = New frmDocument                  ' set the new instance = to the form
    With frmD                                   ' set multiple properties
        .Caption = "Document " & lDocumentCount ' title the form
        .WindowState = vbMaximized              ' make it big
        .Show                                   ' show the instance of the form
    End With
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*******************************************************
' Purpose: Test the unload
' Assumes: modFormUtilities
' Inputs:  Cancel As Integer, UnloadMode As Integer     Returns: None
'*******************************************************
    g_blnUnload = True
    Call UnloadAllForms
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'*******************************************************
' Purpose: Save the setting to the registry for the size & position.
' Inputs: None                      Returns: None
' Comments: these changes are written to the windows registry (see previous note in activate event.)
' Author:   James R. Fleming
'*******************************************************
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Left
        SaveSetting App.Title, "Settings", "MainTop", Top
        SaveSetting App.Title, "Settings", "MainWidth", Width
        SaveSetting App.Title, "Settings", "MainHeight", Height
        SaveSetting App.Title, "Settings", "WindowState", WindowState
    End If
End Sub

Private Sub mnuFAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuFClip_Click()
    frmClipboard.Show
End Sub

Private Sub mnuFError_Click()
    frmErrorCodes.Show
End Sub

Private Sub mnuFExt_Click()
    frmExt.Show
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo SaveErr

    If Me.ActiveForm.Name = "frmtblTips" Then
        If ftblTips.cmdTip(2).Tag = "New" Or ftblTips.cmdTip(4).Tag = "Update" Then
            ftblTips.cmdTip_Click (4)
        Else
            Beep
        End If
    ElseIf Me.ActiveForm.Name = "frmFileView" Then
        With Me.ActiveForm
            If .Picture1.Left = 0 Then
                InputErrBox ("No save code associated with this action")
            ElseIf .rtbFileView.Left = 0 Then
                .rtbFileView.SaveFile dlgMDI.FileName, rtfText
                .fDirty = False
            ElseIf .ImgEdit1.Left = 0 Then
                Call FileSave(dlgMDI, .ImgEdit1, strFileExt & "|*." & strFileExt)
            Else
            End If
        End With
    ElseIf Me.ActiveForm.Name = "frmDocument" Then
        With Me.ActiveForm
            If Left(.Caption, 8) = "Document" Then
                .Caption = TextSave(.txtText.Text, Me.dlgMDI)
            Else
                Call TextSaveChanges(.txtText.Text)
            End If
            .fDirty = False
        End With
    End If
    Exit Sub
SaveErr:
    Select Case Err.Number
    Case 91
        Exit Sub
        Beep
    Case Else
        ErrMsgBox (Err.Description & " " & Err.Number & " in mnuFileSave of MDIMain")
        Resume Next
    End Select
End Sub

Private Sub mnuFReadMe_Click()
    frmFeatures.Show
End Sub

Private Sub mnuFSearch_Click()
    Set fSearch = New frmSearch     ' create a new instance of frmSearch
    fSearch.Show                    ' display it
End Sub

Private Sub mnuFSubType_Click()
    Set fSubType = New frmSubType ' create a new instance of frmSubType
    fSubType.Show
End Sub

Private Sub mnuFWeb_Click()
    frmWeb.Show
End Sub

Private Sub mnuImpCopy_Click()
    If fSearch.msgSearch(1).Col = 1 Then
      TypeView = 1
    Else
      TypeView = 0
    End If
    StringToFind = fSearch.msgSearch(1).TextMatrix(fSearch.msgSearch(1).Row, 0)
    fSearch.msgSearch(1).Col = 1
    FileFSToOpen = fSearch.msgSearch(1).TextMatrix(fSearch.msgSearch(1).Row, 1)
    frmProgress.myIndex = 4
    frmProgress.Show 1
End Sub

Private Sub mnuImpView_Click()
    If fSearch.msgSearch(1).Col = 1 Then
      TypeView = 1
    Else
      TypeView = 0
    End If
    StringToFind = fSearch.msgSearch(1).TextMatrix(fSearch.msgSearch(1).Row, 0)
    fSearch.msgSearch(1).Col = 1
    FileFSToOpen = fSearch.msgSearch(1).TextMatrix(fSearch.msgSearch(1).Row, 1)
    frmProgress.myIndex = 2
    frmProgress.Show 1
End Sub

Private Sub mnuSFOpen_Click()
On Error GoTo myErrorHandler
 '   Set fFileView = New frmFileView
    strFileExt = FileOpen(dlgMDI, fMDI, fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 1) & fSearch.msgSearch(0).TextMatrix(fSearch.msgSearch(0).Row, 0), fSearch.cmbExt.Text)
    Exit Sub
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub mnuViewStatusBar_Click()
'*******************************************************
' Purpose: Save the setting to the registry for the status bar.
' Inputs: None                      Returns: None
' Comments: these changes are written to the windows registry (see previous note in activate event.)
' Author:   James R. Fleming
'*******************************************************
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
        SaveSetting App.Title, "Settings", "sbStatusBarVisible", False
        SaveSetting App.Title, "Settings", "mnuViewStatusBarChecked", False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
        SaveSetting App.Title, "Settings", "sbStatusBarVisible", True
        SaveSetting App.Title, "Settings", "mnuViewStatusBarChecked", True
    End If
End Sub

Private Sub mnuViewToolbar_Click()
'*******************************************************
' Purpose: Save the setting to the registry for the tool bar.
' Inputs: None                      Returns: None
' Comments: these changes are written to the windows registry (see previous note in activate event.)
' Author:   James R. Fleming
'*******************************************************
    If mnuViewToolbar.Checked Then
        tbToolbar.Visible = False
        mnuViewToolbar.Checked = False
        SaveSetting App.Title, "Settings", "tbToolbarVisible", False
        SaveSetting App.Title, "Settings", "mnuViewToolbarChecked", False
    Else
        tbToolbar.Visible = True
        mnuViewToolbar.Checked = True
        SaveSetting App.Title, "Settings", "tbToolbarVisible", True
        SaveSetting App.Title, "Settings", "mnuViewToolbarChecked", True
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key

        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            If TypeOf Screen.ActiveControl Is TextBox Then
                Screen.ActiveControl.FontBold = _
                    Not Screen.ActiveControl.FontBold
            End If
        Case "Italic"
            If TypeOf Screen.ActiveControl Is TextBox Then
                Screen.ActiveControl.FontItalic = _
                    Not Screen.ActiveControl.FontItalic
            End If
        Case "Underline"
            If TypeOf Screen.ActiveControl Is TextBox Then
                Screen.ActiveControl.FontUnderline = _
                    Not Screen.ActiveControl.FontUnderline
            End If
        Case "Left"
            If TypeOf Screen.ActiveControl Is TextBox Then
                Screen.ActiveControl.Alignment = 0
            End If
        Case "Center"
            If TypeOf Screen.ActiveControl Is TextBox Then
                Screen.ActiveControl.Alignment = 2
            End If
        Case "Right"
            If TypeOf Screen.ActiveControl Is TextBox Then
                Screen.ActiveControl.Alignment = 1
            End If
    End Select
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    Arrange vbCascade
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuEditCopy_Click()
On Error GoTo myErrorHandler
    Clipboard.SetText Screen.ActiveControl.SelText
    Exit Sub
myErrorHandler:
    Select Case Err
        Case 438 ' Object doesn't support this property or method
            Resume Next
        Case Else
    ErrMsgBox (Err.Description & " " & Err.Number)
    End Select
End Sub

Private Sub mnuEditCut_Click()
On Error GoTo myErrorHandler
    Clipboard.SetText Screen.ActiveControl.SelText
    Screen.ActiveControl.SelText = ""
    Exit Sub
myErrorHandler:
    Select Case Err
        Case 438 'Object doesn't support this property or method
            Beep
            Resume Next
        Case Else
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
    End Select
End Sub

Private Sub mnuEditPaste_Click()
On Error GoTo myErrorHandler
    Screen.ActiveControl.SelText = Clipboard.GetText
    Exit Sub
myErrorHandler:
    Select Case Err
        Case 438 'Object doesn't support this property or method
            Beep
        Case 91
            Beep
            Resume Next
        Case Else
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
    End Select
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo myErrorHandler
    Set fFileView = New frmFileView
    strFileExt = FileOpen(dlgMDI, fFileView)
    Exit Sub
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub mnuFileClose_Click()
' Assumes: modFormUtilities
    Unload fMDI.ActiveForm
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo SaveErr
    If ActiveForm.Name = "frmtblTips" Then
        If ftblTips.cmdTip(2).Tag = "New" Or ftblTips.cmdTip(4).Tag = "Update" Then
            ftblTips.cmdTip_Click (4)
        Else
            InputErrBox ("No save code associated with this action")
        End If
    ElseIf ActiveForm.Name = "frmFileView" Then
        With ActiveForm
            If .Picture1.Left = 0 Then
                 InputErrBox ("No save code associated with this action")
            ElseIf .rtbFileView.Left = 0 Then
                Call RTFTXTSave(dlgMDI, .rtbFileView, strFileExt & "|*." & strFileExt, ActiveForm)
            ElseIf .ImgEdit1.Left = 0 Then
                InputErrBox ("No save code associated with this action")
            Else
                InputErrBox ("No save code associated with this action")
            End If
        End With
    ElseIf ActiveForm.Name = "frmDocument" Then
        With ActiveForm
            .Caption = FileSave(dlgMDI, .txtText, "Text|*.txt|All|*.*")
        End With
    End If
    Exit Sub
SaveErr:
    Select Case Err.Number
    Case 91
        Beep
    Case Else
        ErrMsgBox (Err.Description & " " & Err.Number & " in mnuFileSaveAs of MDIMain")
        Resume Next
    End Select
End Sub

Private Sub mnuFilePrint_Click()
On Error GoTo SaveErr
    If ActiveForm.Name = "frmtblTips" Then

        With ActiveForm
                strPrintBuffer = vbCrLf & vbCrLf & .txtFields(0).Text & vbCrLf & vbCrLf
            If .SSTab1.Tab = 0 Then
                strPrintBuffer = strPrintBuffer & .txtFields(3)
                Call PrintField(strPrintBuffer, dlgMDI)
            ElseIf .SSTab1.Tab = 1 Then
                Call PrintField(strPrintBuffer, dlgMDI)
            Else
                strPrintBuffer = vbCrLf & vbCrLf & "Title: " & .txtFields(0).Text & vbCrLf & vbCrLf
                strPrintBuffer = strPrintBuffer & "Keyword/website: " & .txtFields(1) & vbCrLf & vbCrLf
                strPrintBuffer = strPrintBuffer & "Language: " & .cmbTipType.Text & vbTab
                strPrintBuffer = strPrintBuffer & "Subtype: " & .cmbTipSubType.Text & vbCrLf & vbCrLf
                strPrintBuffer = strPrintBuffer & "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & vbCrLf
                If Len(.txtFields(3)) > 0 Then
                    strPrintBuffer = strPrintBuffer & "Notes: " & vbCrLf & .txtFields(3) & vbCrLf
                    strPrintBuffer = strPrintBuffer & "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------" & vbCrLf & vbCrLf
                End If
                If Len(.txtFields(4)) > 0 Then strPrintBuffer = strPrintBuffer & "Source Code: " & vbCrLf & vbCrLf & .txtFields(4) & vbCrLf & vbCrLf
            End If
            Call PrintField(strPrintBuffer, dlgMDI)
        End With
    ElseIf ActiveForm.Name = "frmFileView" Then
        With ActiveForm
            If .Picture1.Left = 0 Then
                  .PrintForm
            ElseIf .rtbFileView.Left = 0 Then
                Call PrintField(.rtbFileView, dlgMDI)
            ElseIf .ImgEdit1.Left = 0 Then
                .ImgEdit1.PrintImage
            Else
                InputErrBox ("No print function associated with this action")
            End If
        End With
    ElseIf ActiveForm.Name = "frmDocument" Then
        With ActiveForm
            Call PrintField(.txtText, dlgMDI)
        End With
    ElseIf ActiveForm.Name = "frmClipboard" Then
        With ActiveForm
            Call PrintField(.txtClipboard, dlgMDI)
        End With
    Else
        Dim iResponse As Integer
        iResponse = YesNo("There is no associated text or image with this form." & vbCrLf & "Shall I print out the active form?")
        If iResponse = 1 Then ActiveForm.PrintForm
    End If
    Exit Sub
SaveErr:
    Select Case Err.Number
    Case 91
        Beep
        Exit Sub
    Case Else
        ErrMsgBox (Err.Description & " " & Err.Number & " in mnuFileSaveAs of MDIMain")
        Resume Next
    End Select
End Sub
Private Sub mnuEditUndo_Click()
    If fMDI.ActiveForm.Name = "frmtblTips" Then
        If fMDI.ActiveForm.cmdTip(3).Caption = "C&ancel" Then
            fMDI.ActiveForm.cmdTip_Click (3)
        Else
            MsgBox ("No undo code associated with this action.")
        End If
    End If
End Sub
Private Sub mnuFileExit_Click()
    'unload the form
    Call MDIForm_QueryUnload(0, 0)
End Sub

Private Sub mnuHTips_Click()
    frmShowTips.Show
End Sub

Private Sub mnuV_DFSplash_Click()
    fSplash.Show
End Sub

Private Sub mnuViewDatatblTipType_Click()
 Set fLanguage = New frmLanguage   'create a new instance of frmtblTips
    fLanguage.Show
End Sub

Private Sub mnuViewDatatblTips_Click()
    Dim ftblTips As New frmtblTips
    ftblTips.Show
End Sub

Private Sub mnuViewRefresh_Click()
    On Error Resume Next
    Me.ActiveForm.Refresh
End Sub

Private Sub mnuViewDatatblAuthor_Click()
    Set fAuthor = New frmAuthor
    fAuthor.Show
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuViewOptions_Click()
    'To Do
    MsgBox "Options Dialog Code goes here!"
End Sub

Private Sub mnuFileNew_Click()
On Error GoTo myErrorHandler
    LoadNewDoc
    Exit Sub
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub cmbTipSubType_Change()
    cmdTlbrST.ToolTipText = "Press to view only " & Me.cmbTipType.Text & " " & Me.cmbTipSubType.Text & " tips."
End Sub

Private Sub cmbTipSubType_GotFocus()
    cmbTipSubType.Tag = cmbTipSubType.Text
End Sub

Private Sub cmbTipSubType_LostFocus()
    If cmbTipSubType.Tag <> cmbTipSubType.Text Then
        cmbTipSubType.Tag = cmbTipSubType.Text
        cmdTlbrST.Caption = "Search"
    End If
End Sub

Private Sub cmbTipType_Change()
    cmdTlbrST.ToolTipText = "Press to view only " & Me.cmbTipType.Text & " " & Me.cmbTipSubType.Text & " tips."
End Sub

Private Sub cmbTipType_GotFocus()
    cmbTipType.Tag = cmbTipType.Text    ' set the flag equal to the current value.
End Sub

Private Sub cmbTipType_LostFocus()
    If cmbTipType.Tag <> cmbTipType.Text Then   ' the flag is dirty
        cmbTipType.Tag = cmbTipType.Text        ' clean it
        cmdTlbrST.Caption = "Search"            ' we are on a new search
        Call SubTypeLoad(SubTypeQry(cmbTipType.ListIndex), cmbTipSubType) ' load the subtype list box.
        cmbTipSubType.ListIndex = 0         ' set the value = to the first record in the set
    End If
End Sub

Private Sub cmdTlbrST_Click()
Dim strSql As String
    If cmdTlbrST.Caption = "Search" Then    ' create the SQL statement
        strSql = qryToolBarSTcmb(cmbTipType.ListIndex, cmbTipSubType.Text)
        Call StatusMsgDisplay(TipCount(ListPopulate(strSql, ftblTips.lstTitle)), 2)
        Call ftblTips.FirstRecordShow              ' populate the list
        cmdTlbrST.Caption = "All"
        cmdTlbrST.ToolTipText = "Press to view all tips."
        ftblTips.cmdTip(0).Caption = "View all"
    Else
        cmdTlbrST.Caption = "Search"
        cmdTlbrST.ToolTipText = "Press to view only the selected subtype."
        Call ftblTips.cmdTip_Click(0)
    End If
End Sub
