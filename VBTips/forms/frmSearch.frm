VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Programming Utilities"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.CheckBox chkSetDefault 
      Caption         =   "Save Settings"
      Height          =   255
      Left            =   10080
      TabIndex        =   22
      ToolTipText     =   "Remember your language preference and search path."
      Top             =   4800
      Width           =   1695
   End
   Begin VB.DriveListBox dlbDrive 
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   500
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Include ''_"" char"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10032
      TabIndex        =   19
      ToolTipText     =   "Include those functions and subs that use the underscore character."
      Top             =   4320
      Width           =   1668
   End
   Begin VB.ComboBox cmbTipType 
      Height          =   315
      ItemData        =   "frmSearch.frx":0442
      Left            =   10032
      List            =   "frmSearch.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   15
      Text            =   "Language"
      ToolTipText     =   "Select the source code type here."
      Top             =   480
      Width           =   1668
   End
   Begin VB.ComboBox cmbExt 
      Height          =   315
      ItemData        =   "frmSearch.frx":0446
      Left            =   9000
      List            =   "frmSearch.frx":0448
      TabIndex        =   14
      Text            =   "*.FRM"
      ToolTipText     =   "Select the source code extension type here."
      Top             =   500
      Width           =   855
   End
   Begin VB.TextBox txtSearch 
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Text            =   "path"
      ToolTipText     =   "Enter a path to search or use the browse feature."
      Top             =   500
      Width           =   7695
   End
   Begin VB.Frame frmFile 
      Caption         =   "File Search:"
      Height          =   1095
      Left            =   10032
      TabIndex        =   9
      ToolTipText     =   "You may search either a drive, a path or you may browse to open a folder."
      Top             =   1440
      Width           =   1668
      Begin VB.OptionButton optSearch 
         Caption         =   "Path"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Drive"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optSearch 
         Caption         =   "Browse"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame frmFunction 
      Caption         =   "Function Search:"
      Height          =   795
      Left            =   10032
      TabIndex        =   6
      ToolTipText     =   "Select one file to search or search the entire returned subset by selecting here."
      Top             =   3360
      Width           =   1668
      Begin VB.OptionButton optFunct 
         Caption         =   "All Files"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   840
      End
      Begin VB.OptionButton optFunct 
         Caption         =   "Selected"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   925
      End
   End
   Begin MSComDlg.CommonDialog dlgSearch 
      Left            =   1800
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Files"
      Height          =   375
      Index           =   0
      Left            =   10032
      TabIndex        =   5
      ToolTipText     =   "Find files of the specified type"
      Top             =   975
      Width           =   1668
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Functions && Subs"
      Height          =   375
      Index           =   1
      Left            =   10032
      TabIndex        =   4
      ToolTipText     =   "Currently you may only import functions and subroutines."
      Top             =   2760
      Width           =   1668
   End
   Begin TabDlg.SSTab sstSearch 
      Height          =   7095
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Files"
      TabPicture(0)   =   "frmSearch.frx":044A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSearch(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "msgSearch(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Procedures"
      TabPicture(1)   =   "frmSearch.frx":0466
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "msgSearch(1)"
      Tab(1).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid msgSearch 
         Height          =   6015
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   10610
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         GridLinesFixed  =   1
         AllowUserResizing=   3
      End
      Begin MSFlexGridLib.MSFlexGrid msgSearch 
         Height          =   6015
         Index           =   1
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   10610
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
      End
      Begin VB.Label lblSearch 
         AutoSize        =   -1  'True
         Caption         =   "Files:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   360
      End
   End
   Begin VB.Label lblSearch 
      Caption         =   "Drive"
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
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   465
   End
   Begin VB.Label lblSearch 
      Caption         =   "Type :"
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
      Index           =   9
      Left            =   10032
      TabIndex        =   18
      Top             =   240
      Width           =   1668
   End
   Begin VB.Label lblSearch 
      Caption         =   "Ext:"
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
      Left            =   9000
      TabIndex        =   17
      Top             =   240
      Width           =   495
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
      Left            =   1200
      TabIndex        =   16
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rightmouse As Boolean, okMNU As Boolean, fLoad As Boolean
Dim posx As Single, posy As Single

'*******************************************************
' Purpose:  This is the interface for importing code from preexisting projects.
' Assumes:  modConstants, modControls, modFindAllFiles, modFormUtilities, modstartup, modDBEngine
' Project:  Source_Code_Tool_Box
' Comment:  This part of the app is based loosely on source code found at Planet-source-code by Mark Joyal,
'           although for the record, I had planned to do this input part as part of the original,
'           yet here I have been able to save quite a bit of time by using what was done by others, so their names are left in where they occur
' Authors:   James R. Fleming
'*******************************************************
Private Sub Form_Activate()
    Call StatusFlip("The search form is activated.", "You may begin your search.", 2, 1)
End Sub

Private Sub Form_Load()
'*******************************************************
' Purpose:  Load the form
' Inputs : None     Returns : None.
'*******************************************************
    On Error GoTo myErrHandler
    Dim i As Integer
    Dim bflag As Boolean
    fLoad = True                ' set the on load flag
    Dim lngDefaults As Long
    Call ComboboxLoad(cmbTipType, qryCombo) ' loads the tips combo box

    ' See if we should be shown at startup
    lngDefaults = GetSetting(App.Title, "Search", "Save Search Defaults", 1)
    If lngDefaults = 0 Then  ' skip it and use hard coded values
        For i = 0 To Forms.Count - 1 ' test to see if the form is loaded without calling a new instance of it
        If Forms(i).Name = "frmtblTips" Then    ' found it
            bflag = True                        ' set a flag to that effect
            cmbTipType.ListIndex = ftblTips.cmbTipType.ListIndex
            Height = fMDI.ScaleHeight           ' base size on parent form.
            Left = fMDI.Left
            Exit For                            ' drop out of for loop
        End If                                  ' end if
        Next i                                  ' test the next form
        txtSearch.Text = App.path
        dlbDrive.Drive = "C:\"                  ' set the default drive
        cmdSearch(1).Enabled = False            ' disable function controls
        Check1.Enabled = False
        fLoad = False                           ' reset the on load flag
        'Call ExtensionLoad(cmbTipType.ListIndex) ' load cmbExt
        Call cmbTipType_Click
    Else        ' load defaults.
        chkSetDefault.Value = GetSetting(App.Title, "Search", "Save Search Defaults", 1)
        txtSearch.Text = GetSetting(App.Title, "Search", "Path", txtSearch.Text)
        dlbDrive.Drive = GetSetting(App.Title, "Search", "Drive", dlbDrive.Drive)
        cmbTipType.ListIndex = GetSetting(App.Title, "Search", "Type", cmbTipType.ListIndex)
        cmdSearch(1).Enabled = False                ' disable function controls
        Check1.Enabled = False
        fLoad = False                               ' reset the on load flag
        Call ExtensionLoad(cmbTipType.ListIndex)    ' load cbmExt.
        cmbExt.Text = GetSetting(App.Title, "Search", "Extension", cmbExt.Text)
    End If
    Exit Sub                                    ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in Form Load of frmSearch.")  ' call the error message box
    Resume Next                                     ' continue
End Sub

Private Sub cmbExt_Click()
'*******************************************************
' Purpose: Test to see if the extension of the combobox cmbExt matches
'           the value of the end of the path.
' Assumes:  A valid value in the string
' Inputs:   None               Returns: None
' Comments: This was really a bit of a trick
'          as the value of cmbExt can have varying lengths.
'*******************************************************
On Error GoTo myErrHandler
    
    txtSearch = ExtStrip(txtSearch) ' this routine strips out any extension in the path
Exit Sub                            ' exit the routine
myErrHandler:
    Select Case Err.Number
        Case 5
            Resume Next
        Case Else
            ErrMsgBox (Err.Description & " " & Err.Number & " in cmbExt_Click of frmSearch")  ' call the error message box
            Resume Next                      ' continue
    End Select
End Sub

Private Sub cmbTipType_Change()
'*******************************************************
' Purpose:  Populate the combo cmbExt based on a change in the calling from
' Inputs:   None                Returns: None
' Comments: Drops out during the form load
'*******************************************************
    Call cmbTipType_Click
End Sub

Private Sub cmbTipType_Click()
'*******************************************************
' Purpose:  Populate the combo cmbExt based on a change in the calling from
' Inputs:   None               Returns: None
' Comments: Drops out during the form load
'*******************************************************
On Error GoTo myErrHandler
    If fLoad = True Then                            ' ignore on load
        Exit Sub
    Else
        cmbExt.Clear                                ' clear cbmExt.
        Call ExtensionLoad(cmbTipType.ListIndex)    ' load cbmExt.
        cmbExt.Text = cmbExt.List(0)
        If Len(cmbExt) > 1 Then cmdSearch(1).Caption = "Find " & Right$(cmbExt.Text, Len(cmbExt) - 2) & " files"
        cmdSearch(1).Enabled = False
    End If                                          ' end if
    Exit Sub                                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in cmbTipType_Click of frmSearch")  ' call the error message box
    Resume Next                                     ' continue

End Sub

Private Sub cmbExt_LostFocus()
'*******************************************************
' Purpose: Visual clue so that the search button text matches what is being searched for
' Inputs:   None           Returns: None
'*******************************************************
    If Len(cmbExt) > 1 Then
        cmdSearch(1).Caption = "Find " & Right$(cmbExt.Text, Len(cmbExt) - 2) & " files"
    End If
End Sub

Public Sub AddFunctionSub(ByVal item As String, ByVal fpath As String)
'*******************************************************
' Purpose: Adds result from search to the grid
' Assumes: Private Sub Progress in frmProgress'
' Effects: Adds a string to the MSFlexGrid
' Inputs:  the Funct name & Path    Returns: None
' Comments: Setting multiple properties using With
'*******************************************************
On Error GoTo myErrHandler
    With msgSearch(1)       ' the function/sub grid.
        .AddItem item       ' here is where we add to the grid
        .Row = .Rows - 1    ' the grid title is row 1
        .Col = 1
        .Text = fpath
    End With                ' end with
    Exit Sub                ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & "in AddFunctionSub of frmSearch")  ' call the error message box
    Resume Next             ' continue
End Sub
Private Sub cmdSearch_Click(Index As Integer)
'*******************************************************
' Purpose: Calls frmProgress and passes it which type of search to perform
' Inputs:  the index of the button being pressed    Returns: None
' Comments: This is calling source code that I got from Planet-Source-Code,
'           and although it saved me quite a bit of work, there were some
'           poor coding practices. The original used 3 different progress forms.
'           I combined their functionality into one and just call the one I want based on
'           an index passed before the form is loaded.
'
'    At this time this form can search for files of this type but can currently import Visual basic functions. and HTML documents.
'    To modify this function find section to find functions other than Public Sub, Private Sub, Private Function and Public Function you must modify the following:
'    in frmProgress: Private Sub Progress(), Private Sub ReturnOneFile()
'*******************************************************
On Error GoTo myErrHandler
Dim found As Integer

    Select Case Index
        Case 0                                      ' search for files
            If ValidateSearch = True Then Exit Sub  ' the form isn't filled in properly
            sstSearch.TabVisible(1) = False         ' this sets the File tab
            sstSearch.TabVisible(1) = True          ' to the front
            If optSearch(0).Value = True Then Call TestOptOne ' make sure the path in txtSearch is valid.
            ProgressCancel = False  ' make certain the progress bar isn't active
            InitmsgSearch (0)       ' initialize file flexGrid
            InitmsgSearch (1)       ' initialize function flexGrid
            NbFile = 0              ' set flag values
  '          okMNU = False
            cmdSearch(1).Enabled = False    'disable function controls
            Check1.Enabled = False
            frmProgress.myIndex = 1         ' we're replacing 3 separate progress bars,
            frmProgress.Show 1              ' so we call them by index
            Check1.Enabled = True
  '          okMNU = True
            optFunct(0).Value = True
        Case 1                                  ' search for functions/subroutines
            sstSearch.TabVisible(0) = False     ' this sets the Function tab
            sstSearch.TabVisible(0) = True      ' to the front
            If optFunct(1).Value = True Then    ' looking through only 1 file
                If TestPath = True Then         ' test for proper value in text box
                    InitmsgSearch (1)           ' initialize it
                    frmProgress.myIndex = 0     ' we're replacing 3 separate progress bars,
                    frmProgress.Show 1          ' so we call them by index
        '            okMNU = True
                End If                          ' end the inner if statement
            Else                                ' we're looking for all of the functions in all of the files
                InitmsgSearch (1)               ' initialize it
                frmProgress.myIndex = 0         ' we're replacing 3 separate progress bars,
                frmProgress.Show 1              ' so we call them by index
      '          okMNU = True
            End If                              ' end outer if statement
            found = msgSearch(1).Rows - 1       ' display the results
            If found = 0 Then
                Call StatusMsgDisplay("No " & UCase(Right(cmbExt, Len(fSearch.cmbExt) - 2)) & " Procedures Found", 2)
            ElseIf found = 1 Then
                Call StatusMsgDisplay("1 " & UCase(Right(cmbExt, Len(fSearch.cmbExt) - 2)) & " Procedure Found", 2)
            Else
                Call StatusMsgDisplay(found & " " & UCase(Right(cmbExt, Len(fSearch.cmbExt) - 2)) & " Procedures Found", 2)
            End If                              ' end if
        Case Else

    End Select                                  ' end select
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & "in cmdSearch_Click of frmSearch")  ' call the error message box
    Resume Next             ' continue
End Sub
Private Function ValidateSearch() As Boolean
On Error GoTo VSErr
    If optSearch(0).Value = True Then       ' we search the path
        If Len(txtSearch.Text) = 0 Then
            InputErrBox ("The textbox containing the path information is blank and needs to be filled in.")
            ValidateSearch = True
            Exit Function
        ElseIf cmbExt.Text = "" Or cmbTipType.Text = "" Then
            ValidateSearch = True
            InputErrBox ("One of the combo boxes is blank and needs to be filled in.")
            Exit Function
        Else
            ValidateSearch = False
            Exit Function
        End If
    ElseIf optSearch(1).Value = True Then   ' we search the drive
    
    Else                                    ' we browse
    
    End If
    Exit Function
VSErr:
    ErrMsgBox (Err.Description & " # " & Err.Number & " In ValidateSearch of frmSearch.")
    Resume Next
End Function

Private Function TestPath() As Boolean
'*******************************************************
' Purpose: Performs some simple tests based on the value of txtSearch
' Inputs:  None    Returns: True if successful
' Comments: Not foolproof testing, but it does check for some basic slip ups.
'*******************************************************
On Error GoTo myErrHandler
    If msgSearch(0).Row = 1 Then                ' there is only one file, so take its arguement from the grid
        txtSearch.Text = msgSearch(0).TextMatrix(msgSearch(0).Row, 1) & msgSearch(0).TextMatrix(msgSearch(0).Row, 0)
    ElseIf txtSearch.Text = "" Or Len(txtSearch.Text) < 8 Then  ' something's missing
        InputErrBox ("This doesn't appear to be a valid path." & vbCrLf & "You must select from one of the returned files by clicking on it.")
        sstSearch.TabVisible(1) = False         ' this sets the File tab
        sstSearch.TabVisible(1) = True          ' to the front
    ElseIf Right(txtSearch.Text, 1) = "\" Then  ' don't end on a \ for a function search
       InputErrBox ("This doesn't appear to be a valid path." & vbCrLf & "You must select from one of the returned files by clicking on it.")
        sstSearch.TabVisible(1) = False         ' this sets the File tab
        sstSearch.TabVisible(1) = True          ' to the front
    ElseIf Mid(txtSearch.Text, (Len(txtSearch.Text) - (Len(cmbExt) - 2)), 1) <> "." Then
        InputErrBox ("This doesn't appear to be a valid path." & vbCrLf & "You must select from one of the returned files by clicking on it.")
        sstSearch.TabVisible(1) = False         ' this sets the File tab
        sstSearch.TabVisible(1) = True          ' to the front
    Else
        TestPath = True                         ' hey it worked.
    End If
    Exit Function                               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in TextPath of frmSearch")  ' call the error message box
    Resume Next                                     ' continue
End Function
Private Sub TestOptOne()
'*******************************************************
' Purpose: Performs some simple tests based on the value of txtSearch
' Inputs:  None    Returns: True if successful
' Comments: Again not foolproof testing, but it does check for some basic slip ups.
'           The requirements of finding one tip are slightly different than a more general search.
'*******************************************************
On Error GoTo myErrHandler
    If cmbExt.Text = "" Then cmbExt.Text = cmbExt.Tag
    If txtSearch.Text = "" Or Len(txtSearch.Text) < 3 Then
        txtSearch.Text = App.path
    ElseIf Right(txtSearch.Text, 1) = "\" Then
       txtSearch.Text = Left(txtSearch.Text, Len(txtSearch.Text) - 1)
    ElseIf Mid(txtSearch.Text, (Len(txtSearch.Text) - (Len(cmbExt) - 2)), 1) = "." Then
        txtSearch = StringSplit(txtSearch.Text, 1)
    Else
    End If
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in TestOptOne")  ' call the error message box
    Resume Next             ' continue
End Sub
Private Sub dlbDrive_Change()
'*******************************************************
' Purpose: Change the drive letter in txtSearch to match the dlbDrive value
' Inputs:   None           Returns: None
'*******************************************************
    If Len(txtSearch.Text) > 2 Then
        txtSearch.Text = dlbDrive.Drive & Right(txtSearch.Text, Len(txtSearch.Text) - 2)
    Else: txtSearch.Text = dlbDrive.Drive & "\"
    End If
End Sub


'*******************************************************
' Purpose: Reinitialize the msgSearch grids
' Inputs:   None           Returns: None
'*******************************************************

Private Sub InitmsgSearch(Index As Integer)
On Error GoTo myErrHandler
    Select Case Index
        Case 0  ' file grid
             With msgSearch(Index)  ' set multiple properites
                .Cols = 3
                .Rows = 1
                .Row = 0
                .Col = 0
                .Text = "File Name"
                .Col = 1
                .Text = "Path"
                .Col = 2
                .Text = "Size"
                .Width = sstSearch.Width - (sstSearch.Width * 0.06)
                .ColWidth(0) = (msgSearch(Index).Width * 0.27)
                .ColWidth(1) = (msgSearch(Index).Width * 0.6)
                .ColWidth(2) = (msgSearch(Index).Width * 0.1)
            End With
        Case 1  ' function
            With msgSearch(Index) ' set multiple properites
                .Cols = 2
                .Rows = 1
                .Row = 0
                .Col = 0
                .Text = "Function/Sub"
                .Col = 1
                .Text = "File"
                .Width = sstSearch.Width - (sstSearch.Width * 0.06)
                .ColWidth(0) = (msgSearch(Index).Width * 0.4)
                .ColWidth(1) = (msgSearch(Index).Width * 0.57)
            End With
    End Select
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in InitmsgSearch(" & Index & ") of frmSearch") ' call the error message box
    Resume Next             ' continue
End Sub

Public Sub FindFile(ByVal path As String, ByVal ftype As String)
'*******************************************************
' Purpose: Finds the requested file(s)
' Comment: Function FindFile is From Planet-Source-Code Strongly modified by Carlos 09-10-99
' Inputs:  The Path and file type (extension)   Returns: None
'*******************************************************
    On Error GoTo myErrHandler
    Dim hFile As Long, ts As String, WFD As WIN32_FIND_DATA
    Dim result As Long, sAttempt As String, szPath As String
    Dim strTemp
    If ProgressCancel Then Exit Sub     ' the cancel button was pressed
    frmProgress.ProgressBar1.Value = 1
    If Me.optSearch(2).Value = True Then
        szPath = path & Chr$(0)
    Else
       szPath = path & "*.*" & Chr$(0)
    End If
        
    putFileInPath path, ftype            'Start asking windows for files.
    hFile = FindFirstFile(szPath, WFD)
    Do
      If WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then   'Hey look, we've got a directory!
          ts = StripNull(WFD.cFileName)
          If Not (ts = "." Or ts = "..") Then       'Don't look for hidden or system directories
              If Not (WFD.dwFileAttributes And (FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM)) Then
                  FindFile path & ts & "\", ftype   'Search directory recursively
              End If                                ' close inner if
          End If                                    ' close middle if
        End If                                      ' close outer if
        WFD.cFileName = ""
        result = FindNextFile(hFile, WFD)
          Call StatusMsgDisplay("Searching in: " & path, 1)
        DoEvents
       If frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Max Then frmProgress.ProgressBar1.Value = 1
       frmProgress.ProgressBar1.Value = frmProgress.ProgressBar1.Value + 1
       Loop Until result = 0
    Exit Sub                                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & "In FindFile of frmSearch")  ' call the error message box
    Resume Next                                     ' continue
End Sub

'**********************************************
'* Function putfileinpath is From Planet-Source-Code
'* Modified by Carlos 09-10-99
'***********************************************

Private Sub putFileInPath(ByVal zpath As String, ByVal FileType As String)
On Error GoTo myErrHandler
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
    Exit Sub                                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " In putFileInPath of frmSearch")  ' call the error message box
    Resume Next                                     ' continue
End Sub

Private Sub AddAfile(WFDP As WIN32_FIND_DATA, ByVal path As String)
'*******************************************************
' Purpose: Add files to the msgSearch File grid
' Inputs:   The file name and path     Returns: None
'*******************************************************
On Error GoTo myErrHandler
    NbFile = NbFile + 1
    With msgSearch(0)
       .AddItem Trim(WFDP.cFileName)
       .Row = NbFile
       .Col = 1
       .Text = path
       .Col = 2
       .Text = WFDP.nFileSizeLow / 1000 & " Kb   "
    End With
    Call StatusMsgDisplay(NbFile & " " & UCase(Right(cmbExt, Len(cmbExt) - 2)) & " Files Found.", 2)
    Exit Sub                                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & "In AddAfile of frmSearch")  ' call the error message box
    Resume Next                                     ' continue
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'*******************************************************
' Purpose: Sends message to MDI Status bar.
' Inputs:  The defaults.   Returns: None
'*******************************************************
    Call StatusMsgDisplay("There are no active forms", 2)
    If chkSetDefault.Value = 1 Then
        ' save the path and other information if so desired.
        SaveSetting App.Title, "Search", "Save Search Defaults", chkSetDefault.Value
        SaveSetting App.Title, "Search", "Path", txtSearch.Text
        SaveSetting App.Title, "Search", "Drive", dlbDrive.Drive
        SaveSetting App.Title, "Search", "Extension", cmbExt.Text
        SaveSetting App.Title, "Search", "Type", cmbTipType.ListIndex
    Else
        SaveSetting App.Title, "Search", "Save Search Defaults", chkSetDefault.Value
    End If
End Sub

Private Sub Form_Resize()
'*******************************************************
' Purpose: Handles the repositioning of the controls based on the size of the form.
' Inputs:  None.   Returns: None
'*******************************************************
On Error GoTo myErrHandler
Dim iSize As Integer
    If WindowState = vbNormal Or WindowState = vbMaximized Then
        If Width < 9660 Then Width = 9660
        If Height < 5715 Then Height = 5715
        sstSearch.Width = ScaleWidth * 0.8
        sstSearch.Height = (ScaleHeight - sstSearch.Top) - 150
        msgSearch(0).Height = (sstSearch.Height) - 900
        msgSearch(1).Height = msgSearch(0).Height
        msgSearch(0).Width = sstSearch.Width * 0.9
        msgSearch(1).Width = sstSearch.Width * 0.9
        txtSearch.Width = sstSearch.Width - 1950
        cmbExt.Left = txtSearch.Width + txtSearch.Left + 105
        lblSearch(7).Left = cmbExt.Left
        
        iSize = ScaleWidth - ((3 * sstSearch.Left) + sstSearch.Width)
        cmbTipType.Width = iSize
        lblSearch(9).Width = iSize
        cmdSearch(0).Width = iSize
        cmdSearch(1).Width = iSize
        frmFile.Width = iSize
        Me.frmFunction.Width = iSize
        optSearch(0).Width = iSize * 0.6
        optSearch(0).Left = iSize * 0.1
        optSearch(1).Width = iSize * 0.6
        optSearch(1).Left = iSize * 0.1
        optSearch(2).Width = iSize * 0.6
        optSearch(2).Left = iSize * 0.1
        optFunct(0).Width = iSize * 0.6
        optFunct(1).Width = iSize * 0.6
        Check1.Width = iSize
        chkSetDefault.Width = iSize
        If optFunct(0).Width < 925 Then ' the form is too small to display properly
            optFunct(0).Caption = "All"
            optFunct(1).Caption = "One"
            frmFunction.Caption = "Function"
            optSearch(2).Caption = "Find"
            Check1.Caption = "''_'' Char"
            cmdSearch(1).Caption = "Subs"
            chkSetDefault.Caption = "Settings"
        Else
            optFunct(0).Caption = "All Files"
            optFunct(1).Caption = "Selected"
            optSearch(2).Caption = "Browse"
            cmdSearch(1).Caption = "Functions && Subs"
            frmFunction.Caption = "Function Search"
            Check1.Caption = "Include ''_'' char"
            chkSetDefault.Caption = "Save Settings"
        End If
        iSize = sstSearch.Width + (2 * sstSearch.Left)
        cmbTipType.Left = iSize
        lblSearch(9).Left = iSize
        cmdSearch(0).Left = iSize
        cmdSearch(1).Left = iSize
        frmFile.Left = iSize
        frmFunction.Left = iSize
        Check1.Left = iSize
        chkSetDefault.Left = iSize
    End If
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)  ' call the error message box
    Resume Next             ' continue
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If g_blnUnload = True Then Exit Sub
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub

Private Sub msgSearch_Click(Index As Integer)
'*******************************************************
' Purpose: Puts row value from grid (0) into txtSearch, else fires right mouse event.
' Inputs:  the control's index   Returns: None
'*******************************************************
    Select Case Index
        Case 0  ' files
            txtSearch.Text = msgSearch(Index).TextMatrix(msgSearch(Index).Row, 1) & msgSearch(Index).TextMatrix(msgSearch(Index).Row, 0)
            optFunct(1).Value = True
        Case 1  ' functions
          '  Call msgSearch_MouseUp(Index, 2)
    End Select
End Sub

Private Sub msgSearch_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'*******************************************************
' Purpose: Sends message to MDI Status bar.
' Inputs:  The control index, mouse button    Returns: None
'*******************************************************
    On Error GoTo myErrHandler
    Select Case Index
    Case 0  ' files
        If Button = 2 Then
            fMDI.mnuSFOpen.Caption = "Open " & msgSearch(Index).TextMatrix(msgSearch(Index).Row, 0)
            PopupMenu fMDI.mnuSearchFiles
        End If
    Case 1  ' functions
        If msgSearch(Index).Row > 0 Then
            If Button = 2 Then

                PopupMenu fMDI.mnuImport    ' Show import menu

            End If
        Else                                ' they're on the title bar.
            If Button = 2 Then
                PopupMenu fMDI.mnuHelp      ' Show the help menu
            End If                          ' close inner if
        End If                              ' outer inner if
    End Select                              ' end select
    Exit Sub                                ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)  ' call the error message box
    Resume Next             ' continue
End Sub

Private Sub msgSearch_RowColChange(Index As Integer)
On Error GoTo myErrHandler
    If msgSearch(0).Rows > 1 Then
        If cmbExt.Text = "*.bas" Or cmbExt.Text = "*.BAS" Or cmbExt.Text = "*.cls" Or cmbExt.Text = "*.CLS" Or cmbExt.Text = "*.frm" Or cmbExt.Text = "*.FRM" Or cmbExt.Text = "*.FRX" Or cmbExt.Text = "*.frx" Then
            cmdSearch(1).Enabled = True
        Else: cmdSearch(1).Enabled = False
        End If
    Else: cmdSearch(1).Enabled = False
    End If
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in msgSearch_RowColChange(" & Index & ") of frmSearch")  ' call the error message box
    Resume Next                                     ' continue
End Sub

Private Sub optFunct_Click(Index As Integer)
'*******************************************************
' Purpose: Find either all or one of the functions
' Inputs:  the index         Returns: None
'*******************************************************
On Error GoTo myErrHandler
    Select Case Index
    Case 0                              ' find all functions contained in all the files in msgSearch(0)
        txtSearch.Tag = txtSearch.Text  ' set the text to the tag
        txtSearch.Text = ""             ' clear the text box
    Case 1                              ' find based on one selected file
        If txtSearch.Text = "" Then     ' there must only be one function to select and the user didn't put it into the textbox txtSearch.
            msgSearch(0).Row = 1
            txtSearch.Text = msgSearch(0).TextMatrix(msgSearch(0).Row, 1) & msgSearch(0).TextMatrix(msgSearch(0).Row, 0)
            If cmbExt.Text <> "" Then cmbExt.Tag = cmbExt.Text
        End If
    End Select
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)  ' call the error message box
    Resume Next             ' continue
End Sub

Private Sub optFunct_DblClick(Index As Integer)
'*******************************************************
' Purpose: Refire the click event
' Inputs:  The option index   Returns: None
'*******************************************************
    optFunct_Click (Index)
End Sub

Private Sub optSearch_Click(Index As Integer)
'*******************************************************
' Purpose: Search an increasingly narrow scope depending on the selection of the option box
' Inputs:  index         Returns: None
'*******************************************************
On Error GoTo OptSearchErr
    Select Case Index
        Case 0  ' search by path
            Call TestOptOne
        Case 1  ' search by drive
            If cmbExt.Text = "" Then cmbExt.Text = cmbExt.Tag
            txtSearch.Text = dlbDrive.Drive
        Case 2  ' browse
            If cmbExt.Text <> "" Then cmbExt.Tag = cmbExt.Text
            If txtSearch.Text = Me.dlbDrive.Drive Then
                txtSearch.Text = txtSearch.Text & "\" & cmbExt.Text ' App.path
            End If
            With dlgSearch  ' set the properties of the common dialog control
                .FileName = txtSearch.Text
                .DialogTitle = App.Title
                .Filter = Right(cmbExt, Len(cmbExt) - 2)
                .CancelError = True
                .ShowOpen
            End With

            If Len(dlgSearch.FileName) Then                     ' a file was found & returned (a return null produces an error)
                txtSearch = StringSplit(dlgSearch.FileName, 1)  ' so we put the returned name into the text box txtSearch.
            End If
            cmdSearch_Click (0)                                 ' fire the click event and launch the search based on the returned path.
    End Select                                                  ' end the select
    Exit Sub                                                    ' exit the function
OptSearchErr:
    Select Case Err.Number
    Case 32755 ' common dialog error, when user cancel's browse.
        optSearch(0).SetFocus
        optSearch(0).Value = True
        Exit Sub
    Case Else
        ErrMsgBox (Err.Description & " " & Err.Number & " in optSearch of frmSearch.")   ' call the error message box
        Resume Next
    End Select
    
End Sub

Private Sub optSearch_DblClick(Index As Integer)
'*******************************************************
' Purpose:  Fires of the click event.
' Inputs:   the index of the array.   Returns: None.
'*******************************************************
    optSearch_Click (Index)
End Sub

Private Sub txtSearch_Change()
'*******************************************************
' Purpose:  preforms preliminary testing of the text box.
' Comment:  Not very robust. It sets the extension of any path to equal what is in the cmbExt.
' Inputs:   None.   Returns: None.
'*******************************************************
On Error GoTo myErrHandler
    If Len(txtSearch) > 3 Then
        If Mid(txtSearch.Text, (Len(txtSearch.Text) - 3), 1) = "." Then
            If Right(txtSearch.Text, 3) <> Right(cmbExt, Len(cmbExt) - 2) Then
                cmbExt.Text = "*." & Right(txtSearch.Text, 3)
            End If
        End If
    End If
    Exit Sub               ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in txtSearch_Change of frmSearch.")   ' call the error message box
    Resume Next             ' continue
End Sub
Private Sub ExtensionLoad(comboIndex As Integer)
On Error GoTo myErrHandler
'*******************************************************
' Purpose:  Load the extension into the combo box based on the
'           value of the index of cmbTipType
' Assumes:  A valid value loaded into cmbTipType
' Inputs:   cmbTipType.cmbTipType.ListIndex    Returns: None
' Comments: In reality, this is not the best way to handle
'           this sort of thing, but I have hard coded the values
'           to here to reduce overheard, and show how to load values at run time into a combobox
'*******************************************************
If fLoad = True Then Exit Sub

    Call ComboboxLoad(cmbExt, ExtLoad(comboIndex))
    Exit Sub                        ' exit the routine
myErrHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & "in ExtensionLoad of frmSearch")  ' call the error message box
    Resume Next                      ' continue
End Sub
'
