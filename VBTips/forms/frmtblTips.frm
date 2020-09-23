VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmtblTips 
   Caption         =   "Source Code Toolbox"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   9840
   Icon            =   "frmtblTips.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   9840
   Tag             =   "tblTips"
   Begin VB.CommandButton cmdTip 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   8940
      TabIndex        =   37
      ToolTipText     =   "To commit a change. (Ctrl + U)"
      Top             =   8040
      Width           =   1342
   End
   Begin VB.CommandButton cmdTip 
      Caption         =   "Delete"
      Height          =   255
      Index           =   3
      Left            =   7605
      TabIndex        =   36
      Tag             =   "&Delete"
      ToolTipText     =   "To delete the current tip."
      Top             =   8040
      Width           =   1342
   End
   Begin VB.CommandButton cmdTip 
      Caption         =   "&Add"
      Height          =   255
      Index           =   2
      Left            =   6270
      TabIndex        =   35
      Tag             =   "&Add"
      ToolTipText     =   "To add a new tip (Ctrl + A)"
      Top             =   8040
      Width           =   1342
   End
   Begin VB.CommandButton cmdTip 
      Caption         =   "&Most Recent"
      Height          =   255
      Index           =   1
      Left            =   4935
      TabIndex        =   21
      Tag             =   "&Add"
      ToolTipText     =   "View all tips sorted by most recent entry"
      Top             =   8040
      Width           =   1342
   End
   Begin VB.CommandButton cmdTip 
      Caption         =   "View Ty&pe"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   19
      Tag             =   "&Add"
      ToolTipText     =   "View one types of tips. "
      Top             =   8040
      Width           =   1342
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&<"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   16
      ToolTipText     =   "Move backward one tip. (Ctrl+Shift+<)"
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">&|"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   18
      ToolTipText     =   "Move to last record. (Ctrl + Shift + |)"
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "&>"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   17
      ToolTipText     =   "Move forward one tip. (Ctrl+Shift+>)"
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "|<"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   15
      ToolTipText     =   "Move to first record."
      Top             =   8040
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "vbTips.mdb"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   13573
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Notes"
      TabPicture(0)   =   "frmtblTips.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtFields(3)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Code"
      TabPicture(1)   =   "frmtblTips.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFields(4)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Details"
      TabPicture(2)   =   "frmtblTips.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblDetail(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblDetail(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblDetail(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblDetail(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblTitleWarn"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblDetail(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtFields(0)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmbTipType"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtFields(1)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame1"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtFields(2)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdFind"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmbTipSubType"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      Begin VB.ComboBox cmbTipSubType 
         Height          =   315
         Left            =   4080
         Sorted          =   -1  'True
         TabIndex        =   8
         Text            =   "Category"
         ToolTipText     =   "Select/modify the source code subtype here."
         Top             =   1440
         Width           =   1830
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         Height          =   300
         Left            =   4920
         TabIndex        =   11
         Tag             =   "&Add"
         ToolTipText     =   "Perform a search by entering a keyword and pressing this button."
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Type your query here. Double click the Find label to widen or narrow your search."
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Frame Frame1 
         Caption         =   "View"
         Height          =   1335
         Left            =   1320
         TabIndex        =   12
         ToolTipText     =   "Select any of the options to view only that type of source code. Double click to view all tips."
         Top             =   2160
         Visible         =   0   'False
         Width           =   4575
         Begin VB.OptionButton optViewType 
            Caption         =   "Access"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   22
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   240
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "C, C++"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   23
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   480
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "CGI"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   24
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   720
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "HTML"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   25
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   960
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "Java"
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   26
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   240
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "Javascript"
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   27
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   480
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "MSOffice"
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   28
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   720
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "SQL"
            Height          =   255
            Index           =   7
            Left            =   1680
            TabIndex        =   29
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   960
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "VB5"
            Height          =   255
            Index           =   8
            Left            =   3120
            TabIndex        =   30
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   240
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "VB6"
            Height          =   255
            Index           =   9
            Left            =   3120
            TabIndex        =   31
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   480
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "Visual Basic"
            Height          =   255
            Index           =   10
            Left            =   3120
            TabIndex        =   32
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   720
            Width           =   1148
         End
         Begin VB.OptionButton optViewType 
            Caption         =   "Other"
            Height          =   255
            Index           =   11
            Left            =   3120
            TabIndex        =   33
            ToolTipText     =   "Use the combo boxes on the toolbar to view only one Language tip subtype."
            Top             =   960
            Width           =   1148
         End
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   250
         TabIndex        =   5
         ToolTipText     =   "You may store a web site where the tip was found or keywords to better index this tip in this field."
         Top             =   1080
         Width           =   4575
      End
      Begin VB.ComboBox cmbTipType 
         Height          =   315
         ItemData        =   "frmtblTips.frx":0496
         Left            =   1320
         List            =   "frmtblTips.frx":0498
         Sorted          =   -1  'True
         TabIndex        =   7
         Text            =   "Language"
         ToolTipText     =   "Select the source code type here."
         Top             =   1440
         Width           =   1830
      End
      Begin VB.TextBox txtFields 
         DataField       =   "title"
         Height          =   285
         Index           =   0
         Left            =   1320
         MaxLength       =   250
         TabIndex        =   3
         ToolTipText     =   "Enter Unique Source Code Tip Titles Here"
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Tip"
         DataSource      =   "Data1"
         Height          =   5955
         Index           =   4
         Left            =   -74520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Only source code goes here."
         Top             =   720
         Width           =   5535
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Notes"
         DataSource      =   "Data1"
         Height          =   5715
         Index           =   3
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "All tip notes go here."
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         Caption         =   "SubType"
         Height          =   195
         Index           =   4
         Left            =   3360
         TabIndex        =   34
         ToolTipText     =   "Select/modify the source code subtype here."
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label lblTitleWarn 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1320
         TabIndex        =   20
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T&ype"
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   6
         ToolTipText     =   "Select/modify the source code type here."
         Top             =   1470
         Width           =   765
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Index"
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   4
         ToolTipText     =   "Enter web site where tip was found. Or Keywords to better index a search for this tip"
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Title"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   2
         ToolTipText     =   "Enter title of tips here. Duplicate titles are incremented."
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblDetail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Find Al&l"
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   9
         ToolTipText     =   "Click on this label to toggle between a quick and total search!"
         Top             =   1800
         Width           =   765
      End
   End
   Begin VB.ListBox lstTitle 
      Height          =   7665
      ItemData        =   "frmtblTips.frx":049A
      Left            =   120
      List            =   "frmtblTips.frx":049C
      TabIndex        =   0
      ToolTipText     =   "Click to move to a selected tip."
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmtblTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************************
' Purpose:  This is the main form for this project
' Assumes:  modConstants, modControls, modFormEffects, modFormUtilities, modstartup, modDBEngine
' Project:  Source_Code_Tool_Box
' Depends:  Is called from sub Main in modstartup
' Comment:  It is called by ftblTips in modConstants
' Authors:   James R. Fleming
'*******************************************************
'*****************************************************
'
'Purpose: Declare all module level variable
'
'*****************************************************
Dim bForm As Boolean        ' flag for the form was resized
Dim bFirst As Boolean       ' flag for the form is first Loaded
Dim bSubList As Boolean     ' flag for if a partial list of tips if being viewed
Dim moRS As Recordset       ' module level recordset
Dim strStatMsg As String    ' string for holding the status message
Dim iFKey As Integer        ' this is for holding the value of the primary key of the active record (it is also the foreign key in tblCode, tblNotes
Dim bNoteLen As Boolean     ' this is for a flag that there exists a child record in tblNotes
Dim bCodeLen As Boolean     ' this is for a flag that there exists a child record in tblCode
'**********************************************************************************************************
'
'Purpose:  All of the form subs
'
'**********************************************************************************************************

'*****************************************************
' All Form subs and events
'*****************************************************
Private Sub Form_Initialize()
' Assumes: modFormUtilities
    Call StatusMsgDisplay("The Tips form is initialized...", 2)
End Sub

Private Sub Form_Load()
'*****************************************************
' Purpose:  the form Load event is a driver for all events
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler

    bFirst = True                   ' Borrow this flag to skip the update
    Visible = False
    strStatMsg = LoadAllRecords(lstTitle, cmbTipType, fMDI.cmbTipType)              ' handle load of all data
    Set moRS = BuildRecordSet       ' load the other info
    If moRS.RecordCount > 0 Then Call ListNavigate(0, lstTitle, moRS, bSubList)     ' show the first record
    WindowState = vbMaximized    ' make it big
    Visible = True
    bFirst = False                  ' put the flag back to the default
    Exit Sub                        ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in Form_Load of frmtblTips")
    Resume Next
End Sub
Private Sub Form_Activate()
    Call StatusFlip("The Tips Form is Active...", strStatMsg, 2)
End Sub

Private Sub Form_Resize()
'*****************************************************
' Purpose:  calls the resize sub
' Inputs:   None    ' Returns:  None
'*****************************************************
 On Error GoTo myErrorHandler
    Static blnHeight As Boolean ' two flags to prevent this sub from refiring itself
    Static blnWidth As Boolean
    If WindowState <> vbMinimized And fMDI.WindowState <> vbMinimized Then
        If Height < 5505 Then   ' too short
            blnHeight = True    ' set a flag to prevent recursive calling of resizing.
            Height = 5505
        End If
        If Width < 9990 Then    ' too thin
            blnWidth = True     ' set a flag to prevent recursive calling of resizing.
            Width = 9990
        End If
    Else
        Exit Sub            ' the quick way out
    End If
    If blnHeight = True Then
       blnHeight = False    ' reset the flag
       Exit Sub             ' keep from going through form resizing and positioning n times.
    End If
    If blnWidth = True Then
        blnWidth = False    ' reset the flag
        Exit Sub            ' keep from going through form resizing and positioning n times.
    End If
    ' all of the form resizing and positioning.
    Call SSTabResize(Me)        ' resize the tab control
    Call NaviButtonsMove(cmdNav, (ScaleHeight - 360), (lstTitle.Width / 4)) ' adjust the buttons on the main form
    Call SortButtonsMove(cmdNav(0).Top, SSTab1.Left, (SSTab1.Width / 5), cmdTip())
    Call TabControlsMove(Me)    ' adjust the tab controls
    bForm = True
    Exit Sub                    ' exit the routine
myErrorHandler:
    Select Case Err.Number
    Case 384
        Resume Next
    Case Else
        ErrMsgBox (Err.Description & " " & Err.Number & " in Form_Resize")
        Resume Next
    End Select
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
' Purpose:  Put a simple menu on the form
' Inputs:   button, shift and mouse position   ' Returns:  None
'*****************************************************
    If Button = 2 Then              ' the right mouse click
        PopupMenu fMDI.mnuEdit      ' a sample menu
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*****************************************************
' Purpose:  That's all folks!
' Inputs:   None     Returns:  None
'*****************************************************
    If g_blnUnload = True Then Exit Sub
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub
'*****************************************************
' All button and control events
'*****************************************************
Private Sub cmbTipSubType_GotFocus()
'*****************************************************
' Purpose: To set a flag at any change in the value of the combo box
' Assumes: modControls
' Inputs:   None    ' Returns:  None
'*****************************************************
    cmbTipSubType.Tag = cmbTipSubType.Text      ' set the flag
    Call ControlGotFocus(lblDetail(4))          ' switch off visual focus
End Sub

Private Sub cmbTipSubType_LostFocus()
'*****************************************************
' Purpose:  Test the value of the flag set during got focus
' Assumes: modControls
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
    Dim strTemp As String
    strTemp = cmbTipSubType.Text
    If Trim$(cmbTipSubType.Text) = "" Then cmbTipSubType.Text = cmbTipSubType.Tag
    If cmbTipSubType.Tag <> cmbTipSubType.Text Then
        Call SubTypeLoad(SubTypeQry(cmbTipType.ListIndex), cmbTipSubType)
        Call DataHasChanged                     ' tests the value, toggle the buttons off
        cmbTipSubType.Tag = cmbTipSubType.Text  ' reset the flag
    End If                                      ' end if statement
    Call ControlLostFocus(lblDetail(4))         ' switch off visual focus
    Exit Sub                                    ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in cmbTipSubType_LostFocus() of frmtblTips")
    Resume Next
End Sub

Private Sub cmbTipType_GotFocus()
'*****************************************************
' Purpose: To set a flag at any change in the value of the combo box
' Assumes: modControls
' Inputs:   None    ' Returns:  None
'*****************************************************
    cmbTipType.Tag = cmbTipType.Text    ' set the flag
    Call ControlGotFocus(lblDetail(3))  ' light up the Label
End Sub

Private Sub cmbTipType_LostFocus()
'*****************************************************
' Purpose: Test the value of the flag set during got focus
' Assumes: modControls
' Inputs:  None    ' Returns:  None
'*****************************************************
Dim i As Integer
    If Trim$(cmbTipType.Text) = "" Then cmbTipType.Text = cmbTipType.Tag
    If cmbTipType.Tag <> cmbTipType.Text Then
        Call DualSubTypeLoad(SubTypeQry(cmbTipType.ListIndex), cmbTipSubType, fMDI.cmbTipSubType)
        Call DataHasChanged                 ' tests the value, toggle the buttons off
        cmbTipType.Tag = cmbTipType.Text
    End If
    Call ControlLostFocus(lblDetail(3))     ' switch off visual focus
    cmbTipType.Tag = cmbTipType.ListIndex   '  reset the flag
    
End Sub

Private Sub cmdFind_Click()
'*****************************************************
' Purpose:  Create the SQL string based on the content of
'           the txtFields(2).text. Passes the string to the appropriate sub
' Inputs:   None  ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler

    If Caption = "Find" And txtFields(2).Text = "" Then    ' we wish to search for a tip  ' test for a value. If blank reset the listbox
        Call LoadAllRecords(lstTitle, cmbTipType, fMDI.cmbTipType)         ' reset the tips to start up status.
        Call FirstRecordShow        ' show the first record
        txtFields(2).SetFocus       ' search text box
        Exit Sub                    ' exit the routine
    Else
        Call qrySearch(cmdFind, txtFields(2), lblDetail(2), lstTitle)
    End If                          ' close inner if statement
    FirstRecordShow                 ' populate the list
    Exit Sub                        ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in cmdFind_Click() of frmtblTips")
    Resume Next
End Sub
Private Sub cmdNav_Click(Index As Integer)
'*****************************************************
' Purpose: Navigate through the records of the list
' Comment:  We have to test to make sure we're not on the first or last record
'           We have to handle moving through the sublists differently than
'           the entire collection
' Inputs:   the index of the control.   Returns:  None
' assume:   All the nav buttons are an array.
'*****************************************************
On Error GoTo cmdNavErr
    Call ListNavigate(Index, lstTitle, moRS, bSubList)     ' show the first record
    If lblTitleWarn.Caption <> "" Then lblTitleWarn.Caption = "" ' hide the title warning
    Exit Sub
cmdNavErr:
    ErrMsgBox (Err.Description & " " & Err.Number & ". In cmdNav_Click of frmTblTips")
    Resume Next
End Sub

Public Sub cmdTip_Click(Index As Integer)
'*****************************************************
' Purpose:  Takes the index of the value in cmbTipType as an arguement
'            sets focus to the option button with the matching index
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
Select Case Index
    Case 0              ' View Ty&pe
        Call RecordViewType
    Case 1              ' view by date/subtype
        Call RecordSort
    Case 2              ' add
        Call RecordAdd
    Case 3
        Call RecordDelete
    Case 4
        Call RecordUpdate
    Case Else
End Select
    Exit Sub            ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in cmdTip_Click of frmTblTips")
    Resume Next
End Sub

Public Sub RecordUpdate()
'*****************************************************
' Purpose:  Here we update the record to the recordset
' Inputs:   None    ' Returns:  None
' comment:  This command has 3 parts: validate the data, update the record, reset the form
'*****************************************************
    On Error GoTo Update_ERR
    Dim strProperTitle As String
    Dim bNoteKill As Boolean, bCodeKill As Boolean      ' the deletes should occur after the updates
    ' validate the input and the title which must be unique
    If FormCheck(Me) = True Then Exit Sub ' make sure that the tip is filled in  fall out of the routine
    strProperTitle = TitleCaps(txtFields(0).Text)       ' Put title in proper case
    strProperTitle = TitleValidate(strProperTitle, lblTitleWarn)    'This may be redundant JAMES test for a valid title
    If cmdTip(2).Tag = "New" Then                       ' prepare to add/update the data
        strProperTitle = TitleReTitle(strProperTitle)   ' I need to pass an actual string - Open quote
        strProperTitle = TestTitle(strProperTitle)      ' test for duplicate titles
        moRS.AddNew                                     ' prepare to add to the recordset
    Else                                                ' else it's an update we are modifying an existing record
        moRS.Edit                                       ' prepare to edit the recordset
    End If                                              ' end if
    With moRS                                           ' begin recordset modification
        !strTitle = Trim$(strProperTitle)               ' add the title
        !strIndex = Trim$(txtFields(1))                 ' add the web site if found on line
        !strSTTitle = Trim$(cmbTipSubType)              ' add the source if other than a web site
        !intTypeID = Trim$(cmbTipType.ListIndex)        ' This stores the id but not the id label
        !datTipDate = Now()                             ' timestamp my entry (timestamps are not visible on form)
        If Len(txtFields(3)) Then                       ' there are new notes so we generate a new record
            !memNotes = Trim$(txtFields(3))             ' add the notes (if blank skip it as these are separate tables: keep size down)
            !lngNoteTipsFK = !lngTblTipsID              ' add the foreign key
        ElseIf Len(txtFields(3)) = 0 And bNoteLen = True Then
            bNoteKill = True                            ' there was a record and we are deleting it
            bNoteLen = False                            ' reset the flag
        End If                                          ' else we're not adding anything and there was no record
        If Len(txtFields(4)) Then                       ' there is new source code so we generate a new record
            !memCode = Trim$(txtFields(4))              ' add the code  (if blank skip it as these are separate tables: keep size down)
            !lngCodeTipsFK = !lngTblTipsID              ' add the foreign key
        ElseIf Len(txtFields(4)) = 0 And bCodeLen = True Then
            bCodeKill = True                            ' there was a record and we are deleting it
            bCodeLen = False                            ' reset the flag
        End If                                          ' else we're not adding anything and there was no record
        iFKey = !lngTblTipsID                           ' assign the foreign key to a variable ( this may not need to be modLevel)
    End With                                            ' end the record set modification
    moRS.Update                                         ' update the recordset
    If bNoteKill = True Then Call KillNote(iFKey)       ' delete the note record in tblNotes
    If bCodeKill = True Then Call KillCode(iFKey)       ' delete the code record in tblCode
    Set moRS = BuildRecordSet 'rebuild the recordset to reorder with new/edited record
    Call StatusMsgDisplay(TipCount(ListPopulate(qryList, lstTitle)), 2) ' loads the list box, build the string, display it
    moRS.FindFirst "strTitle='" & strProperTitle & "'"      'find the new position of the new record in the list and highlight it
 '   If moRS.AbsolutePosition < 38 Then                      ' set index to the beginning
    lstTitle.ListIndex = moRS.AbsolutePosition               ' however if the list is longer than the list box
'    Else                                                    ' it always seems to be one 'short' of where it belongs
'        lstTitle.ListIndex = moRS.AbsolutePosition + 1      ' so we bump it forward one place
'    End If                                                  ' end if
    Call ToggleButtons                                       ' reset the buttons
    Exit Sub                                                ' exit the routine
Update_ERR:
    ErrMsgBox (Err.Description & " " & Err.Number & " in RecordUpdate of frmTblTips")
    Resume Next
End Sub
Private Sub RecordViewType()
'*****************************************************
' Purpose:  prepares the form for data entry
' Inputs: None  ' Returns:   None
'*****************************************************
On Error GoTo myErrorHandler
    If cmdTip(0).Caption = "View Ty&pe" Then
        ' make the frame visible, size it and dynamically title the options
        Frame1.Visible = True   ' show the container for the option buttons
        With fMDI
            .cmbTipSubType.Visible = True
            .cmbTipSubType.Text = Me.cmbTipSubType.Text
            .cmbTipSubType.ListIndex = cmbTipSubType.ListIndex
            .cmbTipType.ListIndex = cmbTipType.ListIndex
            .cmbTipType.Text = cmbTipType.Text
        End With
        Call optTypeName(qryOptName, optViewType())
        Call FrameSize(Frame1, ((cmdFind.Left + cmdFind.Width) - txtFields(0).Left), txtFields(0).Left, optViewType)        ' resize the frame based on the forms width
        Call AlignOpt(optViewType())
        ' change the button captions, sizes, tooltips
        cmdTip(0).Caption = "View all"                      ' set button caption
        cmdTip(1).Caption = "Sort SubType"                  ' set button caption
        cmdTip(0).ToolTipText = "Press here to view all types of tips."
        cmdTip(1).ToolTipText = "Show " & cmbTipType.ToolTipText & " sorted by tip SubType."
        Call optViewType_GotFocus(cmbTipType.ListIndex)     ' light up the control
        optViewType(cmbTipType.ListIndex).SetFocus          ' set focus
        OptionSelect (cmbTipType.ListIndex)                 ' This is to view a sublist of all the available tips
    ElseIf cmdTip(0).Caption = "View all" Then
        fMDI.cmdTlbrST.Caption = "Search"
        Call StatusMsgDisplay(TipCount(ListPopulate(qryList, lstTitle)), 2)
        Frame1.Visible = False
        cmdTip(0).Caption = "View Ty&pe"
        cmdTip(0).ToolTipText = "Press here to view one type of tips."
        cmdTip(1).ToolTipText = "View all tips sorted by most recent"
        cmdTip(1).Caption = "&Most Recent"
        lstTitle.ToolTipText = "Click to move to a selected tip."
        bSubList = False
        Call FirstRecordShow                    'set it to the first record in the list
    ElseIf cmdTip(0).Caption = "&Update" Then
        Call RecordUpdate
        cmdTip(0).Caption = "View Ty&pe"
        cmdTip(1).Caption = "&Most Recent"
    End If
    Exit Sub                        ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in RecordViewType of frmTblTips")
    Resume Next
End Sub

Private Sub RecordSort()
'*****************************************************
' Purpose:  Rearrange the order of the list box in order of data Last entry first.
' Inputs:   None    ' Returns:  None
'*****************************************************
  On Error GoTo myErrorHandler
   If cmdTip(1).Caption = "&Most Recent" Then
        Call StatusMsgDisplay(TipCount(ListPopulate(qryDateSort, lstTitle)), 2)
        Call FirstRecordShow
        cmdTip(1).Caption = "&Alphabetical"
        cmdTip(1).ToolTipText = "Press here to view all tips in alphabetical order."
    ElseIf cmdTip(1).Caption = "&Alphabetical" Then
        Call StatusMsgDisplay(TipCount(ListPopulate(qryList, lstTitle)), 2)  ' loads the list box
        Call FirstRecordShow
        cmdTip(1).Caption = "&Most Recent"
        cmdTip(1).ToolTipText = "View all tips sorted by most recent entry."
    ElseIf cmdTip(1).Caption = "Sort SubType" Then
         Call StatusMsgDisplay(TipCount(ListPopulate(ViewBySubType(cmbTipType.ListIndex), lstTitle)), 2)    ' creates a sublist of an access datatable
         Call FirstRecordShow
         cmdTip(1).Caption = "&Alphabetical"
    Else
        Call CancelUpdate
    End If
    Exit Sub                ' exit the routine
myErrorHandler:
        ErrMsgBox (Err.Description & " " & Err.Number & " in RecordSort of frmTblTips")
        Resume Next
    End Sub
Private Sub RecordAdd()
'*****************************************************
' Purpose:  prepares the form for data entry
' Inputs: None  ' Returns:   None
'*****************************************************
On Error GoTo myErrorHandler
    cmdTip(2).Tag = "New"           ' set the flag
    cmdTip(0).Enabled = False       ' disable the control
    cmdTip(1).Enabled = False       ' disable the control
    DataHasChanged                  ' tests the value, toggle the buttons off
    Call FormClear(txtFields)       ' clear the Details tab text and combo controls
    txtFields(0).SetFocus           ' set focus on the first text box
    moRS.MoveLast                   ' move to the end of the record set
    Exit Sub                        ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub RecordDelete()
'*****************************************************
' Purpose:  Warn the user they are about to delet a record
'           Process the result of their action
' Inputs:   None    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
    Dim intResponse As Integer
    If cmdTip(3).Caption = "Delete" Then   'allow user to undo action
        intResponse = YesNo("Delete the current tip?")
        If intResponse = vbYes Then
            If lstTitle.ListIndex > 0 Then
                intResponse = (lstTitle.ListIndex) - 1 ' set the index to the tip id being deleted.
            Else
                intResponse = 0 ' reset the index
            End If              ' end the nested if
            moRS.Delete         ' delete the record set
            Call StatusMsgDisplay(TipCount(ListPopulate(qryList, lstTitle)), 2) ' reload the list w/o the deleted record
            Call cmdNav_Click(intResponse) ' set to first item in list
        End If                  ' end the inner if statement
    Else                        ' it said "C&ancel" and was an edit or add
        Call CancelUpdate       ' skip it!
    End If                      ' end outer if
    Exit Sub                    ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in RecordDelete of frmTblTips")
    Resume Next
End Sub

Private Sub lstTitle_DblClick()
'*****************************************************
' Purpose:  hide  lblTitleWarn if visible once the active record is moved
' Inputs:   None    ' Returns:  None
'*****************************************************
    If lblTitleWarn.Caption <> "" Then lblTitleWarn.Caption = ""
End Sub

Private Sub lstTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
' Purpose:  Put a simple menu on the list
' Inputs:   button, shift and mouse position   ' Returns:  None
'*****************************************************
    If Button = 2 Then
        PopupMenu fMDI.mnuHelp
    End If
End Sub

Private Sub SSTab1_GotFocus()
    Call SSTab1Focus(SSTab1, txtFields)
End Sub

Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'*****************************************************
' Purpose:  Put a simple menu on the tab control
' Inputs:   button, shift and mouse position   ' Returns:  None
'*****************************************************
    If Button = 2 Then
        PopupMenu fMDI.mnuHelp
    End If
End Sub

Private Sub txtFields_Change(Index As Integer)
'*****************************************************
' Purpose:  All of the text & combobox handling routines
' Inputs:   the control index   ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler                    ' set the error handler
    If bFirst = False Then                      ' skips change event on form load
        If Index <> 2 Then                      ' the query text box
            Call DataHasChanged                 ' tests the value, toggle the buttons off
            cmdFind.Caption = "Find"            ' reset the query button
            If cmdTip(2).Tag = "New" Or cmdTip(4).Tag = "Update" Then   ' set the flag Then the title is too big
                lblTitleWarn.Caption = TxtLenWarn(txtFields(0), 42)
            End If
        Else                                    ' it is the query field
            If txtFields(Index).Text = "" Then  ' assume they want to view all
                cmdFind.Caption = "Find"        ' reset the query button
                Call cmdFind_Click              ' Create the SQL string
            End If                              ' close inner if
        End If                                  ' close middle if
    End If                                      ' close outer if
    Exit Sub                                    ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in txtFields_Change of frmTblTips")
    Resume Next
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
'*****************************************************
' Purpose:  Call the label and textbox onfocus routines
' Assumes:  modControls
' Inputs:   the Index of the current control    ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
    txtFields(Index).Tag = txtFields(Index).Text ' set the flag
    Call ControlGotFocus(lblDetail(Index))  ' light up the control
    Call TextboxSelect(txtFields(Index))    ' Highlight any text
    If Index = 2 Then
        If txtFields(Index).Text = "" Or cmdFind.Caption = "Clear" Then
            cmdFind.Caption = "Find"            ' reset the label caption
        End If
    End If
    SSTab1.ForeColor = &H800000 ' Give the tab control added focus
    Exit Sub                ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in txtFields_GotFocus of frmTblTips")
    Resume Next
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 18 Then
        If cmdTip(2).Tag = "New" And Index = 3 Then ' test the flag
            SSTab1.TabVisible(0) = False            ' this sets the File tab
            SSTab1.TabVisible(0) = True             ' to the front
            txtFields(4).SetFocus
        ElseIf cmdTip(2).Tag = "New" And Index = 4 Then ' test the flag
            SSTab1.TabVisible(0) = False            ' this sets the File tab
            SSTab1.TabVisible(1) = False            ' this sets the File tab
            SSTab1.TabVisible(0) = True             ' to the front
            SSTab1.TabVisible(1) = True             ' to the front
        End If
    End If
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = vbTab Then
        If cmdTip(2).Tag = "New" And Index = 3 Then ' test the flag
            SSTab1.TabVisible(0) = False            ' this sets the File tab
            SSTab1.TabVisible(0) = True             ' to the front
            txtFields(4).SetFocus
        ElseIf cmdTip(2).Tag = "New" And Index = 4 Then ' test the flag
            SSTab1.TabVisible(0) = False            ' this sets the File tab
            SSTab1.TabVisible(1) = False            ' this sets the File tab
            SSTab1.TabVisible(0) = True             ' to the front
            SSTab1.TabVisible(1) = True             ' to the front
        End If
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
'*****************************************************
' Purpose:  Call the label and textbox onfocus routines
' Inputs:   the Index of the current control    ' Returns:  None
' Comments: Since the web address text box does change without actually getting a value, we want to
'*****************************************************

    If txtFields(Index).Tag <> txtFields(Index).Text Then ' the text has changed
        cmdTip(4).Enabled = True
        If Index = 0 Then                               ' the title has changed.
            Dim strTemp As String
            strTemp = txtFields(Index).Text
            strTemp = TitleReTitle(strTemp)             ' I need to pass an actual string - Open quote
            txtFields(Index).Text = TestTitle(strTemp)  ' test for duplicate titles
        End If
    End If
    Call ControlLostFocus(lblDetail(Index))     ' switch off visual focus
    SSTab1.ForeColor = vbDefault                ' switch off visual focus

End Sub
Private Sub Frame1_DblClick()
'*****************************************************
' Purpose:  To close the sublist view & reset the list to view all
'           We are using the frame as a short cut
' Inputs:   None ' Returns:  None
'*****************************************************
On Error GoTo myErrorHandler
    Dim i As Integer                        ' dimension a temp variable
    bSubList = False                        'reset the flag
     For i = i To optViewType.Count - 1     ' clears any values from the frame.
        If optViewType(i).Value = True Then ' found the one value that was selected
            optViewType(i).Value = False
            optViewType(i).ForeColor = vbDefault
            optViewType_LostFocus (i)
            Exit For                        ' exit this routine
        End If                              ' close if statement
        i = i + 1                           ' increment the counter
     Next i                                 ' search the next
     Call cmdTip_Click(0)                 ' rebuild the recordset
     Set moRS = BuildRecordSet              ' reset the list
     FirstRecordShow                        ' display the first record
    Exit Sub                                ' exit the routine
myErrorHandler:
    ErrMsgBox (Err.Description & " " & Err.Number & " in Frame1_DblClick of frmTblTips")
    Resume Next
End Sub

Private Sub lblDetail_DblClick(Index As Integer)
'*****************************************************
' Purpose:  One label acts as a parameter for the type of query
'           That is performed. Whether the entire database
'           or just the title is searched
' Inputs:   the label index
' Returns:  None
'*****************************************************
    Select Case Index
                
        Case 2  ' Find Al&l
            If lblDetail(Index).Caption = "Find &Quick" Then
                lblDetail(Index).Caption = "Find Al&l"
                lblDetail(Index).ToolTipText = "Double Click on this label to do a title only search."
            Else
                lblDetail(Index).Caption = "Find &Quick"
                lblDetail(Index).ToolTipText = "Double Click on this label to do a more thorough search."
            End If
        Case 3  ' T&ype
            lblDetail(Index).Caption = "T&ype"
            Visible = False
            WindowState = vbNormal
            Move (fMDI.ScaleWidth - Me.Width), fMDI.Top
            Visible = True
            fLanguage.Show
            fMDI.Arrange vbCascade
        Case 4  ' SubType
            Visible = False
            WindowState = vbNormal
            Move (fMDI.ScaleWidth - Me.Width), fMDI.Top
            Visible = True
            fSubType.Show
    End Select
    If cmdFind.Caption = "Clear" Then
        cmdFind.Caption = "Find"
    End If
End Sub

Public Sub lstTitle_Click()
'*****************************************************
' Purpose:  This is the event that populates the list'
'           and loads the selected tip's info into the form fields
' Inputs:   None    ' Returns:  None
'*****************************************************
    On Error GoTo lstTitle_ERR
    Dim strSql As String
    
    cmdTip(4).Tag = "Show" 'tells DataHasChanged that it's just being shown, not changed
    ' load the record into the fields
    moRS.FindFirst "strTitle='" & lstTitle.Text & "'"   ' find the selected tip. load its info onto the form
    txtFields(0) = moRS!strTitle & ""                   ' Populate the title
    txtFields(3) = moRS!memNotes & ""                   ' Populate the Notes
    txtFields(4) = moRS!memCode & ""                    ' Populate the Code
    txtFields(1) = moRS!strIndex & ""                   ' Populate the Webpage
    cmbTipSubType = moRS!strSTTitle & ""                ' Populate the From Combobox
    cmbTipType.ListIndex = moRS!intTypeID & ""          ' Populate the Tip Type Combobox
    cmbTipType.Text = moRS!strLang & ""                 ' Populate the name of the current tiptype
    cmbTipType.ToolTipText = moRS!strToolTip & ""       ' add the tool tip text
    Call DualSubTypeLoad(SubTypeQry(cmbTipType.ListIndex), cmbTipSubType, fMDI.cmbTipSubType)
    fMDI.cmbTipType.ListIndex = cmbTipType.ListIndex
    fMDI.cmbTipType.Text = cmbTipType.Text
    fMDI.cmbTipSubType.Text = cmbTipSubType.Text
    cmbTipSubType.ToolTipText = moRS!strSTToolTip & ""      ' label the subType combo box
    cmdTip(0).ToolTipText = "View all " & moRS!strToolTip & " tips."
    cmdTip(4).Tag = ""                          ' clear the flag
    ' set the flag values. This will allow us to add, modify and delete data in the child tables.
    iFKey = moRS!lngTblTipsID                   ' foreign key to both tables
    If Len(txtFields(3)) Then
        bNoteLen = True
    Else: bNoteLen = False                      ' there is preexisting data
    End If
    If Len(txtFields(4)) Then
        bCodeLen = True
    Else: bCodeLen = False
    End If
    Call FormCaption(Me, txtFields(0).Text & " (" & cmbTipType.Text & " Tip For " & cmbTipSubType.Text & ")")                          ' the caption has changed
    Refresh
    Exit Sub                                    ' exit the routine
lstTitle_ERR:
    Select Case Err.Number
        Case 94     'we don't care if a field is empty
            Resume Next
        Case Else
            ErrMsgBox (Err.Description & " " & Err.Number & " in lstTitle_Click of frmTblTips")
            Resume Next
    End Select
End Sub

Private Sub optViewType_Click(Index As Integer)
'*****************************************************
' Purpose:  This is to view a sublist of all the available tips
' Inputs:   None    ' Returns:  None
'*****************************************************
    
    OptionSelect (Index)
    lstTitle.ToolTipText = "Click to move to a selected " & cmbTipType.ToolTipText & " tip."
   
End Sub

Private Sub optViewType_GotFocus(Index As Integer)
'*****************************************************
' Purpose:  Set the current option button to visually look active
' Assumes:  modControls
' Inputs:   None    ' Returns:  None
'*****************************************************
    Call ControlGotFocus(optViewType(Index))    ' light up the label
    Frame1.ForeColor = optViewType(Index).ForeColor ' match frame to label
End Sub

Private Sub optViewType_LostFocus(Index As Integer)
'*****************************************************
' Purpose:  Reset the control to default
' Inputs:   None    ' Returns:  None
'*****************************************************
    Call ControlLostFocus(optViewType(Index))   ' reset the label
    Frame1.ForeColor = optViewType(Index).ForeColor ' match frame to label
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'*****************************************************
' Purpose:  if the form was resized, resize the tab just clicked on
' Inputs:   None  ' Returns:  None
'*****************************************************
    If bForm = True Then Call SSTabResize(Me)
    Call TabControlsMove(Me)
End Sub

'**********************************************************************************************************
'
'Purpose:  All of the general subs and functions
'
'**********************************************************************************************************

Private Function OptionSelect(optIndex As Integer)
'*****************************************************
' Purpose:  This is to view a sublist of all the available tips
' Inputs:   the option Index    ' Returns:  None
' Comment:  This has an int value assigned to it to pass that value
'           back to the calling function.
'*****************************************************
    bSubList = True          ' set a flag based on the sublist
    If txtFields(2).Text <> "" Then txtFields(2).Text = ""
    Call StatusMsgDisplay(TipCount(ListPopulate(ViewOneTypeTip(optIndex), lstTitle)), 2)  ' creates a sublist of an access datatable
    Call FirstRecordShow     'set it to the first record in the list

End Function

Public Sub FirstRecordShow()
'*****************************************************
' Purpose:  Populate the details text controls with the first record
' Inputs:   None    ' Returns:  None
'*****************************************************
    'set it to the first record in the list
    If lstTitle.ListCount > 0 Then
        Call ListNavigate(0, lstTitle, moRS, bSubList) 'triggers the lstTitle_click event
    Else                ' Title only search
        InputInfoBox ("There are no records. List Count = " & lstTitle.ListCount)
    End If
End Sub

'**********************************************************************************************************
'
'Purpose:  Form level validation routines
'
'**********************************************************************************************************
Private Sub CancelUpdate()
    If lstTitle.ListCount <> 0 Then     ' there are still records
        If cmdTip(2).Tag = "New" Then      ' if it was an add
            Call ListNavigate(0, lstTitle, moRS, bSubList)        ' load the first tip
            Call ToggleButtons          ' reset the buttons
            cmdTip(2).Tag = "&Add"
        Else                            ' if it was an edit
            Call ToggleButtons          ' reset the buttons
        End If                          ' end nested if
    Else
    '    Call form_activate 'this needs to check for no tips
    End If
    Call lstTitle_Click                 'This is the event that populates the list
End Sub
'**********************************************************************************************************
'
'Purpose:  Form visual feedback and positioning
'
'**********************************************************************************************************

Private Sub ToggleButtons()
'*****************************************************
' Purpose: toggle buttons enabled/disabled when data changes or is saved/canceled
' Inputs:   None    ' Returns:  None
'*****************************************************
    Dim i As Integer
    Do While i < 4
        cmdNav(i).Enabled = Not cmdNav(i).Enabled ' toggle buttons
        i = i + 1
    Loop

    cmdTip(0).Enabled = Not cmdTip(0).Enabled
    cmdTip(1).Enabled = Not cmdTip(1).Enabled
    cmdTip(2).Enabled = Not cmdTip(2).Enabled
    cmdTip(4).Enabled = Not cmdTip(4).Enabled
    lstTitle.Enabled = Not lstTitle.Enabled
    txtFields(2).Enabled = Not txtFields(2).Enabled ' search text box
    lblDetail(2).Enabled = Not lblDetail(2).Enabled
    cmdFind.Enabled = Not cmdFind.Enabled
    
    If cmdTip(3).Caption = "Delete" Then
        cmdTip(3).Caption = "C&ancel"
    Else
        cmdTip(3).Caption = "Delete"
    End If
    Exit Sub
myErrorHandler:
    ErrMsgBox ("Error has occured during Private sub 'ToggleButtons' " & Chr(13) & Chr(10) & Err.Description & " " & Err.Number)
    Resume Next
End Sub

Private Sub DataHasChanged()
'*****************************************************
' Purpose:  if the data is being edited--not just shown
' Inputs:   None    ' Returns:  None
'*****************************************************

    If cmdTip(4).Tag <> "Show" Then
        'if the buttons haven't already been toggled
        If cmdTip(3).Caption = "Delete" Then
            Call ToggleButtons
        End If
        If cmdTip(4).Tag = "Update" Then
            cmdTip(4).Tag = ""
        ElseIf cmdTip(4).Tag = "" Then
            cmdTip(4).Tag = "Update"
            cmdTip(0).Enabled = False       'disable the control
            cmdTip(1).Enabled = False       'disable the control
        End If
    End If
End Sub
