VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Programmers Toolbox"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About VBTips"
   Begin VB.CommandButton cmdAuthor 
      Cancel          =   -1  'True
      Caption         =   "About Author"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   2640
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   1785
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4245
      TabIndex        =   1
      Tag             =   "&System Info..."
      Top             =   2235
      Width           =   1245
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   240
      TabIndex        =   3
      Tag             =   "App Description"
      Top             =   120
      Width           =   5085
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   1905
      Left            =   240
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   5465
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   240
      X2              =   5450
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const KEY_ALL_ACCESS = &H2003F

' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdAuthor_Click()
    Unload Me
    fAuthor.Show
End Sub

Private Sub Form_Activate()
' Assumes: modFormUtilities
    Call StatusMsgDisplay("The About form is active.", 2)
End Sub

Private Sub Form_Load()
    Me.Caption = "About box " & App.Title & " Version: " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.FileDescription
    lblDisclaimer.Caption = "Warning: ... There has been a great deal of care in developing this application, but it is far from perfect." & _
    "This application has not been tested extensively. Currently this program is being released as freeware without any warrantee, but please report bugs to umbcifsm@netscape.net. I will do what I can to resolve any issues for you." & vbCrLf & "James"

End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub StartSysInfo()
On Error GoTo SysInfoErr

    Dim rc As Long
    Dim SysInfoPath As String
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
            ' Validate Existance Of Known 32 Bit File Version
            If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                    SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            ' Error - File Can Not Be Found...
            Else
                GoTo SysInfoErr
            End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    Call Shell(SysInfoPath, vbNormalFocus)
    Exit Sub
SysInfoErr:
        MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                            
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit

GetKeyError:    ' Cleanup After An Error Has Occured...
        KeyVal = ""                                         ' Set Return Val To Empty String
        GetKeyValue = False                                 ' Return Failure
        rc = RegCloseKey(hKey)                              ' Close Registry Key
End Function

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub
