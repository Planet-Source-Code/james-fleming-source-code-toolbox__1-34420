VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#1.0#0"; "IMGEDIT.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFileView 
   Caption         =   "File Viewer"
   ClientHeight    =   6405
   ClientLeft      =   -72540
   ClientTop       =   345
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   7545
   Begin VB.HScrollBar HScroll1 
      Height          =   250
      Left            =   10680
      TabIndex        =   4
      Top             =   8400
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8000
      LargeChange     =   5
      Left            =   11640
      Max             =   100
      Min             =   1
      TabIndex        =   3
      Top             =   0
      Value           =   100
      Width           =   250
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   5640
      ScaleHeight     =   6375
      ScaleWidth      =   1935
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   2655
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   4683
      _StockProps     =   0
      ImageControl    =   "ImgEdit1"
      AnnotationBackColor=   12632256
      BorderStyle     =   0
      AutoRefresh     =   -1  'True
   End
   Begin RichTextLib.RichTextBox rtbFileView 
      Height          =   6375
      Left            =   -7200
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   11245
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmFileView.frx":0000
   End
End
Attribute VB_Name = "frmFileView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fLoad As Boolean
Public fDirty As Boolean

Private Sub Form_Activate()
    fLoad = False
       If rtbFileView.Left = 0 Then            ' disable the inappropriate menu items
        VScroll1.Visible = False
        HScroll1.Visible = False
        Picture1.Visible = False
        ImgEdit1.Visible = False
        rtbFileView.Visible = True
    ElseIf Picture1.Left = 0 Then           ' disable the inappropriate menu items
        fMDI.mnuFileSave.Enabled = False
        fMDI.mnuFileSaveAs.Enabled = False
        rtbFileView.Visible = False
        ImgEdit1.Visible = False
        Picture1.Visible = True
   ElseIf ImgEdit1.Left = 0 Then            ' disable the inappropriate menu items
        fMDI.mnuFileSave.Enabled = False
        fMDI.mnuFileSaveAs.Enabled = False
        rtbFileView.Visible = False
        VScroll1.Visible = False
        HScroll1.Visible = False
        ImgEdit1.Visible = True
    End If
    Call StatusMsgDisplay("The File Viewer form is active.", 2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Picture1.Left = 0 Then Call PicScroll(KeyCode, VScroll1, HScroll1)
End Sub

Private Sub Form_Load()
    fLoad = True
    Call StatusMsgDisplay("The File Viewer form is active", 2)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If fDirty = True Then
        Dim intResponse As Integer
        intResponse = YesNo("The text has been changed. Save changes?")
        If intResponse = 1 Then Me.rtbFileView.SaveFile fMDI.dlgMDI.FileName, rtfText
    End If
End Sub

Private Sub Form_Resize()
    If fLoad = True Then Exit Sub
    If rtbFileView.Left = 0 Then
        rtbFileView.Move 0, 0, (ScaleWidth), ScaleHeight
    ElseIf Picture1.Left = 0 Then
        If Me.WindowState = vbMaximized Then
            Call PicBoxResize(Me, Picture1, VScroll1, HScroll1)
        Else
            VScroll1.Visible = False
            HScroll1.Visible = False
        End If
    ElseIf ImgEdit1.Left = 0 Then
        ImgEdit1.Move 0, 0, (ScaleWidth), ScaleHeight
        ImgEdit1.Visible = True
    End If
    Me.Refresh
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'*****************************************************
' Purpose:  That's all folks!
' Inputs:   None     Returns:  None
'*****************************************************
    Screen.MousePointer = vbDefault ' reset the mouse pointer
    If g_blnUnload = True Then Exit Sub
    fMDI.mnuFileSave.Enabled = True
    fMDI.mnuFileSaveAs.Enabled = True
    Call StatusMsgDisplay("There are no active forms", 2)
End Sub

Private Sub rtbFileView_Change()
    If fLoad = False Then
        fDirty = True
    End If
End Sub

Private Sub VScroll1_Change()
   If Picture1.Left = 0 Then Call VScrollChange(Picture1, VScroll1)
End Sub
