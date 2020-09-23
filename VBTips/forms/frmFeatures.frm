VERSION 5.00
Begin VB.Form frmFeatures 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Application Features"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmFeatures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFeatureShow 
      BackColor       =   &H80000018&
      Caption         =   "Don't show this again"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label lblDetails 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on one of the feature categories at left to learn more about this application."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3000
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   6975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFeatures 
      BackStyle       =   0  'Transparent
      Caption         =   "Programming:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   4110
      Width           =   1815
   End
   Begin VB.Label lblFeatures 
      BackStyle       =   0  'Transparent
      Caption         =   "Import Projects:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Top             =   3645
      Width           =   1815
   End
   Begin VB.Label lblFeatures 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sorting && Searching:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   3195
      Width           =   1815
   End
   Begin VB.Label lblFeatures 
      BackStyle       =   0  'Transparent
      Caption         =   "Cataloging:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2730
      Width           =   1815
   End
   Begin VB.Label lblFeatures 
      BackStyle       =   0  'Transparent
      Caption         =   " Application:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblOverview 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
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
      Height          =   1800
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9045
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFeatures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    chkFeatureShow = GetSetting(App.Title, "Settings", "chkFeatureShow", 1)
    lblOverview.Caption = "First I would like to say thank you for downloading SourceCode from CyberSpace: " & _
    "The programmer 's productivity tool! " & _
    "This application is designed to make creating software applications faster and easier by giving " & _
    "developers a searchable database for source code." & vbCrLf & vbCrLf & _
    "Second, I would like to ask: Please vote for me at http://www.planetsourcecode.com." & vbCrLf & _
    "They hold a monthly contest for programmer of the month, and I need your vote."

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "chkFeatureShow", chkFeatureShow.Value
End Sub

Private Sub lblFeatures_Click(Index As Integer)
Static iPrevious As Integer
    lblFeatures(iPrevious).ForeColor = &H8000&
    lblFeatures(Index).ForeColor = &HC000&
Select Case Index
    Case 0
        lblDetails = "Catalogues source code in multiple languages." & vbCrLf & _
            "Extensive Searching and sorting options" & vbCrLf & _
            "Function import features." & vbCrLf & _
            "Microsoft windows design." & vbCrLf & _
            "Tip of the Day feature lets you reinforce your knowledge." & vbCrLf & _
            "MS Visual Basic error codes table for non-fatal errors." & vbCrLf & _
            "File viewer lets you open up and view multiple file types." & vbCrLf & _
            "A 12 Page Systems Analysis Survey in MS Word." & vbCrLf & _
            "An HTML and VB BAS module template generator."

    Case 1
        lblDetails = "Programming tips are catalogued by type (ASP, C++, VB, etc)." & vbCrLf & _
            "Up to 12 languages catagories can be stored." & vbCrLf & _
            "Language catagories can be modified throught the interface." & vbCrLf & _
            "Tips can be cataloged by language subtype (string handling, math, etc)." & vbCrLf & _
            "Infinitely many subtype catagories can be added through the interface." & vbCrLf & _
            "Tips can be further catalogued by coding tip title, and keyword." & vbCrLf & _
            "Duplicate titles are auto-incremented so you don't have to worry about data integrity constraints."
    
    Case 2
        lblDetails = "Code tips can be sorted and viewed by title, language, or most recent entry." & vbCrLf & _
            "Tips can be further sorted by language and subtype." & vbCrLf & _
            "Notes about source code can be stored independant of the code." & vbCrLf & _
            "You can store additional comments separate from the source code." & vbCrLf & _
            "Keyword searches can be of a broad or narrow scope." & vbCrLf & _
            "You may perform a keyword search in only the title and keyword fields or you may search all fields."
    Case 3
        lblDetails = "Import feature lets you import VB functions and subs from previous projects." & vbCrLf & _
            "Import searches can be by drive, path or you may browse for a folder." & vbCrLf & _
            "Import form can save your last path searched." & vbCrLf & _
            "Import allows you to copy code directly into the database or clipboard."

    Case 4
        lblDetails = "Over 100 coding tips included in the database." & vbCrLf & _
            "All open source code." & vbCrLf & _
            "Clear and generous notes (about 2000 lines of notes)" & vbCrLf & _
            "Documentation on naming convention used (see modNamingConventions)." & vbCrLf & _
            "Small reusable BAS Modules."
End Select
    
    iPrevious = Index ' set the last index to a static var.

End Sub
