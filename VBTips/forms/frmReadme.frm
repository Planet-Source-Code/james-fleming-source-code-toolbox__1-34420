VERSION 5.00
Begin VB.Form frmReadme 
   BackColor       =   &H80000018&
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   7920
   Begin VB.Label lblDetails 
      BackStyle       =   0  'Transparent
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
      Height          =   3855
      Left            =   2400
      TabIndex        =   6
      Top             =   2400
      Width           =   5055
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
      Top             =   4350
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
      Top             =   3885
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
      Top             =   3435
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
      Top             =   2970
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
      Top             =   2520
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
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblOverview.Caption = "First I would like to say thank you for downloading SourceCode from CyberSpace: " & _
    "a programmer 's productivity tool!" & vbCrLf & vbCrLf & _
    "Second, I would like to ask: Please vote for me at http://www.planetsourcecode.com." & _
    "They hold a monthly contest for programmer of the month, and I need your vote." & vbCrLf & vbCrLf & _
    "This application is designed to make creating software applications faster and easier by giving" & vbCrLf & _
    "developers a searchable database for source code."
    
End Sub

Private Sub lblFeatures_Click(Index As Integer)
Select Case Index
    Case 0
        lblDetails = "Catalogues source code in multiple languages." & vbCrLf & _
            "Extensive Searching and sorting options" & vbCrLf & _
            "Function import features." & vbCrLf & _
            "Microsoft windows design." & vbCrLf & _
            "Tip of the Day feature lets you reinforce your knowledge." & vbCrLf & _
            "MS Visual Basic error codes table for non-fatal errors." & vbCrLf & _
            "File viewer lets you open up and view multiple file types." & vbCrLf & _
            "An HTML and VB BAS module template generator."

    Case 1
        lblDetails = "Programming tips are catalogued by type (ASP, C++, VB, etc)." & vbCrLf & _
            "Up to 12 languages catagories can be stored." & vbCrLf & _
            "Language catagories can be modified throught the interface without any recoding." & vbCrLf & _
            "Tips can be cataloged by language subtype (string handling, math, date & time, etc)." & vbCrLf & _
            "Infinitely many subtype catagories can be added through the interface." & vbCrLf & _
            "Tips can be further catalogued by coding tip title, and keyword." & vbCrLf & _
            "Duplicate titles are auto-incremented so you don't have to worry about data integrity constraints."
    
    Case 2
        lblDetails = "Source code tips can be sorted and viewed by title, language, or most recent entry." & vbCrLf & _
            "Tips can be further sorted by language and subtype (alphabetically, by all subtypes alphabetically, or only one subtype)." & vbCrLf & _
            "Notes about source code can be stored independant of the code: You can store additional comments separate from the source code." & vbCrLf & _
            "Keyword searches can be of a broad or narrow scope: You may perform a keyword search in only the title and keyword fields or you may search all fields."
    Case 3
        lblDetails = "Import feature lets you search for and import Visual Basic functions and subroutines from previous projects." & vbCrLf & _
            "Import searches can be by drive, path or you may browse for the desired folder." & vbCrLf & _
            "Import form can save your last path searched." & vbCrLf & _
            "Import allows you to copy code directly into the database or clipboard."

    Case 4
        lblDetails = "Over 100 coding tips included in the database." & vbCrLf & _
            "All open source code." & vbCrLf & _
            "Clear and generous notes (about 2000 lines of notes)" & vbCrLf & _
            "Documentation on naming convention used (see modNamingConventions)." & vbCrLf & _
            "Small reusable BAS Modules."
            
End Select

End Sub
