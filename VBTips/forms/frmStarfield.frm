VERSION 5.00
Begin VB.Form frmStars 
   BackColor       =   &H00000000&
   Caption         =   "Author: James Robert Fleming"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   8610
   Begin VB.Timer Timer1 
      Interval        =   4
      Left            =   1080
      Top             =   3480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   2040
      Picture         =   "frmStarfield.frx":0000
      ScaleHeight     =   3945
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   1560
      Width           =   4530
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "He may be reached at umbcifsm@netscape.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   5760
      Width           =   8175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "James Rober Fleming is a self taught VB Programmer and student at UMBC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8175
   End
End
Attribute VB_Name = "frmStars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Gen Declarations

' You can adjust the variable "intSpeed" for speed of stars

' and constant "intNumStars" for the number you want.

Const intNumStars = 200
Dim intSpeed As Integer
Dim dStars(8, intNumStars) As Double
Dim dDistFromOrigin As Double
Dim dAngle As Double

Private Sub Form_Activate()

    ' Make standard X,Y coordinates, i.e X=0 and Y = 0 is center of form

    Me.ScaleTop = 1000 ' Set scale for top of grid.
    Me.ScaleLeft = -1000 ' Set scale for left of grid.
    Me.ScaleWidth = 2000 ' Set scale (-1 to 1).
    Me.ScaleHeight = -2000
    Me.BackColor = &H0&
    
    Randomize
    
    intSpeed = 5
    
    Dim Y As Integer
    ' Intitialize
    
    For Y = 0 To intNumStars
        dDistFromOrigin = (2 * Me.ScaleTop * Rnd) ' * (-1) ^ Int(1 + Rnd * 2)
        dAngle = (360 * Rnd)
        dStars(1, Y) = dAngle
        dStars(2, Y) = intSpeed
        dStars(3, Y) = dDistFromOrigin
        dStars(4, Y) = RGB(250, Rnd * 255, Rnd * 255)   ' Color of star
        dStars(5, Y) = dStars(3, Y) * Cos(dStars(1, Y)) ' New X position
        dStars(6, Y) = dStars(3, Y) * Sin(dStars(1, Y)) ' New Y position
        dStars(7, Y) = dStars(5, Y) ' Old X position
        dStars(8, Y) = dStars(6, Y) ' Old Y position
    Next Y
 End Sub

Private Sub Form_Resize()

    Me.Cls

    If Me.WindowState <> 1 Then
        ' Make standard X,Y coordinates, i.e X=0 and Y = 0 is center of form
        Me.ScaleTop = 1000 ' Set scale for top of grid.
        Me.ScaleLeft = -1000 ' Set scale for left of grid.
        Me.ScaleWidth = 2000 ' Set scale (-1 to 1).
        Me.ScaleHeight = -2000
    End If
End Sub



Private Sub Timer1_Timer()

    Dim Y As Integer
    
    Randomize
    
    For Y = 0 To intNumStars
        ' If star goes out of the forms boundry create reset it.

        If ((dStars(7, Y) > Me.ScaleLeft + Me.ScaleWidth + 5) Or (dStars(7, Y) < Me.ScaleLeft - 5)) Or ((dStars(8, Y) > Me.ScaleTop + 5) Or (dStars(8, Y) < -(Me.ScaleTop - 5))) Then
            dDistFromOrigin = (1 + 1000 * Rnd) ' * (-1) ^ Int(1 + Rnd * 2)
            dAngle = (360 * Rnd)
            
            dStars(1, Y) = dAngle ' Angle about origin.
            dStars(2, Y) = intSpeed ' Speed of star.
            dStars(3, Y) = dDistFromOrigin ' Distance from origin.
            dStars(4, Y) = RGB(255, 255, 255)  ' Color of star
            dStars(5, Y) = dStars(3, Y) * Cos(dStars(1, Y)) ' New X position
            dStars(6, Y) = dStars(3, Y) * Sin(dStars(1, Y)) ' New Y position
            dStars(7, Y) = dStars(5, Y) ' Old X position
            dStars(8, Y) = dStars(6, Y) ' Old Y position
        End If

        
        ' Place new star.

        PSet (dStars(5, Y), dStars(6, Y)), dStars(4, Y) ' RGB(255, 255, 255)
        ' Add another pixel to make it a little brighter if you want

        'PSet (dStars(5, y) + 1, dStars(6, y)), dStars(4, y) ' RGB(255, 2

        '     55, 255)

        
        ' Erase star at old position

        PSet (dStars(7, Y), dStars(8, Y)), RGB(0, 0, 0)
        ' Add another pixel to make it a little brighter if you want

        'PSet (dStars(7, y) + 1, dStars(8, y)), RGB(0, 0, 0)

        
        ' Increase the distance from the origin for the star.

        dStars(3, Y) = dStars(3, Y) + dStars(2, Y)
        dStars(7, Y) = dStars(5, Y) ' Old X position
        dStars(8, Y) = dStars(6, Y) ' Old Y position
        dStars(5, Y) = dStars(3, Y) * Cos(dStars(1, Y)) ' New X position
        dStars(6, Y) = dStars(3, Y) * Sin(dStars(1, Y)) ' New Y position
        
        ' Increase the stars speed each time so that it appears to

        ' move faster as it nears the edge of the form.

        dStars(2, Y) = dStars(2, Y) + intSpeed
    Next Y

End Sub


