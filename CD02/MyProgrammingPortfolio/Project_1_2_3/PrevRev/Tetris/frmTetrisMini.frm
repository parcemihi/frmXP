VERSION 5.00
Begin VB.Form frmTetris 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   1800
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   11
      Left            =   2160
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   10
      Left            =   1800
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   9
      Left            =   1440
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   8
      Left            =   2160
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   7
      Left            =   1800
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   6
      Left            =   1440
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   5
      Left            =   2160
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   1800
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   1440
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   2160
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   1800
      Top             =   720
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   1440
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBlue 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":0000
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image imgPurple 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":0702
      Top             =   2160
      Width           =   360
   End
   Begin VB.Image imgRed 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":0E04
      Top             =   2520
      Width           =   360
   End
   Begin VB.Image imgTurquoise 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":1506
      Top             =   1800
      Width           =   360
   End
   Begin VB.Image imgGreen 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":1C08
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image imgGray 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":230A
      Top             =   1080
      Width           =   360
   End
   Begin VB.Image imgYellow 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetrisMini.frx":2A0C
      Top             =   720
      Width           =   360
   End
End
Attribute VB_Name = "frmTetris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i1

Private Sub redBlock()
If i1 > 11 Then
    i1 = 0
Else
    Image1(i1).Visible = True
    Image1(i1) = imgRed.Picture
    i1 = i1 + 3
End If
End Sub
Private Sub Command1_Click()
Const row1 = 1
Const row2 = 2
Const row3 = 3
Const row4 = 4
Const row5 = 5
Const row6 = 6
Const row7 = 7
Const row8 = 8
Const row9 = 9
Const row10 = 10
Const row11 = 11
Const row12 = 12
Const row13 = 13
Const row14 = 14
Const row15 = 15


Dim check As Integer, count As Integer
Dim A As Integer, B As Integer

Select Case check
    Case check = A To B
        check = 0
        For check = 0 To 2
        If Image1(check).Visible = True Then
        count = count + 1
            If count = 3 Then
'            count = 0
                For count = 0 To Row 'Move all blocks above, down
                    count = count + 10
                    
                Next count
            End If
        
        
        
        End If
 
Next check
 
 
 'Select Case check
    Case check = 3 To 5
    
                For count = 3 To 5
            
                Next count
    
    Case check = 6 To 8
                For count = 6 To 8
            
                Next count


End Select
End Sub

Private Sub Form_Load()





'    Timer1.Enabled = True
'    Timer1.Interval = 500
End Sub

Private Sub Timer1_Timer()
    redBlock
End Sub
