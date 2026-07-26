VERSION 5.00
Begin VB.Form frmTicTacToe 
   Caption         =   "TicTacToe"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   2430
   LinkTopic       =   "Form1"
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   162
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox T1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox T2 
      Height          =   735
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox T3 
      Height          =   735
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox T4 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox T5 
      Height          =   735
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox T6 
      Height          =   735
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox T7 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox T8 
      Height          =   735
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox T9 
      Height          =   735
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.Image imgWhite 
      Height          =   720
      Left            =   840
      Picture         =   "tictactoe.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   720
   End
   Begin VB.Image imgBlank 
      Height          =   720
      Left            =   840
      Picture         =   "tictactoe.frx":1C96
      Stretch         =   -1  'True
      Top             =   840
      Width           =   720
   End
   Begin VB.Image imgBlack 
      Height          =   720
      Left            =   840
      Picture         =   "tictactoe.frx":392C
      Stretch         =   -1  'True
      Top             =   840
      Width           =   720
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu computer 
         Caption         =   "Play the computer"
         Checked         =   -1  'True
      End
      Begin VB.Menu New 
         Caption         =   "New"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public whosTurn As Boolean
Public message
Public comply As Boolean
Dim comp As Integer, makeMove As Boolean
'Dim sw1 As Boolean, sw2 As Boolean, sw3 As Boolean
'Dim sw4 As Boolean, sw5 As Boolean, sw6 As Boolean
'Dim sw7 As Boolean, sw8 As Boolean, sw9 As Boolean

Private Sub Form_Load()
    ResetGame
    
End Sub
Public Function ResetGame()
    'sw1 = False
    'sw2 = False
    'sw3 = False
    'sw4 = False
    'sw5 = False
    'sw6 = False
    'sw7 = False
    'sw8 = False
    'sw9 = False
    T1.Picture = imgBlank
    T2.Picture = imgBlank
    T3.Picture = imgBlank
    T4.Picture = imgBlank
    T5.Picture = imgBlank
    T6.Picture = imgBlank
    T7.Picture = imgBlank
    T8.Picture = imgBlank
    T9.Picture = imgBlank

    If computer.Checked = True Then
        comply = False
        CompPlay
    End If
End Function

Public Function CompPlay()
makeMove = False
'//ROW 1
If T3.Picture = imgWhite And T2.Picture = imgWhite Or _
   T7.Picture = imgWhite And T4.Picture = imgWhite Or _
   T9.Picture = imgWhite And T5.Picture = imgWhite Then
   T1.Picture = imgWhite
   turn
ElseIf T1.Picture = imgWhite And T3.Picture = imgWhite Or _
   T8.Picture = imgWhite And T5.Picture = imgWhite Then
   T2.Picture = imgBlack
   turn
ElseIf T1.Picture = imgWhite And T2.Picture = imgWhite Or _
   T9.Picture = imgWhite And T6.Picture = imgWhite Or _
   T7.Picture = imgWhite And T5.Picture = imgWhite Then
   T3.Picture = imgBlack
'//ROW 2
   turn
ElseIf T1.Picture = imgWhite And T7.Picture = imgWhite Or _
   T5.Picture = imgWhite And T6.Picture = imgWhite Then
   T4.Picture = imgBlack
   turn
ElseIf T1.Picture = imgWhite And T9.Picture = imgWhite Or _
   T3.Picture = imgWhite And T7.Picture = imgWhite Or _
   T2.Picture = imgWhite And T8.Picture = imgWhite Or _
   T4.Picture = imgWhite And T6.Picture = imgWhite Then
   T5.Picture = imgBlack
   turn
ElseIf T3.Picture = imgWhite And T9.Picture = imgWhite Or _
   T4.Picture = imgWhite And T5.Picture = imgWhite Then
   T6.Picture = imgBlack
'//ROW 3
   turn
ElseIf T1.Picture = imgWhite And T4.Picture = imgWhite Or _
   T3.Picture = imgWhite And T5.Picture = imgWhite Or _
   T9.Picture = imgWhite And T8.Picture = imgWhite Then
   T7.Picture = imgBlack
   turn
ElseIf T7.Picture = imgWhite And T9.Picture = imgWhite Or _
   T2.Picture = imgWhite And T5.Picture = imgWhite Then
   T8.Picture = imgBlack
   turn
ElseIf T3.Picture = imgWhite And T6.Picture = imgWhite Or _
   T7.Picture = imgWhite And T8.Picture = imgWhite Or _
   T1.Picture = imgBlack And T5.Picture = imgWhite Then
   T9.Picture = imgWhite
   turn

'//Computer Win!
ElseIf T3.Picture = imgWhite And T2.Picture = imgWhite Or _
   T7.Picture = imgWhite And T4.Picture = imgWhite Or _
   T9.Picture = imgWhite And T5.Picture = imgWhite Then
   T1.Picture = imgWhite
   turn
ElseIf T1.Picture = imgWhite And T3.Picture = imgWhite Or _
   T8.Picture = imgWhite And T5.Picture = imgWhite Then
   T2.Picture = imgWhite
   turn
ElseIf T1.Picture = imgWhite And T2.Picture = imgWhite Or _
   T9.Picture = imgWhite And T6.Picture = imgWhite Or _
   T7.Picture = imgWhite And T5.Picture = imgWhite Then
   T3.Picture = imgWhite
'//ROW 2
   turn
ElseIf T1.Picture = imgWhite And T7.Picture = imgWhite Or _
   T5.Picture = imgWhite And T6.Picture = imgWhite Then
   T4.Picture = imgWhite
   turn
ElseIf T1.Picture = imgWhite And T9.Picture = imgWhite Or _
   T3.Picture = imgWhite And T7.Picture = imgWhite Or _
   T2.Picture = imgWhite And T8.Picture = imgWhite Or _
   T4.Picture = imgWhite And T6.Picture = imgWhite Then
   T5.Picture = imgWhite
   turn
ElseIf T3.Picture = imgWhite And T9.Picture = imgWhite Or _
   T4.Picture = imgWhite And T5.Picture = imgWhite Then
   T6.Picture = imgWhite
'//ROW 3
   turn
ElseIf T1.Picture = imgWhite And T4.Picture = imgWhite Or _
   T3.Picture = imgWhite And T5.Picture = imgWhite Or _
   T9.Picture = imgWhite And T8.Picture = imgWhite Then
   T7.Picture = imgWhite
   turn
ElseIf T7.Picture = imgWhite And T9.Picture = imgWhite Or _
   T2.Picture = imgWhite And T5.Picture = imgWhite Then
   T8.Picture = imgWhite
   turn
ElseIf T3.Picture = imgWhite And T6.Picture = imgWhite Or _
   T7.Picture = imgWhite And T8.Picture = imgWhite Or _
   T1.Picture = imgWhite And T5.Picture = imgWhite Then
   T9.Picture = imgWhite
   turn
Else
'// Computer First Move
    Do
    Randomize
    comp = Rnd * 9
        If comp = 1 And T1.Picture = imgBlank Then
        T1.Picture = imgBlack
        Exit Do
        ElseIf comp = 2 And T2.Picture = imgBlank Then
        T2.Picture = imgBlack
        Exit Do
        ElseIf comp = 3 And T3.Picture = imgBlank Then
        T3.Picture = imgBlack
        Exit Do
        ElseIf comp = 4 And T4.Picture = imgBlank Then
        T4.Picture = imgBlack
        Exit Do
        ElseIf comp = 5 And T5.Picture = imgBlank Then
        T5.Picture = imgBlack
        Exit Do
        ElseIf comp = 6 And T6.Picture = imgBlank Then
        T6.Picture = imgBlack
        Exit Do
        ElseIf comp = 7 And T7.Picture = imgBlank Then
        T7.Picture = imgBlack
        Exit Do
        ElseIf comp = 8 And T8.Picture = imgBlank Then
        T8.Picture = imgBlack
        Exit Do
        ElseIf comp = 9 And T9.Picture = imgBlank Then
        T9.Picture = imgBlack
        Exit Do
        End If
    Loop
End If
turn
End Function

Public Function checkWin()
Dim winIs As String, i As Boolean

winIs = "And the winner is..."
If T1.Picture = imgWhite And _
   T2.Picture = imgWhite And _
   T3.Picture = imgWhite Or _
   T4.Picture = imgWhite And _
   T5.Picture = imgWhite And _
   T6.Picture = imgWhite Or _
   T7.Picture = imgWhite And _
   T8.Picture = imgWhite And _
   T9.Picture = imgWhite Or _
   T1.Picture = imgWhite And _
   T4.Picture = imgWhite And _
   T7.Picture = imgWhite Or _
   T2.Picture = imgWhite And _
   T5.Picture = imgWhite And _
   T8.Picture = imgWhite Or _
   T3.Picture = imgWhite And _
   T6.Picture = imgWhite And _
   T9.Picture = imgWhite Or _
   T1.Picture = imgWhite And _
   T5.Picture = imgWhite And _
   T9.Picture = imgWhite Or _
   T3.Picture = imgWhite And _
   T5.Picture = imgWhite And _
   T7.Picture = imgWhite Then
        message = MsgBox("White Wins! Would you like to play again?", vbYesNo, "Game Over")
            If message = vbYes Then
                Form_Load
            End If

ElseIf T1.Picture = imgBlack And _
   T1.Picture = imgBlack And T2.Picture = imgBlack And _
   T3.Picture = imgBlack Or _
 _
   T4.Picture = imgBlack And _
   T5.Picture = imgBlack And _
   T6.Picture = imgBlack Or _
   T7.Picture = imgBlack And _
   T8.Picture = imgBlack And _
   T9.Picture = imgBlack Or _
   T1.Picture = imgBlack And _
   T4.Picture = imgBlack And _
   T7.Picture = imgBlack Or _
   T2.Picture = imgBlack And _
   T5.Picture = imgBlack And _
   T8.Picture = imgBlack Or _
   T3.Picture = imgBlack And _
   T6.Picture = imgBlack And _
   T9.Picture = imgBlack Or _
   T1.Picture = imgBlack And _
   T5.Picture = imgBlack And _
   T9.Picture = imgBlack Or _
   T3.Picture = imgBlack And _
   T5.Picture = imgBlack And _
   T7.Picture = imgBlack Then
        message = MsgBox("Black Wins! Would you like to play again?", vbYesNo, "Game Over")
            If message = vbYes Then
                Form_Load
            End If
        

ElseIf (T1.Picture = imgWhite Or T1.Picture = imgBlack) And _
       (T2.Picture = imgWhite Or T2.Picture = imgBlack) And _
       (T3.Picture = imgWhite Or T3.Picture = imgBlack) And _
       (T4.Picture = imgWhite Or T4.Picture = imgBlack) And _
       (T5.Picture = imgWhite Or T5.Picture = imgBlack) And _
       (T6.Picture = imgWhite Or T6.Picture = imgBlack) And _
       (T7.Picture = imgWhite Or T7.Picture = imgBlack) And _
       (T8.Picture = imgWhite Or T8.Picture = imgBlack) And _
       (T9.Picture = imgWhite Or T9.Picture = imgBlack) Then
        message = MsgBox("Cat's Game! Would you like to play again?", vbYesNo, "Game Over")
            If message = vbYes Then
                Form_Load
            End If
End If
If whosTurn = False Then
    CompPlay
    ElseIf whosTurn = True Then
    '
End If
End Function

Private Sub T1_Click()
If T1.Picture = imgBlank And whosTurn = True Then
    T1.Picture = imgWhite
End If
turn
End Sub
Private Sub T2_Click()
If T2.Picture = imgBlank And whosTurn = True Then
    T2.Picture = imgWhite
End If
turn
End Sub
Private Sub T3_Click()
If T3.Picture = imgBlank And whosTurn = True Then
    T3.Picture = imgWhite
End If
turn
End Sub
Private Sub T4_Click()
If T4.Picture = imgBlank And whosTurn = True Then
    T4.Picture = imgWhite
End If
turn
End Sub
Private Sub T5_Click()
If T5.Picture = imgBlank And whosTurn = True Then
    T5.Picture = imgWhite
End If
turn
End Sub
Private Sub T6_Click()
If T6.Picture = imgBlank And whosTurn = True Then
    T6.Picture = imgWhite
End If
turn
End Sub
Private Sub T7_Click()
If T7.Picture = imgBlank And whosTurn = True Then
    T7.Picture = imgWhite
End If
turn
End Sub
Private Sub T8_Click()
If T8.Picture = imgBlank And whosTurn = True Then
    T8.Picture = imgWhite
End If
turn
End Sub
Private Sub T9_Click()
If T9.Picture = imgBlank And whosTurn = True Then
    T9.Picture = imgWhite
End If
turn
End Sub
Public Function turn()
    If whosTurn = True Then
        whosTurn = False
    Else: whosTurn = True
    End If
checkWin
End Function
Private Sub New_Click()
    Form_Load
End Sub
Private Sub exit_Click()
End
End Sub

