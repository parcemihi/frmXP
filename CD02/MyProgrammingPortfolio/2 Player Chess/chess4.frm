VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Chess 
   Caption         =   "Keith's Chess Game"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   63
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   120
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Game File"
      Filter          =   "*.keg"
      InitDir         =   "app.path"
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   62
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   61
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   60
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   59
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   58
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   57
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   56
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   55
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   54
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   53
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   52
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   51
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   50
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   49
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   48
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   47
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   46
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   45
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   44
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   43
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   42
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   41
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   40
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   39
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   38
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   37
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   36
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   35
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   34
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   33
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   32
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   31
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   30
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   29
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   28
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   27
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   26
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   25
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   24
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   23
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   22
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   21
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   20
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   19
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   18
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   17
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   16
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   15
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   14
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   13
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   12
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   11
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   10
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   9
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   7
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   6
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   5
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   4
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   2
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   7
      Left            =   6240
      Picture         =   "chess4.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   1
      Left            =   5640
      Picture         =   "chess4.frx":0511
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   12
      Left            =   6240
      Picture         =   "chess4.frx":0998
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   11
      Left            =   6240
      Picture         =   "chess4.frx":0E56
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   10
      Left            =   6240
      Picture         =   "chess4.frx":137C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   9
      Left            =   6240
      Picture         =   "chess4.frx":18BC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   8
      Left            =   6240
      Picture         =   "chess4.frx":1DBE
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   3
      Left            =   5640
      Picture         =   "chess4.frx":23A4
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   2
      Left            =   5640
      Picture         =   "chess4.frx":2816
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   0
      Left            =   5640
      Picture         =   "chess4.frx":2D17
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   6
      Left            =   5640
      Picture         =   "chess4.frx":3095
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   5
      Left            =   5640
      Picture         =   "chess4.frx":34E2
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   615
   End
   Begin VB.Image pieces 
      Height          =   615
      Index           =   4
      Left            =   5640
      Picture         =   "chess4.frx":3977
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   615
   End
   Begin VB.Menu cmdFile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu cmdNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu cmdOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu cmdSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Seperator 
         Caption         =   "-"
      End
      Begin VB.Menu cmdExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Chess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P2M As Integer
Dim InPlay As Boolean
Dim BoardColor As String

Private Sub board_Click(Index As Integer)
If InPlay = True Then
    board(P2M).BackColor = BoardColor
    board(Index).Picture = board(P2M).Picture
    If Not P2M = Index Then
        'legal_move P2M, Index
        Board_Layout(Index) = Board_Layout(P2M)
        Board_Layout(P2M) = 0
        board(P2M).Picture = pieces(Board_Layout(P2M)).Picture
        board(Index).Picture = pieces(Board_Layout(Index)).Picture
    End If
    InPlay = False
Else
    P2M = Index
    BoardColor = board(Index).BackColor
    board(Index).BackColor = vbGreen
    InPlay = True
End If
End Sub

Private Sub cmdNew_Click()
NewGame
SetUpBoard
End Sub

Private Sub cmdOpen_Click()
CommonDialog1.Filter = "Chess Files (*.keg)|*.keg|All Files (*.*)|*.*;*.ico"
CommonDialog1.Action = 1
LoadGame
SetUpBoard
End Sub

Private Sub cmdSave_Click()
CommonDialog1.Filter = "Chess Files (*.keg)|*.keg|All Files (*.*)|*.*;*.ico"
CommonDialog1.ShowSave
SaveGame
End Sub

Private Sub Form_Load()
Game_Default_Values
SetUpBoard
End Sub

Private Sub cmdExit_Click()
End
End Sub

