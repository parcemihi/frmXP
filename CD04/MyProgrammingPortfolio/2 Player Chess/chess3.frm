VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Chess 
   Caption         =   "Keith's Chess Game"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Game File"
      Filter          =   "*.keg"
      InitDir         =   "app.path"
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   63
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   62
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   61
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   60
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   59
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   58
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   57
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   56
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   55
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   54
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   53
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   52
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   51
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   50
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   49
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   48
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   47
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   46
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   45
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   44
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   43
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   42
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   41
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   40
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   39
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   38
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   37
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   36
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   35
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   34
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   33
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   32
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   31
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   30
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   29
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   28
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   27
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   26
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   25
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   24
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   23
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   22
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   21
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   20
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   19
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   18
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   17
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   16
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   15
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   14
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   13
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   12
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   11
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   10
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   9
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   8
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   7
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   6
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   5
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   4
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   3
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   615
      Index           =   1
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   615
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4800
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

Dim Piece_Name(0 To 31) As String


Dim P2M As Integer
Dim InPlay As Boolean
Dim BoardColor As String

Private Sub board_Click(Index As Integer)


If InPlay = True Then
    board(P2M).BackColor = BoardColor
    board(Index).Picture = board(P2M).Picture
    If Not Index = P2M Then
        board(P2M).Picture = LoadPicture(App.Path & "\images\Blank.gif")
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
Dim x
SaveGame
End Sub

Private Sub Form_Load()
Dim i As Integer
Game_Default_Values
SetUpBoard
End Sub
Private Sub cmdExit_Click()
End
End Sub

