VERSION 5.00
Begin VB.Form Chess1 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   63
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   62
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   61
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   60
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   59
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   58
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   57
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   56
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   55
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   54
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   53
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   52
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   51
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   50
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   49
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   48
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   47
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   46
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   45
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   44
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   43
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   42
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   41
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   40
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   39
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   38
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   37
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   36
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   35
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   34
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   33
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   32
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   31
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   30
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   29
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   28
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   27
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   26
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   25
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   24
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   23
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   22
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   21
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   20
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   19
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   18
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   17
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   16
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   15
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   14
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   13
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   12
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   11
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   10
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   9
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   8
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   7
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   6
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   5
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   4
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   3
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000007&
      Height          =   855
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton board 
      BackColor       =   &H80000009&
      Height          =   855
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6000
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   12
      Left            =   8040
      Picture         =   "Chess.frx":0000
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   11
      Left            =   8040
      Picture         =   "Chess.frx":04BE
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   10
      Left            =   8040
      Picture         =   "Chess.frx":09E4
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   9
      Left            =   8040
      Picture         =   "Chess.frx":0F24
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   8
      Left            =   8040
      Picture         =   "Chess.frx":1426
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   7
      Left            =   8040
      Picture         =   "Chess.frx":1A0C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   6
      Left            =   7200
      Picture         =   "Chess.frx":1F1D
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   5
      Left            =   7200
      Picture         =   "Chess.frx":236A
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   4
      Left            =   7200
      Picture         =   "Chess.frx":27FF
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   3
      Left            =   7200
      Picture         =   "Chess.frx":2C99
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   2
      Left            =   7200
      Picture         =   "Chess.frx":310B
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   1
      Left            =   7200
      Picture         =   "Chess.frx":360C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   855
   End
   Begin VB.Image pieces 
      Height          =   855
      Index           =   0
      Left            =   7200
      Picture         =   "Chess.frx":3A93
      Stretch         =   -1  'True
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Chess1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
