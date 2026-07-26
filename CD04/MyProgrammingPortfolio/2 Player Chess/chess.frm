VERSION 5.00
Begin VB.Form chess 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   7425
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Image pieces 
      Height          =   720
      Index           =   4
      Left            =   3360
      Picture         =   "chess.frx":0000
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   3
      Left            =   2640
      Picture         =   "chess.frx":0536
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   2
      Left            =   1920
      Picture         =   "chess.frx":0A7A
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   1
      Left            =   1200
      Picture         =   "chess.frx":0F74
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   0
      Left            =   480
      Picture         =   "chess.frx":1562
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   5
      Left            =   4080
      Picture         =   "chess.frx":1A87
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   6
      Left            =   4800
      Picture         =   "chess.frx":1F81
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   7
      Left            =   5520
      Picture         =   "chess.frx":256F
      Top             =   480
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   8
      Left            =   480
      Picture         =   "chess.frx":2A94
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   9
      Left            =   1200
      Picture         =   "chess.frx":2F45
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   10
      Left            =   1920
      Picture         =   "chess.frx":33F6
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   11
      Left            =   2640
      Picture         =   "chess.frx":38A7
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   12
      Left            =   3360
      Picture         =   "chess.frx":3D58
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   13
      Left            =   4080
      Picture         =   "chess.frx":4209
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   14
      Left            =   4800
      Picture         =   "chess.frx":46BA
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   15
      Left            =   5520
      Picture         =   "chess.frx":4B6B
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   31
      Left            =   5520
      Picture         =   "chess.frx":501C
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   30
      Left            =   4800
      Picture         =   "chess.frx":54A3
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   29
      Left            =   4080
      Picture         =   "chess.frx":59A4
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   28
      Left            =   3360
      Picture         =   "chess.frx":5E16
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   27
      Left            =   2640
      Picture         =   "chess.frx":62AB
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   26
      Left            =   1920
      Picture         =   "chess.frx":6745
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   25
      Left            =   1200
      Picture         =   "chess.frx":6BB7
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   24
      Left            =   480
      Picture         =   "chess.frx":70B8
      Top             =   5520
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   23
      Left            =   5520
      Picture         =   "chess.frx":753F
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   22
      Left            =   4800
      Picture         =   "chess.frx":798C
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   21
      Left            =   4080
      Picture         =   "chess.frx":7DD9
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   20
      Left            =   3360
      Picture         =   "chess.frx":8226
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   19
      Left            =   2640
      Picture         =   "chess.frx":8673
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   18
      Left            =   1920
      Picture         =   "chess.frx":8AC0
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   17
      Left            =   1200
      Picture         =   "chess.frx":8F0D
      Top             =   4800
      Width           =   720
   End
   Begin VB.Image pieces 
      Height          =   720
      Index           =   16
      Left            =   480
      Picture         =   "chess.frx":935A
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   63
      Left            =   5520
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   62
      Left            =   4800
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   61
      Left            =   4080
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   60
      Left            =   3360
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   59
      Left            =   2640
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   58
      Left            =   1920
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   57
      Left            =   1200
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   56
      Left            =   480
      Top             =   5520
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   55
      Left            =   5520
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   54
      Left            =   4800
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   53
      Left            =   4080
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   52
      Left            =   3360
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   51
      Left            =   2640
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   50
      Left            =   1920
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   49
      Left            =   1200
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   48
      Left            =   480
      Top             =   4800
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   47
      Left            =   5520
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   46
      Left            =   4800
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   45
      Left            =   4080
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   44
      Left            =   3360
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   43
      Left            =   2640
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   42
      Left            =   1920
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   41
      Left            =   1200
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   40
      Left            =   480
      Top             =   4080
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   39
      Left            =   5520
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   38
      Left            =   4800
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   37
      Left            =   4080
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   36
      Left            =   3360
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   35
      Left            =   2640
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   34
      Left            =   1920
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   33
      Left            =   1200
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   32
      Left            =   480
      Top             =   3360
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   31
      Left            =   5520
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   30
      Left            =   4800
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   29
      Left            =   4080
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   28
      Left            =   3360
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   27
      Left            =   2640
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   26
      Left            =   1920
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   25
      Left            =   1200
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   24
      Left            =   480
      Top             =   2640
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   23
      Left            =   5520
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   22
      Left            =   4800
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   21
      Left            =   4080
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   20
      Left            =   3360
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   19
      Left            =   2640
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   18
      Left            =   1920
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   17
      Left            =   1200
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   16
      Left            =   480
      Top             =   1920
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   15
      Left            =   5520
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   14
      Left            =   4800
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   13
      Left            =   4080
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   12
      Left            =   3360
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   11
      Left            =   2640
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   10
      Left            =   1920
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   9
      Left            =   1200
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   8
      Left            =   480
      Top             =   1200
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   7
      Left            =   5520
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   6
      Left            =   4800
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   5
      Left            =   4080
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   4
      Left            =   3360
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   3
      Left            =   2640
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   2
      Left            =   1920
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   720
      Index           =   1
      Left            =   1200
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape board 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   720
      Index           =   0
      Left            =   480
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "chess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim P2M As Integer

Private Sub pieces_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
P2M = Index
pieces(Index).Drag 1
End Sub

Private Sub Move_Piece(source As Control, X As Single, Y As Single)
Dim w As Integer, h As Integer, i As Integer
For i = pieces.LBound To pieces.UBound
    If Not i = P2M Then
        pieces(i).Enabled = False
    End If
Next i
w = board(0).Left
h = board(0).Top
Do Until w > X
    w = w + 720
Loop
Do Until h > Y
    h = h + 720
Loop
w = w - 720
h = h - 720
'Insert function for move legality
Call source.Move(w, h)
For i = pieces.LBound To pieces.UBound
    pieces(i).Enabled = True
    If pieces(i).Top = pieces(P2M).Top And pieces(i).Left = pieces(P2M).Left And Not P2M = i Then
        pieces(i).Visible = False
        pieces(i).Enabled = False
    End If
Next i


End Sub


Private Sub Form_DragDrop(source As Control, X As Single, Y As Single)
Move_Piece source, X, Y
End Sub


