VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   7
      Left            =   6840
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   6
      Left            =   5880
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   5
      Left            =   4920
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   4
      Left            =   3960
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   3
      Left            =   3000
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   2
      Left            =   2040
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   1
      Left            =   1080
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox board 
      Height          =   975
      Index           =   0
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin PicClip.PictureClip pieces 
      Left            =   1800
      Top             =   3000
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub Form_Click()
'pieces.ClipX = 1
'pieces.ClipY = 1
'board(0).Picture = pieces.Clip
'End Sub
Dim P2M As Integer

Private Sub board_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
pieces(Index).Drag 1

End Sub
Private Sub board_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
Move_Piece Source, X, Y
End Sub

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
Call Source.Move(w, h)
For i = pieces.LBound To pieces.UBound
    pieces(i).Enabled = True
    If pieces(i).Top = pieces(P2M).Top And pieces(i).Left = pieces(P2M).Left And Not P2M = i Then
        pieces(i).Visible = False
        pieces(i).Enabled = False
    End If
Next i






Private Sub Form_Load()
pieces.Picture = LoadPicture(App.Path & "\images\pieces2.bmp")
   Dim SaveMode As Integer
   board(0).ScaleMode = 3
'   SaveMode = board(0).ScaleMode
   pieces.ClipX = 100
   pieces.ClipY = 100
   pieces.ClipWidth = 100
   pieces.ClipHeight = 100
   pieces.StretchX = board(0).ScaleWidth
   pieces.StretchY = board(0).ScaleHeight
   board(0).Picture = pieces.Clip
'   board(0).ScaleMode = SaveMode
End Sub
