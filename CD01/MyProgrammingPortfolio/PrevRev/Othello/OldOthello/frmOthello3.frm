VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   8
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   7
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   6
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   5
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   4
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   3
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   2
      Left            =   1560
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   1
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Index           =   0
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim whosTurn As Boolean

Private Sub Picture1_Click(Index As Integer)
If whosTurn = True Then
Picture1(i).Picture = imgWhite.Picture
   i = i + 1
ElseIf whosTurn = False Then
Picture1(i).Picture = imgBlack.Picture
   i = i + 1
End If
turn
End Sub


Private Sub turn()
    If whosTurn = True Then
    whosTurn = False
    Else: whosTurn = True
    End If
End Sub

Private Sub Form_Load()
whosTurn = True
For i = 0 To 99
Picture1(i).Picture = imgBlank.Picture
Next i
End Sub

Public Function rowsel()
Select Case rosel
Case 0 To 9
Case 10 To 19
Case 20 To 29
Case 30 To 39
Case 40 To 49
Case 50 To 59
Case 60 To 69
Case 70 To 79
Case 80 To 89
Case 90 To 99


End Function
