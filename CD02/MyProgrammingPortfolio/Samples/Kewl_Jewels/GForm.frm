VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kewl Jewels! - by Simon Price"
   ClientHeight    =   5184
   ClientLeft      =   0
   ClientTop       =   228
   ClientWidth     =   3816
   Icon            =   "GForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Display 
      BackColor       =   &H00FFFFFF&
      Height          =   5088
      Left            =   120
      ScaleHeight     =   420
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   0
      Width           =   2928
      Begin VB.PictureBox Sprites 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1212
         Left            =   720
         ScaleHeight     =   101
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Timer MoveT 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   120
      End
      Begin VB.PictureBox PB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   5088
         Left            =   720
         ScaleHeight     =   420
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   2928
      End
      Begin VB.Timer LevelT 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   120
         Top             =   600
      End
   End
   Begin VB.PictureBox ColumnPic 
      Height          =   1128
      Left            =   3240
      ScaleHeight     =   1080
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   0
      Width           =   408
   End
   Begin VB.Label LevelL 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   252
      Left            =   3120
      TabIndex        =   4
      Top             =   2640
      Width           =   612
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Level"
      Height          =   252
      Left            =   3120
      TabIndex        =   3
      Top             =   2400
      Width           =   612
   End
   Begin VB.Label ScoreL 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   252
      Left            =   3120
      TabIndex        =   2
      Top             =   1680
      Width           =   612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Score"
      Height          =   252
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   612
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuDifficulty 
      Caption         =   "&Difficulty"
      Begin VB.Menu mnuEasy 
         Caption         =   "&Easy"
         Checked         =   -1  'True
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuMedium 
         Caption         =   "&Medium"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuHard 
         Caption         =   "&Hard"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Key As Byte
Dim GameOver As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Key = KeyCode
Select Case Key
  Case vbKeyLeft
    If MoveLeft Then
    Else
    End If
  Case vbKeyRight
    If MoveRight Then
    Else
    End If
  Case vbKeyDown
    MoveT.Interval = 10
  Case vbKeyUp
    RotateColumn
  Case vbKeyEscape
    End
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Key = vbKeyDown Then MoveT.Interval = MoveTime
Key = NOWT
End Sub

Private Sub Form_Load()
Randomize Timer
MoveTime = 300
Sprites = LoadPicture(App.Path & "\Resources\Graphix\Jewels.bmp")
Display = LoadPicture(App.Path & "\Resources\Graphix\Intro.jpg")
Show
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
Key = vbKeyEscape
End Sub

Private Sub LevelT_Timer()
Level = Level + 1
LevelL = Level
If MoveT.Interval > 10 Then MoveT.Interval = MoveT.Interval / 1.25
End Sub

Private Sub mnuAbout_Click()
MsgBox "Kewl Jewels - version 1.0 - by Simon Price - Email : Si@VBgames.co.uk - Website : www.VBgames.co.uk ", vbInformation, "About Kewl Jewels"
End Sub

Private Sub mnuEasy_Click()
Diff = 0
mnuEasy.Checked = True
mnuMedium.Checked = False
mnuHard.Checked = False
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Sub MainLoop()
Do
DoEvents
DrawAll
Animate
If GameOver Then Exit Do
Loop Until Key = vbKeyEscape

MoveT.Enabled = False
LevelT.Enabled = False
MsgBox "Game Over! You scored " & Score & " points and reached level " & Level & "!", vbInformation, "Game Over!"
Display = LoadPicture(App.Path & "\Resources\Graphix\Intro.jpg")

End Sub

Public Sub DoScore()
Score = Score + ChainReaction
ScoreL = Score
End Sub

Private Sub mnuHard_Click()
Diff = 2
mnuEasy.Checked = False
mnuMedium.Checked = False
mnuHard.Checked = True
End Sub

Private Sub mnuMedium_Click()
Diff = 1
mnuEasy.Checked = False
mnuMedium.Checked = True
mnuHard.Checked = False
End Sub

Private Sub mnuNewGame_Click()
Dim x, y As Integer
CreateColumn
CreateColumn
MoveT.Enabled = True
LevelT.Enabled = True
Score = 0
Level = 1
ScoreL = "0"
LevelL = "1"
GameOver = False
Key = NOWT
For x = 0 To XTILES
For y = -3 To YTILES
  Tile(x, y) = NOWT
Next
Next
Select Case Diff
 Case 0
   MoveT.Interval = 500
 Case 1
   MoveT.Interval = 350
 Case 2
   MoveT.Interval = 150
End Select
MainLoop
End Sub

Private Sub MoveT_Timer()
On Error Resume Next
Dim Jewel As Byte

If MoveColumn = STOPPED Then
sndPlaySound App.Path & "\Resources\SoundFX\Land.wav", &H1
  'column stopped, so create new one
  If Special Then
    sndPlaySound App.Path & "\Resources\SoundFX\Special.wav", 0
    Tile(C.x, C.y + Special - 1) = VANISH
    If Special = 3 Then
      Jewel = Tile(C.x, C.y + 3)
    Else
      Jewel = C.Jewel(Special)
    End If
    If Jewel Then
      DeleteAllJewelsOfType Jewel
    End If
  End If
  DoScore
  If ItIsGameOver Then GameOver = True
  CreateColumn
Else
End If

End Sub

Sub DrawAll()
Dim x, y As Integer
'go through each tile, drawing on backbuffer
PB.Cls
For x = 0 To XTILES
For y = 0 To YTILES
    Select Case Tile(x, y)
    Case 1 To 8
      BitBlt PB.hdc, x * 30, y * 30, 30, 30, Sprites.hdc, Frame * 30, (Tile(x, y) - 1) * 30, vbSrcCopy
    Case VANISH To -1
      BitBlt PB.hdc, x * 30, y * 30, 30, 30, Sprites.hdc, Tile(x, y) * -30, 240, vbSrcCopy
    End Select
Next
Next
'copy into view
BitBlt Display.hdc, 0, 0, 240, 420, PB.hdc, 0, 0, vbSrcCopy
End Sub

Public Sub CreateColumn()
Dim i, y As Integer
Special = Special2
Special2 = 0
'swap active column
C = C2
'position new column
C2.x = Int(Rnd * 2) + 3
C2.y = -3
'choose new jewels
For i = 0 To 2
TryAgain:
    If Int(Rnd * 50) Then
        C2.Jewel(i) = Int(Rnd * 6) + 2
    Else
        '1 in 50 jewels are special
        If Special2 Then GoTo TryAgain
        C2.Jewel(i) = CLEAR
        Special2 = i + 1
    End If
Next
'paint new column in preview box
For y = 0 To 2
  BitBlt ColumnPic.hdc, 0, y * 30, 30, 30, Sprites.hdc, Frame * 30, (C2.Jewel(y) - 1) * 30, vbSrcCopy
Next
End Sub

