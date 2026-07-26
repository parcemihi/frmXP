Attribute VB_Name = "KewlJewels"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'tiles that make up the game screen
Public Const XTILES = 7
Public Const YTILES = 13
Public Tile(0 To XTILES, -3 To YTILES) As Integer
Public VTile(0 To XTILES, 0 To YTILES) As Boolean
Public DTile(0 To XTILES, 0 To YTILES) As Boolean


'frame of animation
Public Frame As Byte

'jewel constants
Public Const VANISH = -4
Public Const NOWT = 0
Public Const CLEAR = 1
Public Const CYAN = 2
Public Const GREEN = 3
Public Const MAGENTA = 4
Public Const YELLOW = 5
Public Const BLUE = 6
Public Const RED = 7
Public Const ORANGE = 8

Public Score As Long
Public Level As Integer
Public Special As Byte
Public Special2 As Byte
Public MoveTime As Integer
Public Diff As Byte

Public Type tColumn
  x As Byte
  y As Integer
  Jewel(0 To 2) As Byte
End Type

Public C As tColumn
Public C2 As tColumn

Public Const STOPPED = 1
Public Const NOPE = 1

Public Sub RotateColumn()
'rotates the jewels in a column
Dim TempJewel As Byte
TempJewel = C.Jewel(0)
C.Jewel(0) = C.Jewel(2)
C.Jewel(2) = C.Jewel(1)
C.Jewel(1) = TempJewel
If Special Then
If Special = 3 Then
  Special = 1
Else
  Special = Special + 1
End If
End If
End Sub

Public Function MoveLeft() As Byte
Dim y As Integer
On Error Resume Next
If C.x = 0 Then
  'on edge, can't move
  MoveLeft = NOPE
  Exit Function
End If
For y = 0 To 2
  If Tile(C.x - 1, C.y + y) Then
    'cannot move left (tile in way)
    MoveLeft = NOPE
    Exit Function
  End If
Next
'passed all tests, so move left
'clear away old pos
For y = -1 To 2
  Tile(C.x, C.y + y) = NOWT
Next
'move left
C.x = C.x - 1
'fill in new pos
For y = 0 To 2
  Tile(C.x, C.y + y) = C.Jewel(y)
Next
End Function

Public Function MoveRight() As Byte
Dim y As Integer
On Error Resume Next
If C.x = XTILES Then
  'on edge, can't move
  MoveRight = NOPE
  Exit Function
End If
For y = 0 To 2
  If Tile(C.x + 1, C.y + y) Then
    'cannot move right (tile in way)
    MoveRight = NOPE
    Exit Function
  End If
Next
'passed all tests, so move right
'clear away old pos
For y = -1 To 2
  Tile(C.x, C.y + y) = NOWT
Next
'move left
C.x = C.x + 1
'fill in new pos
For y = 0 To 2
  Tile(C.x, C.y + y) = C.Jewel(y)
Next
End Function

Public Function MoveColumn() As Byte
Dim y As Integer
On Error Resume Next
For y = 0 To 2
  Tile(C.x, C.y + y) = C.Jewel(y)
Next
'move column down, if it's path is clear
If Tile(C.x, C.y + 3) Or C.y > YTILES - 2 Then
  'delete old jewel
  Tile(C.x, C.y - 1) = NOWT
  MoveColumn = 1
Else
  'OK, shuffle down
  For y = 0 To 2
    Tile(C.x, C.y + y) = C.Jewel(y)
  Next
  'delete old jewel
  Tile(C.x, C.y - 1) = NOWT
  'record new column position
  C.y = C.y + 1
End If
End Function

Public Function ChainReaction() As Integer
Dim Round, i, ChainLength As Byte

Do

ChainLength = ChainLength + 1

'check which jewels should disappear
Round = Check4Rows

'animate them vanishing
If Round Then sndPlaySound App.Path & "\Resources\SoundFX\Vanish.wav", &H1
For i = 1 To 4
    Animate
    Sleep 100
    GForm.DrawAll
Next

'drop the jewels that were above those
'which have just disappeared
Do
  If DropJewels Then
    Animate
    GForm.DrawAll
  Else
    Exit Do
  End If
Loop

'check if the chain reaction should
'continue or not
If Round Then
  ChainReaction = ChainReaction + Round * Round * ChainLength * ChainLength
Else
  Exit Do
End If

Loop

End Function

Public Function Check4Rows() As Byte
Dim x, y As Integer
'checks and makes lines of jewels vanish
'empty buffer
On Error Resume Next
For x = 0 To XTILES
For y = 0 To YTILES
  VTile(x, y) = False
Next
Next

'check jewels
For x = 0 To XTILES
For y = 0 To YTILES
  Check4HRows x + 1 - 1, y + 1 - 1
  Check4VRows x + 1 - 1, y + 1 - 1
  Check4D1Rows x + 1 - 1, y + 1 - 1
  Check4D2Rows x + 1 - 1, y + 1 - 1
Next
Next

'count up jewels captured and make them vanish
For x = 0 To XTILES
For y = 0 To YTILES
  If VTile(x, y) Then
    Check4Rows = Check4Rows + 1
    Tile(x, y) = VANISH
  End If
Next
Next

End Function

Public Sub Check4HRows(x As Byte, y As Byte)
On Error GoTo StopIt
Dim xx, xxx As Integer
Dim Jewel As Byte

Jewel = Tile(x, y)
If Jewel = NOWT Then Exit Sub
xx = x

Do
  If Tile(xx, y) <> Jewel Then Exit Do
  xx = xx + 1
Loop
StopIt:
If xx - x > 2 Then
    For xxx = x To xx - 1
      VTile(xxx, y) = True
    Next
End If

End Sub

Public Sub Check4VRows(x As Byte, y As Byte)
On Error GoTo StopIt
Dim yy, yyy As Integer
Dim Jewel As Byte

Jewel = Tile(x, y)
If Jewel = NOWT Then Exit Sub
yy = y

Do
  If Tile(x, yy) <> Jewel Then Exit Do
  yy = yy + 1
Loop
StopIt:
If yy - y > 2 Then
    For yyy = y To yy - 1
      VTile(x, yyy) = True
    Next
End If

End Sub

Public Sub Check4D1Rows(x As Byte, y As Byte)
On Error GoTo StopIt
Dim i, ii As Integer
Dim Jewel As Byte

Jewel = Tile(x, y)
If Jewel = NOWT Then Exit Sub
i = 0
Do
  If Tile(x + i, y + i) <> Jewel Then Exit Do
  i = i + 1
Loop
StopIt:
If i > 2 Then
    For ii = 0 To i - 1
      VTile(x + ii, y + ii) = True
    Next
End If

End Sub

Public Sub Check4D2Rows(x As Byte, y As Byte)
On Error GoTo StopIt
Dim i, ii As Integer
Dim Jewel As Byte

Jewel = Tile(x, y)
If Jewel = NOWT Then Exit Sub
i = 0
Do
  If Tile(x + i, y - i) <> Jewel Then Exit Do
  i = i + 1
Loop
StopIt:
If i > 2 Then
    For ii = 0 To i - 1
      VTile(x + ii, y - ii) = True
    Next
End If

End Sub




Public Function DropJewels() As Boolean
Dim x, y, yy As Integer
'drops tiles and records them
For x = 0 To XTILES
For y = YTILES - 1 To 0 Step -1
   If Tile(x, y) Then
      If Tile(x, y + 1) = NOWT Then
         Tile(x, y + 1) = Tile(x, y)
         Tile(x, y) = NOWT
         DropJewels = True
      End If
   End If
Next
Next
End Function

Public Sub Animate()
Dim x, y As Integer
'animate each vanishing tile
For x = 0 To XTILES
For y = -3 To YTILES
  If Tile(x, y) < NOWT Then Tile(x, y) = Tile(x, y) + 1
Next
Next

'move to next frame of animation
If Frame = 3 Then Frame = 0 Else Frame = Frame + 1
End Sub

Public Sub DeleteAllJewelsOfType(Jewel As Byte)
'delete one sort of jewel
For x = 0 To XTILES
For y = 0 To YTILES
    If Tile(x, y) = Jewel Then Tile(x, y) = VANISH
Next
Next
sndPlaySound App.Path & "\Resources\SoundFX\Vanish.wav", &H1
'animate them vanishing
For i = 1 To 4
    Animate
    Sleep 100
    GForm.DrawAll
Next
'drop the jewels that were above those
'which have just disappeared
Do
  If DropJewels Then
    Animate
    GForm.DrawAll
  Else
    Exit Do
  End If
Loop
End Sub

Public Function ItIsGameOver() As Boolean
Dim x, y As Integer

For x = 0 To XTILES
For y = -3 To -1
  If Tile(x, y) Then ItIsGameOver = True
Next
Next
End Function

