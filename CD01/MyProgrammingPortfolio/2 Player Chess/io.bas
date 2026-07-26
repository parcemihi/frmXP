Attribute VB_Name = "io"
Option Explicit
Global Board_Layout(0 To 63) As Integer
Global Board_Layout1(0 To 5000) As board_area
Type board_area
    a(7) As Integer '1
    b(7) As Integer '2
    c(7) As Integer '3
    d(7) As Integer '4
    e(7) As Integer '5
    f(7) As Integer '6
    g(7) As Integer '7
    h(7) As Integer '8
End Type

Dim WQR_ As PictureBox

Function SetUpBoard()
Dim i As Integer


For i = 0 To 7
Chess.board(i).Picture = Chess.pieces(Board_Layout1(0).a(i))
Chess.board(i + 8).Picture = Chess.pieces(Board_Layout1(0).b(i))
Chess.board(i + 16).Picture = Chess.pieces(Board_Layout1(0).c(i))
Chess.board(i + 24).Picture = Chess.pieces(Board_Layout1(0).d(i))
Chess.board(i + 32).Picture = Chess.pieces(Board_Layout1(0).e(i))
Chess.board(i + 40).Picture = Chess.pieces(Board_Layout1(0).f(i))
Chess.board(i + 48).Picture = Chess.pieces(Board_Layout1(0).g(i))
Chess.board(i + 56).Picture = Chess.pieces(Board_Layout1(0).h(i))
Next i









End Function

Function NewGame()
Chess.CommonDialog1.FileName = ""
Game_Default_Values
End Function

Function SaveGame()
On Error GoTo ErrorHandler
Open Chess.CommonDialog1.FileName For Append As #1
Write #1, Board_Layout(0), Board_Layout(1), Board_Layout(2), Board_Layout(3), Board_Layout(4), Board_Layout(5), Board_Layout(6), Board_Layout(7), Board_Layout(8), Board_Layout(9), Board_Layout(10), Board_Layout(11), Board_Layout(12), Board_Layout(13), Board_Layout(14), Board_Layout(15), Board_Layout(16), Board_Layout(17), Board_Layout(18), Board_Layout(19), Board_Layout(20), Board_Layout(21), Board_Layout(22), Board_Layout(23), Board_Layout(24), Board_Layout(25), Board_Layout(26), Board_Layout(27), Board_Layout(28), Board_Layout(29), Board_Layout(30), Board_Layout(31), _
          Board_Layout(32), Board_Layout(33), Board_Layout(34), Board_Layout(35), Board_Layout(36), Board_Layout(37), Board_Layout(38), Board_Layout(39), Board_Layout(40), Board_Layout(41), Board_Layout(42), Board_Layout(43), Board_Layout(44), Board_Layout(45), Board_Layout(46), Board_Layout(47), Board_Layout(48), Board_Layout(49), Board_Layout(50), Board_Layout(51), Board_Layout(52), Board_Layout(53), Board_Layout(54), Board_Layout(55), Board_Layout(56), Board_Layout(57), Board_Layout(58), Board_Layout(59), Board_Layout(60), Board_Layout(61), Board_Layout(62), Board_Layout(63)
Close #1
ErrorHandler:
End Function

Function LoadGame()
On Error GoTo ErrorHandler
Open Chess.CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
'    Input #1
Board_Layout1(0).a (0): Board_Layout1(0).a (1): Board_Layout1(0).a (2): Board_Layout1(0).a (3): Board_Layout1(0).a (4): Board_Layout1(0).a (5): Board_Layout1(0).a (6): Board_Layout1(0).a (7)
Board_Layout1(0).b (0): Board_Layout1(0).b (1): Board_Layout1(0).b (2): Board_Layout1(0).b (3): Board_Layout1(0).b (4): Board_Layout1(0).b (5): Board_Layout1(0).b (6): Board_Layout1(0).b (7)
Board_Layout1(0).c (0): Board_Layout1(0).c (1): Board_Layout1(0).c (2): Board_Layout1(0).c (3): Board_Layout1(0).c (4): Board_Layout1(0).c (5): Board_Layout1(0).c (6): Board_Layout1(0).c (7)
Board_Layout1(0).d (0): Board_Layout1(0).d (1): Board_Layout1(0).d (2): Board_Layout1(0).d (3): Board_Layout1(0).d (4): Board_Layout1(0).d (5): Board_Layout1(0).d (6): Board_Layout1(0).d (7)
Board_Layout1(0).e (0): Board_Layout1(0).e (1): Board_Layout1(0).e (2): Board_Layout1(0).e (3): Board_Layout1(0).e (4): Board_Layout1(0).e (5): Board_Layout1(0).e (6): Board_Layout1(0).e (7)
Board_Layout1(0).f (0): Board_Layout1(0).f (1): Board_Layout1(0).f (2): Board_Layout1(0).f (3): Board_Layout1(0).f (4): Board_Layout1(0).f (5): Board_Layout1(0).f (6): Board_Layout1(0).f (7)
Board_Layout1(0).g (0): Board_Layout1(0).g (1): Board_Layout1(0).g (2): Board_Layout1(0).g (3): Board_Layout1(0).g (4): Board_Layout1(0).g (5): Board_Layout1(0).g (6): Board_Layout1(0).g (7)
Board_Layout1(0).h (0): Board_Layout1(0).h (1): Board_Layout1(0).h (2): Board_Layout1(0).h (3): Board_Layout1(0).h (4): Board_Layout1(0).h (5): Board_Layout1(0).h (6): Board_Layout1(0).h (7)
    Loop
Close #1
ErrorHandler:
End Function

Function Game_Default_Values()
Board_Layout1(0).a(0) = 1
Board_Layout1(0).a(1) = 2
Board_Layout1(0).a(2) = 3
Board_Layout1(0).a(3) = 4
Board_Layout1(0).a(4) = 5
Board_Layout1(0).a(5) = 3
Board_Layout1(0).a(6) = 2
Board_Layout1(0).a(7) = 1
Board_Layout1(0).b(0) = 6
Board_Layout1(0).b(1) = 6
Board_Layout1(0).b(2) = 6
Board_Layout1(0).b(3) = 6
Board_Layout1(0).b(4) = 6
Board_Layout1(0).b(5) = 6
Board_Layout1(0).b(6) = 6
Board_Layout1(0).b(7) = 6

Dim i As Integer
For i = 0 To 7
Board_Layout1(0).c(i) = 0
Board_Layout1(0).d(i) = 0
Board_Layout1(0).e(i) = 0
Board_Layout1(0).f(i) = 0
Next i

Board_Layout1(0).g(0) = 12
Board_Layout1(0).g(1) = 12
Board_Layout1(0).g(2) = 12
Board_Layout1(0).g(3) = 12
Board_Layout1(0).g(4) = 12
Board_Layout1(0).g(5) = 12
Board_Layout1(0).g(6) = 12
Board_Layout1(0).g(7) = 12
Board_Layout1(0).h(0) = 7
Board_Layout1(0).h(1) = 8
Board_Layout1(0).h(2) = 9
Board_Layout1(0).h(3) = 10
Board_Layout1(0).h(4) = 11
Board_Layout1(0).h(5) = 9
Board_Layout1(0).h(6) = 8
Board_Layout1(0).h(7) = 7
End Function
