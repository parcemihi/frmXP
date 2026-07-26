Attribute VB_Name = "io"
Option Explicit
Global Board_Layout(0 To 63) As Integer

Function SetUpBoard()
Dim i As Integer
For i = 0 To 63
    Chess.board(i).Picture = Chess.pieces(Board_Layout(i)).Picture
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
      Input #1, Board_Layout(0), Board_Layout(1), Board_Layout(2), Board_Layout(3), Board_Layout(4), Board_Layout(5), Board_Layout(6), Board_Layout(7), Board_Layout(8), Board_Layout(9), Board_Layout(10), Board_Layout(11), Board_Layout(12), Board_Layout(13), Board_Layout(14), Board_Layout(15), Board_Layout(16), Board_Layout(17), Board_Layout(18), Board_Layout(19), Board_Layout(20), Board_Layout(21), Board_Layout(22), Board_Layout(23), Board_Layout(24), Board_Layout(25), Board_Layout(26), Board_Layout(27), Board_Layout(28), Board_Layout(29), Board_Layout(30), Board_Layout(31), _
                Board_Layout(32), Board_Layout(33), Board_Layout(34), Board_Layout(35), Board_Layout(36), Board_Layout(37), Board_Layout(38), Board_Layout(39), Board_Layout(40), Board_Layout(41), Board_Layout(42), Board_Layout(43), Board_Layout(44), Board_Layout(45), Board_Layout(46), Board_Layout(47), Board_Layout(48), Board_Layout(49), Board_Layout(50), Board_Layout(51), Board_Layout(52), Board_Layout(53), Board_Layout(54), Board_Layout(55), Board_Layout(56), Board_Layout(57), Board_Layout(58), Board_Layout(59), Board_Layout(60), Board_Layout(61), Board_Layout(62), Board_Layout(63)
    Loop
Close #1
ErrorHandler:
End Function

Function Game_Default_Values()
Dim i As Integer
Open App.Path & "\default.keg" For Input As #1
Input #1, Board_Layout(0), Board_Layout(1), Board_Layout(2), Board_Layout(3), Board_Layout(4), Board_Layout(5), Board_Layout(6), Board_Layout(7), _
          Board_Layout(8), Board_Layout(9), Board_Layout(10), Board_Layout(11), Board_Layout(12), Board_Layout(13), Board_Layout(14), Board_Layout(15), _
          Board_Layout(16), Board_Layout(17), Board_Layout(18), Board_Layout(19), Board_Layout(20), Board_Layout(21), Board_Layout(22), Board_Layout(23), _
          Board_Layout(24), Board_Layout(25), Board_Layout(26), Board_Layout(27), Board_Layout(28), Board_Layout(29), Board_Layout(30), Board_Layout(31), _
          Board_Layout(32), Board_Layout(33), Board_Layout(34), Board_Layout(35), Board_Layout(36), Board_Layout(37), Board_Layout(38), Board_Layout(39), _
          Board_Layout(40), Board_Layout(41), Board_Layout(42), Board_Layout(43), Board_Layout(44), Board_Layout(45), Board_Layout(46), Board_Layout(47), _
          Board_Layout(48), Board_Layout(49), Board_Layout(50), Board_Layout(51), Board_Layout(52), Board_Layout(53), Board_Layout(54), Board_Layout(55), _
          Board_Layout(56), Board_Layout(57), Board_Layout(58), Board_Layout(59), Board_Layout(60), Board_Layout(61), Board_Layout(62), Board_Layout(63)
Close #1

End Function
