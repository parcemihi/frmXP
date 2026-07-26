Attribute VB_Name = "io"
Option Explicit

Function Load_Default()
Dim i As Integer, j As Integer
Open App.Path & "\default.dat" For Input As #1
        For i = 0 To 7
            For j = 0 To 7
                Input #1, Piece(0).xy(i, j)
            Next j
        Next i
Close #1
End Function


Function Load_Moves_Log()
Dim i As Integer, j As Integer
Piece_End_Array = 0
Open App.Path & "\moves.log" For Input As #1
    Do Until EOF(1)
    Piece_End_Array = Piece_End_Array + 1
        For i = 0 To 7
            For j = 0 To 7
                Input #1, Piece(Piece_End_Array).xy(i, j)
            Next j
        Next i
    Loop
Close #1
End Function

Function Save_Moves_Log()
Dim i As Integer
Open App.Path & "\moves.log" For Output As #1
Close #1
End Function

