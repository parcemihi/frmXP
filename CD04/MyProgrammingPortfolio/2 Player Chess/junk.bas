Attribute VB_Name = "junk"
Option Explicit
'Dim global_left As Integer
'Dim global_top As Integer
'
'Dim playlist(0 To 5000) As game
'Private Type game
'    piece(0 To 31) As Integer
'End Type
'
'Dim move_end As Integer
'Dim move_no As Integer
Private Sub Get_New_Layout()
Dim global_width As Integer





End Sub

Private Function open_playlist()
Dim i As Integer
Open App.Path & "\playlist.log" For Input As #1
Do Until EOF(1)
    Input #1, move_no
    For i = 0 To 31
        Input #1, playlist(move_no).piece(i)
    Next i
Loop
move_end = move_no
Close #1
End Function

Private Function write_playlist()
Dim i As Integer

For i = 0 To 15
playlist.piece(i) = i
Next i

For i = 0 To 15
playlist.piece(i + 16) = i + 48
Next i



Open App.Path & "\playlist.log" For Output As #1
    For i = 0 To move_end
    Write #1, i, playlist.piece(0), playlist.piece(1), playlist.piece(2), playlist.piece(3), playlist.piece(4), playlist.piece(5), playlist.piece(6), playlist.piece(7), playlist.piece(8), playlist.piece(9), _
                 playlist.piece(10), playlist.piece(11), playlist.piece(12), playlist.piece(13), playlist.piece(14), playlist.piece(15), playlist.piece(16), playlist.piece(17), playlist.piece(18), playlist.piece(19), _
                 playlist.piece(20), playlist.piece(21), playlist.piece(22), playlist.piece(23), playlist.piece(24), playlist.piece(25), playlist.piece(26), playlist.piece(27), playlist.piece(28), playlist.piece(29), _
                 playlist.piece(30), playlist.piece(31)
    Next i
Close #1
End Function


