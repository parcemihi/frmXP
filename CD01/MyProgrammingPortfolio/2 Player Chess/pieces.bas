Attribute VB_Name = "Game"
Option Explicit
'Define the positions of each piece
Type pawns
    pawn As Integer
End Type
Type Pieces
    Crown As Integer
    Bishop As Integer
    Knight As Integer
    Rook As Integer
    Crowns As pawns
    Bishops As pawns
    Knights As pawns
    Rooks As pawns
End Type
Type Side
    Kings As Pieces
    Queens As Pieces
End Type
Type Teams
    White As Side
    Black As Side
End Type
Global Game As Teams

Function Game_Default_Values()
With Game: With .White
.Queens.Rook = 0: .Queens.Knight = 1: .Queens.Bishop = 2: .Queens.Crown = 3: .Kings.Crown = 4: .Kings.Bishop = 5: .Kings.Knight = 6: .Kings.Rook = 7: .Queens.Rooks.pawn = 8: .Queens.Knights.pawn = 9: .Queens.Bishops.pawn = 10: .Queens.Crowns.pawn = 11: .Kings.Crowns.pawn = 12: .Kings.Bishops.pawn = 13: .Kings.Knights.pawn = 14: .Kings.Rooks.pawn = 15
End With: With .Black
.Queens.Rooks.pawn = 48: .Queens.Knights.pawn = 49: .Queens.Bishops.pawn = 50: .Queens.Crowns.pawn = 51: .Kings.Crowns.pawn = 52: .Kings.Bishops.pawn = 53: .Kings.Knights.pawn = 54: .Kings.Rooks.pawn = 55: .Queens.Rook = 56: .Queens.Knight = 57: .Queens.Bishop = 58: .Queens.Crown = 59: .Kings.Crown = 60: .Kings.Bishop = 61: .Kings.Knight = 62: .Kings.Rook = 63
End With: End With
End Function

Function SetUpBoard()
Dim ImagePath As String
ImagePath = App.Path & "\images\"
With Game
    With .White
        With .Queens
        Chess.board(.Rook).Picture = LoadPicture(ImagePath & "RW.gif")
        Chess.board(.Knight).Picture = LoadPicture(ImagePath & "NW.gif")
        Chess.board(.Bishop).Picture = LoadPicture(ImagePath & "BW.gif")
        Chess.board(.Crown).Picture = LoadPicture(ImagePath & "QW.gif")
        Chess.board(.Rooks.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        Chess.board(.Knights.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        Chess.board(.Bishops.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        Chess.board(.Crowns.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        End With
        With .Kings
        Chess.board(.Crown).Picture = LoadPicture(ImagePath & "KW.gif")
        Chess.board(.Bishop).Picture = LoadPicture(ImagePath & "BW.gif")
        Chess.board(.Knight).Picture = LoadPicture(ImagePath & "NW.gif")
        Chess.board(.Rook).Picture = LoadPicture(ImagePath & "RW.gif")
        Chess.board(.Crowns.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        Chess.board(.Bishops.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        Chess.board(.Knights.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        Chess.board(.Rooks.pawn).Picture = LoadPicture(ImagePath & "PW.gif")
        End With
    End With
    With .Black
        With .Queens
        Chess.board(.Rooks.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Knights.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Bishops.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Crowns.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Rook).Picture = LoadPicture(ImagePath & "RB.gif")
        Chess.board(.Knight).Picture = LoadPicture(ImagePath & "NB.gif")
        Chess.board(.Bishop).Picture = LoadPicture(ImagePath & "BB.gif")
        Chess.board(.Crown).Picture = LoadPicture(ImagePath & "QB.gif")
        End With
        With .Kings
        Chess.board(.Crowns.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Bishops.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Knights.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Rooks.pawn).Picture = LoadPicture(ImagePath & "PB.gif")
        Chess.board(.Crown).Picture = LoadPicture(ImagePath & "KB.gif")
        Chess.board(.Bishop).Picture = LoadPicture(ImagePath & "BB.gif")
        Chess.board(.Knight).Picture = LoadPicture(ImagePath & "NB.gif")
        Chess.board(.Rook).Picture = LoadPicture(ImagePath & "RB.gif")
        End With
    End With
End With
End Function

Function SaveGame()
Open Chess.CommonDialog1.FileName For Append As #1
    With Game
    Write #1, .White.Queens.Rook, .White.Queens.Knight, .White.Queens.Bishop, .White.Queens.Crown, .White.Kings.Crown, .White.Kings.Bishop, .White.Kings.Knight, .White.Kings.Rook, .White.Queens.Rooks.pawn, .White.Queens.Knights.pawn, .White.Queens.Bishops.pawn, .White.Queens.Crowns.pawn, .White.Kings.Crowns.pawn, .White.Kings.Bishops.pawn, .White.Kings.Knights.pawn, .White.Kings.Rooks.pawn, .Black.Queens.Rooks.pawn, .Black.Queens.Knights.pawn, .Black.Queens.Bishops.pawn, .Black.Queens.Crowns.pawn, .Black.Kings.Crowns.pawn, .Black.Kings.Bishops.pawn, .Black.Kings.Knights.pawn, .Black.Kings.Rooks.pawn, .Black.Queens.Rook, .Black.Queens.Knight, .Black.Queens.Bishop, .Black.Queens.Crown, .Black.Kings.Crown, .Black.Kings.Bishop, .Black.Kings.Knight, .Black.Kings.Rook
    End With
Close #1
End Function

Function NewGame()
Game_Default_Values
End Function

Function LoadGame()
On Error GoTo ErrorLog
Open Chess.CommonDialog1.FileName For Input As #1
    With Game
    Do Until EOF(1)
    Input #1, .White.Queens.Rook, .White.Queens.Knight, .White.Queens.Bishop, .White.Queens.Crown, .White.Kings.Crown, .White.Kings.Bishop, .White.Kings.Knight, .White.Kings.Rook, .White.Queens.Rooks.pawn, .White.Queens.Knights.pawn, .White.Queens.Bishops.pawn, .White.Queens.Crowns.pawn, .White.Kings.Crowns.pawn, .White.Kings.Bishops.pawn, .White.Kings.Knights.pawn, .White.Kings.Rooks.pawn, .Black.Queens.Rooks.pawn, .Black.Queens.Knights.pawn, .Black.Queens.Bishops.pawn, .Black.Queens.Crowns.pawn, .Black.Kings.Crowns.pawn, .Black.Kings.Bishops.pawn, .Black.Kings.Knights.pawn, .Black.Kings.Rooks.pawn, .Black.Queens.Rook, .Black.Queens.Knight, .Black.Queens.Bishop, .Black.Queens.Crown, .Black.Kings.Crown, .Black.Kings.Bishop, .Black.Kings.Knight, .Black.Kings.Rook
    Loop
    End With
Close #1

ErrorLog:

End Function

