Attribute VB_Name = "move"
Option Explicit
Global V1 As Integer, V2 As Integer
Global X1 As Integer, X2 As Integer
Global Y1 As Integer, Y2 As Integer

Global Piece_End_Array As Integer, Piece_Pointer As Integer
Global Piece(0 To 5000) As Chess_Matrix

Type Chess_Matrix
    xy(0 To 7, 0 To 7) As Integer
End Type
Dim InPlay As Boolean
Dim BoardColor As String

Function Setup_Board()
Dim i As Integer, j As Integer, counter As Integer
counter = -1
For i = 0 To 7
    For j = 0 To 7
    counter = counter + 1
    Chess1.board(counter).Picture = Chess1.pieces(Piece(Piece_Pointer).xy(i, j))
    Next j
Next i
End Function

Function Move_Piece(Index As Integer)
If InPlay = True Then
    InPlay = False
    V2 = Index
    Chess1.board(V1).BackColor = BoardColor
    If Not V1 = V2 Then
'        Move_Piece V1, V2
    End If
Else
    If Not Chess1.board(Index).Picture = 0 Then
    InPlay = True
    V1 = Index
    BoardColor = Chess1.board(Index).BackColor
    Chess1.board(Index).BackColor = vbGreen
    End If
End If


X1 = -8 * Y1 + V1
Y1 = Int(V1 / 8)
X2 = -8 * Y2 + V2
Y2 = Int(V2 / 8)
End Function
Function Chess_Rules()






End Function
Function New_Game()
End Function

Function Update_Board(Move_Number As Integer)
End Function
