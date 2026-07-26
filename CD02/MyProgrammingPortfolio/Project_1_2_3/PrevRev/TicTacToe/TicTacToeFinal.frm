VERSION 5.00
Begin VB.Form TicTacToe2 
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   1560
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Menu Game 
      Caption         =   "&Game"
      Begin VB.Menu ComputerPlay 
         Caption         =   "Computer Play"
         Checked         =   -1  'True
      End
      Begin VB.Menu Reset 
         Caption         =   "Reset"
      End
   End
End
Attribute VB_Name = "TicTacToe2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theArray(0 To 8) As String, pass As Integer, j As Integer, k As Integer
Dim Trncnt As Integer, Level As Integer, WinGame As Integer
Dim C(2, 2) As String, i As Integer, WhosTurn As Boolean

Private Sub Cmd_Click(Index As Integer)
If Cmd(Index).Caption = "" Then
    Cmd(Index).Caption = "O"
    Check
        For i = 0 To 8
        If Cmd(i).Caption = "X" Or Cmd(i).Caption = "O" Then
        Else
        ComputerGoes
        Exit For
        End If
        Next i
    Check
    End If

End Sub

Private Sub ComputerPlay_Click()
If ComputerPlay.Checked = True Then
ComputerPlay.Checked = False
Else: ComputerPlay.Checked = True
End If
ResetGame
End Sub

Private Sub ResetGame()
For i = 0 To 8
    Cmd(i).Caption = ""
Next i
If ComputerPlay.Checked = True Then
    ComputerGoes
'    Cmd(4).Caption = "X"
End If

End Sub


Private Sub Form_Load()

ComputerPlay.Checked = False

pass = 0
Trncnt = 0

End Sub

Public Function Check()

k = 0
    For i = 0 To 2
        For j = 0 To 2
        C(j, i) = Cmd(k).Caption
        k = k + 1
        Next j
    Next i
    If _
    C(0, 0) = "X" And C(0, 1) = "X" And C(0, 2) = "X" Or _
    C(1, 0) = "X" And C(1, 1) = "X" And C(1, 2) = "X" Or _
    C(2, 0) = "X" And C(2, 1) = "X" And C(2, 2) = "X" Or _
    C(0, 0) = "X" And C(1, 0) = "X" And C(2, 0) = "X" Or _
    C(0, 1) = "X" And C(1, 1) = "X" And C(2, 1) = "X" Or _
    C(0, 2) = "X" And C(1, 2) = "X" And C(2, 2) = "X" Or _
    C(0, 0) = "X" And C(1, 1) = "X" And C(2, 2) = "X" Or _
    C(2, 0) = "X" And C(1, 1) = "X" And C(0, 2) = "X" Then
    MsgBox ("X Wins")
    ResetGame
    ElseIf _
    C(0, 0) = "O" And C(0, 1) = "O" And C(0, 2) = "O" Or _
    C(1, 0) = "O" And C(1, 1) = "O" And C(1, 2) = "O" Or _
    C(2, 0) = "O" And C(2, 1) = "O" And C(2, 2) = "O" Or _
    C(0, 0) = "O" And C(1, 0) = "O" And C(2, 0) = "O" Or _
    C(0, 1) = "O" And C(1, 1) = "O" And C(2, 1) = "O" Or _
    C(0, 2) = "O" And C(1, 2) = "O" And C(2, 2) = "O" Or _
    C(0, 0) = "O" And C(1, 1) = "O" And C(2, 2) = "O" Or _
    C(2, 0) = "O" And C(1, 1) = "O" And C(0, 2) = "O" Then
    MsgBox ("O Wins")
    ResetGame
    ElseIf _
    (C(0, 0) = "X" Or C(0, 0) = "O") And (C(0, 1) = "X" Or C(0, 1) = "O") And _
    (C(0, 2) = "X" Or C(0, 2) = "O") And (C(1, 0) = "X" Or C(1, 0) = "O") And _
    (C(1, 1) = "X" Or C(1, 1) = "O") And (C(1, 2) = "X" Or C(1, 2) = "O") And _
    (C(2, 0) = "X" Or C(2, 0) = "O") And (C(2, 1) = "X" Or C(2, 1) = "O") And _
    (C(2, 2) = "X" Or C(2, 2) = "O") And (C(0, 0) = "X" Or C(0, 0) = "O") And _
    (C(1, 0) = "X" Or C(1, 0) = "O") And (C(2, 0) = "X" Or C(2, 0) = "O") And _
    (C(0, 1) = "X" Or C(0, 1) = "O") And (C(1, 1) = "X" Or C(1, 1) = "O") And _
    (C(2, 1) = "X" Or C(2, 1) = "O") And (C(0, 2) = "X" Or C(0, 2) = "O") And _
    (C(1, 2) = "X" Or C(1, 2) = "O") And (C(2, 2) = "X" Or C(2, 2) = "O") Then
    MsgBox ("Cats Game")
    ResetGame
    End If
End Function
Private Sub Reset_Click()
    ResetGame
End Sub
Public Function ComputerGoes()

C(0, 0) = Cmd(0).Caption: C(1, 0) = Cmd(1).Caption: C(2, 0) = Cmd(2).Caption
C(0, 1) = Cmd(3).Caption: C(1, 1) = Cmd(4).Caption: C(2, 1) = Cmd(5).Caption
C(0, 2) = Cmd(6).Caption: C(1, 2) = Cmd(7).Caption: C(2, 2) = Cmd(8).Caption
    
    If C(1, 1) = "" Then
    Cmd(4).Caption = "X"
    ElseIf C(0, 0) = "" Then ' Or C(1, 0) = "" Then
    Cmd(0).Caption = "X"
    ElseIf C(0, 0) = "" And (C(1, 0) = "X" And C(2, 0) = "X" Or C(0, 1) = "X" And C(0, 2) = "X" Or C(1, 1) = "X" And C(2, 2) = "X") Then   '//win for Square 1
    Cmd(0).Caption = "X"
    ElseIf C(1, 0) = "" And (C(1, 1) = "X" And C(1, 2) = "X" Or C(0, 0) = "X" And C(2, 0) = "X") Then   '//win for Square 2
    Cmd(1).Caption = "X"
    ElseIf C(2, 0) = "" And (C(0, 0) = "X" And C(1, 0) = "X" Or C(1, 1) = "X" And C(0, 2) = "X" Or C(2, 1) = "X" And C(2, 2) = "X") Then    '//win for Square 3
    Cmd(2).Caption = "X"
    ElseIf C(0, 1) = "" And (C(1, 1) = "X" And C(2, 1) = "X" Or C(0, 0) = "X" And C(0, 2) = "X") Then    '//win for Square 4
    Cmd(3).Caption = "X"
    ElseIf C(1, 1) = "" And (C(0, 0) = "X" And C(2, 2) = "X" Or C(1, 0) = "X" And C(1, 2) = "X" Or C(2, 0) = "X" And C(0, 2) = "X" Or C(0, 1) = "X" And C(2, 1) = "X") Then    '//win for Square 5
    Cmd(4).Caption = "X"
    ElseIf C(2, 1) = "" And (C(2, 0) = "X" And C(2, 2) = "X" Or C(0, 1) = "X" And C(1, 1) = "X") Then    '//win for Square 6
    Cmd(5).Caption = "X"
    ElseIf C(0, 2) = "" And (C(0, 0) = "X" And C(0, 1) = "X" Or C(2, 0) = "X" And C(1, 1) = "X" Or C(1, 2) = "X" And C(2, 2) = "X") Then    '//win for Square 7
    Cmd(6).Caption = "X"
    ElseIf C(1, 2) = "" And (C(1, 0) = "X" And C(1, 1) = "X" Or C(0, 2) = "X" And C(2, 2) = "X") Then    '//win for Square 8
    Cmd(7).Caption = "X"
    ElseIf C(2, 2) = "" And (C(0, 0) = "X" And C(1, 1) = "X" Or C(2, 0) = "X" And C(2, 1) = "X" Or C(0, 2) = "X" And C(1, 2) = "X") Then    '//win for Square 9
    Cmd(8).Caption = "X"
    ElseIf C(0, 0) = "" And (C(1, 0) = "O" And C(2, 0) = "O" Or C(0, 1) = "O" And C(0, 2) = "O" Or C(1, 1) = "O" And C(2, 2) = "O") Then   '//win for Square 1
    Cmd(0).Caption = "X"
    ElseIf C(1, 0) = "" And (C(1, 1) = "O" And C(1, 2) = "O" Or C(0, 0) = "O" And C(2, 0) = "O") Then   '//win for Square 2
    Cmd(1).Caption = "X"
    ElseIf C(2, 0) = "" And (C(0, 0) = "O" And C(1, 0) = "O" Or C(1, 1) = "O" And C(0, 2) = "O" Or C(2, 1) = "O" And C(2, 2) = "O") Then    '//win for Square 3
    Cmd(2).Caption = "X"
    ElseIf C(0, 1) = "" And (C(1, 1) = "O" And C(2, 1) = "O" Or C(0, 0) = "O" And C(0, 2) = "O") Then    '//win for Square 4
    Cmd(3).Caption = "X"
    ElseIf C(1, 1) = "" And (C(0, 0) = "O" And C(2, 2) = "O" Or C(1, 0) = "O" And C(1, 2) = "O" Or C(2, 0) = "O" And C(0, 2) = "O" Or C(0, 1) = "O" And C(2, 1) = "O") Then    '//win for Square 5
    Cmd(4).Caption = "X"
    ElseIf C(2, 1) = "" And (C(2, 0) = "O" And C(2, 2) = "O" Or C(0, 1) = "O" And C(1, 1) = "O") Then    '//win for Square 6
    Cmd(5).Caption = "X"
    ElseIf C(0, 2) = "" And (C(0, 0) = "O" And C(0, 1) = "O" Or C(2, 0) = "O" And C(1, 1) = "O" Or C(1, 2) = "O" And C(2, 2) = "O") Then    '//win for Square 7
    Cmd(6).Caption = "X"
    ElseIf C(1, 2) = "" And (C(1, 0) = "O" And C(1, 1) = "O" Or C(0, 2) = "O" And C(2, 2) = "O") Then    '//win for Square 8
    Cmd(7).Caption = "X"
    ElseIf C(2, 2) = "" And (C(0, 0) = "O" And C(1, 1) = "O" Or C(2, 0) = "O" And C(2, 1) = "O" Or C(0, 2) = "O" And C(1, 2) = "O") Then    '//win for Square 9
    Cmd(8).Caption = "X"
    Else
Level = Level + 2
    If Level = 10 Then
    Level = 8
    End If
        Select Case Level
        Case 2
            If C(1, 1) = "X" Then
                If (C(1, 0) = "O" Or C(0, 1) = "O") And C(0, 0) = "" Then
                Cmd(0).Caption = "X"
                WinGame = 3
                ElseIf (C(1, 2) = "O" Or C(2, 1) = "O") And C(2, 2) = "" Then
                Cmd(8).Caption = "X"
                WinGame = 4
                ElseIf (C(0, 0) = "O" Or C(0, 2) = "O") And C(0, 1) = "" Then
                Cmd(3).Caption = "X"
                WinGame = 5
                ElseIf (C(2, 0) = "O" Or C(2, 2) = "O") And C(2, 1) = "" Then
                Cmd(5).Caption = "X"
                WinGame = 5
                End If
            End If
        Case 4
            If WinGame = 3 Then
                If C(0, 1) = "" Then
                Cmd(3).Caption = "X"
                Else
                Cmd(1).Caption = "X"
                End If
            ElseIf WinGame = 4 Then
                If C(1, 2) = "" Then
                Cmd(7).Caption = "X"
                ElseIf C(2, 1) = "" Then
                Cmd(5).Caption = "X"
                End If
            ElseIf WinGame = 5 Then
                If ((C(0, 0) = "O" And C(2, 1) = "O") Or (C(0, 1) = "O" And C(2, 0) = "O")) And C(1, 0) = "" Then
                Cmd(1).Caption = "X"
                Level = 2
                ElseIf ((C(0, 2) = "O" And C(2, 1) = "O") Or (C(0, 1) = "O" And C(2, 2) = "O")) And C(1, 2) = "" Then
                Cmd(7).Caption = "X"
                Level = 2
                ElseIf C(0, 0) = "" Then
                Cmd(0).Caption = "X"
                Level = 2
                ElseIf C(2, 0) = "" Then
                Cmd(2).Caption = "X"
                Level = 2
                ElseIf C(0, 2) = "" Then
                Cmd(6).Caption = "X"
                Level = 2
                ElseIf C(2, 2) = "" Then
                Cmd(8).Caption = "X"
                Level = 2
                End If
            End If
        Case 6
        If WinGame = 3 Then
        End If
        Case 8
        End Select
    End If
End Function
