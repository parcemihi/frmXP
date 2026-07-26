VERSION 5.00
Begin VB.Form TicTacToe 
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
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
      Begin VB.Menu Reset 
         Caption         =   "&New Game"
      End
      Begin VB.Menu ComputerPlay 
         Caption         =   "Computer Plays First"
         Checked         =   -1  'True
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "TicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim j As Integer
Dim k As Integer

Private Sub ComputerPlay_Click()
    If ComputerPlay.Checked = True Then
    ComputerPlay.Checked = False
    ElseIf ComputerPlay.Checked = False Then
    ComputerPlay.Checked = True
    End If
    Reset_Click
End Sub

Private Sub Reset_Click()
    For i = 0 To 8
    Cmd(i).Caption = ""
    Next i
    If ComputerPlay.Checked = True Then
    ComputersTurn
    End If
End Sub

Public Function ComputersTurn()
'ArtificialIntelligence
    Randomize
    j = 1
    For i = 0 To 8
        If Cmd(j).Caption = "X" Or Cmd(j).Caption = "O" Then
        j = j + 1
        End If
    Next j
'//Capture win for X
    If Cmd(0).Caption = "" And _
        Cmd(1).Caption = "X" And Cmd(2).Caption = "X" Or Cmd(3).Caption = "X" And Cmd(6).Caption = "X" Or Cmd(4).Caption = "X" And Cmd(8).Caption = "X" Then
        Cmd(0).Caption = "X"
    ElseIf Cmd(1).Caption = "" And Cmd(4).Caption = "X" And Cmd(7).Caption = "X" Or Cmd(0).Caption = "X" And Cmd(2).Caption = "X" Then
        Cmd(1).Caption = "X"
    ElseIf Cmd(2).Caption = "" And Cmd(0).Caption = "X" And Cmd(1).Caption = "X" Or Cmd(4).Caption = "X" And Cmd(6).Caption = "X" Or Cmd(5).Caption = "X" And Cmd(8).Caption = "X" Then
        Cmd(2).Caption = "X"
    ElseIf Cmd(3).Caption = "" And Cmd(4).Caption = "X" And Cmd(5).Caption = "X" Or Cmd(0).Caption = "X" And Cmd(6).Caption = "X" Then
        Cmd(3).Caption = "X"
    ElseIf Cmd(4).Caption = "" And Cmd(0).Caption = "X" And Cmd(8).Caption = "X" Or Cmd(1).Caption = "X" And Cmd(7).Caption = "X" Or Cmd(2).Caption = "X" And Cmd(6).Caption = "X" Or Cmd(3).Caption = "X" And Cmd(5).Caption = "X" Then
        Cmd(4).Caption = "X"
    ElseIf Cmd(5).Caption = "" And Cmd(2).Caption = "X" And Cmd(8).Caption = "X" Or Cmd(3).Caption = "X" And Cmd(4).Caption = "X" Then
        Cmd(5).Caption = "X"
    ElseIf Cmd(6).Caption = "" And Cmd(0).Caption = "X" And Cmd(3).Caption = "X" Or Cmd(2).Caption = "X" And Cmd(4).Caption = "X" Or Cmd(7).Caption = "X" And Cmd(8).Caption = "X" Then
        Cmd(6).Caption = "X"
    ElseIf Cmd(7).Caption = "" And Cmd(1).Caption = "X" And Cmd(4).Caption = "X" Or Cmd(6).Caption = "X" And Cmd(8).Caption = "X" Then
        Cmd(7).Caption = "X"
    ElseIf Cmd(8).Caption = "" And Cmd(0).Caption = "X" And Cmd(4).Caption = "X" Or Cmd(2).Caption = "X" And Cmd(5).Caption = "X" Or Cmd(6).Caption = "X" And Cmd(7).Caption = "X" Then
        Cmd(8).Caption = "X"
'//Block Win for O
    ElseIf Cmd(0).Caption = "" And Cmd(1).Caption = "O" And Cmd(2).Caption = "O" Or Cmd(3).Caption = "O" And Cmd(6).Caption = "O" Or Cmd(4).Caption = "O" And Cmd(8).Caption = "O" Then
        Cmd(0).Caption = "X"
    ElseIf Cmd(1).Caption = "" And Cmd(4).Caption = "O" And Cmd(7).Caption = "O" Or Cmd(0).Caption = "O" And Cmd(2).Caption = "O" Then
        Cmd(1).Caption = "X"
    ElseIf Cmd(2).Caption = "" And Cmd(0).Caption = "O" And Cmd(1).Caption = "O" Or Cmd(4).Caption = "O" And Cmd(6).Caption = "O" Or Cmd(5).Caption = "O" And Cmd(8).Caption = "O" Then
        Cmd(2).Caption = "X"
    ElseIf Cmd(3).Caption = "" And Cmd(4).Caption = "O" And Cmd(5).Caption = "O" Or Cmd(0).Caption = "O" And Cmd(6).Caption = "O" Then
        Cmd(3).Caption = "X"
    ElseIf Cmd(4).Caption = "" And Cmd(0).Caption = "O" And Cmd(8).Caption = "O" Or Cmd(1).Caption = "O" And Cmd(7).Caption = "O" Or Cmd(2).Caption = "O" And Cmd(6).Caption = "O" Or Cmd(3).Caption = "O" And Cmd(5).Caption = "O" Then
        Cmd(4).Caption = "X"
    ElseIf Cmd(5).Caption = "" And Cmd(2).Caption = "O" And Cmd(8).Caption = "O" Or Cmd(3).Caption = "O" And Cmd(4).Caption = "O" Then
        Cmd(5).Caption = "X"
    ElseIf Cmd(6).Caption = "" And Cmd(0).Caption = "O" And Cmd(3).Caption = "O" Or Cmd(2).Caption = "O" And Cmd(4).Caption = "O" Or Cmd(7).Caption = "O" And Cmd(8).Caption = "O" Then
        Cmd(6).Caption = "X"
    ElseIf Cmd(7).Caption = "" And Cmd(1).Caption = "O" And Cmd(4).Caption = "O" Or Cmd(6).Caption = "O" And Cmd(8).Caption = "O" Then
        Cmd(7).Caption = "X"
    ElseIf Cmd(8).Caption = "" And Cmd(0).Caption = "O" And Cmd(4).Caption = "O" Or Cmd(2).Caption = "O" And Cmd(5).Caption = "O" Or Cmd(6).Caption = "O" And Cmd(7).Caption = "O" Then
        Cmd(8).Caption = "X"
'//Defense for move
    Else
        If j = 1 Then 'Offense
            Cmd(4).Caption = "X"
        End If
        
        If j = 2 Then 'Defense
            Do
            k = Int((4 * Rnd) + 1)
            If Cmd(4).Caption = "" Then
            Cmd(0).Caption = "X": Exit Do
            End If
            If k = 1 And Cmd(0).Caption = "" Then
            Cmd(0).Caption = "X": Exit Do
            End If
            If k = 2 And Cmd(2).Caption = "" Then
            Cmd(2).Caption = "X": Exit Do
            End If
            If k = 3 And Cmd(6).Caption = "" Then
            Cmd(6).Caption = "X": Exit Do
            End If
            If k = 4 And Cmd(8).Caption = "" Then
            Cmd(8).Caption = "X": Exit Do
            End If
            Loop
        End If
        
        If j = 3 Then 'Offense
            'Pick a corner
            Do
            If Cmd(0).Caption = "" And Cmd(2).Caption = "" And Cmd(6).Caption = "" And Cmd(8).Caption = "" Then
                k = Int((4 * Rnd) + 1)
                If k = 1 Then
                Cmd(0).Caption = "X"
                Exit Do
                
                End If
                If k = 2 Then
                Cmd(2).Caption = "X"
                Exit Do
                
                End If
                If k = 3 Then
                Cmd(6).Caption = "X"
                Exit Do
                
                End If
                If k = 4 And Cmd(8).Caption = "" Then
                Cmd(8).Caption = "X"
                Exit Do
                End If
            Loop
            Else
                k = Int((2 * Rnd) + 1)
                If Cmd(0).Caption = "O" Then
                    If k = 1 Then
                    Cmd(1).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(3).Caption = "X"
                    End If
                End If
                If Cmd(2).Caption = "O" Then
                    If k = 1 Then
                    Cmd(1).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(5).Caption = "X"
                    End If
                End If
                If Cmd(6).Caption = "O" Then
                    If k = 1 Then
                    Cmd(3).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(7).Caption = "X"
                    End If
                End If
                If Cmd(8).Caption = "O" Then
                    If k = 1 Then
                    Cmd(5).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(7).Caption = "X"
                    End If
                End If
            End If
        End If
        If j = 4 Then 'Defense
                k = Int((2 * Rnd) + 1)
                
                
                
                If Cmd(0).Caption = "X" Then
                    If k = 1 And Cmd(1).Caption = "" Then
                    Cmd(1).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(3).Caption = "X"
                    End If
                End If
                If Cmd(2).Caption = "O" Then
                    If k = 1 Then
                    Cmd(1).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(5).Caption = "X"
                    End If
                End If
                If Cmd(6).Caption = "O" Then
                    If k = 1 Then
                    Cmd(3).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(7).Caption = "X"
                    End If
                End If
                If Cmd(8).Caption = "O" Then
                    If k = 1 Then
                    Cmd(5).Caption = "X"
                    ElseIf k = 2 Then
                    Cmd(7).Caption = "X"
                    End If
                End If
        
        End If
        If j = 5 Then 'Offense
        End If
        If j = 6 Then 'Defense
        End If
        If j = 7 Then 'Offense
        End If
        If j = 8 Then 'Defense
        End If
        If j = 9 Then 'Offense
        End If
    
    
    End If
End Function












Public Function Check()

k = 0
    For i = 0 To 2
        For j = 0 To 2
        C(j, i) = Cmd(k).Caption
        k = k + 1
        Next j
    Next i
    If _
    Cmd(0).Caption = "X" And Cmd(3).Caption = "X" And Cmd(6).Caption = "X" Or _
    Cmd(1).Caption = "X" And Cmd(4).Caption = "X" And Cmd(7).Caption = "X" Or _
    Cmd(2).Caption = "X" And Cmd(5).Caption = "X" And Cmd(8).Caption = "X" Or _
    Cmd(0).Caption = "X" And Cmd(1).Caption = "X" And Cmd(2).Caption = "X" Or _
    Cmd(3).Caption = "X" And Cmd(4).Caption = "X" And Cmd(5).Caption = "X" Or _
    Cmd(6).Caption = "X" And Cmd(7).Caption = "X" And Cmd(8).Caption = "X" Or _
    Cmd(0).Caption = "X" And Cmd(4).Caption = "X" And Cmd(8).Caption = "X" Or _
    Cmd(2).Caption = "X" And Cmd(4).Caption = "X" And Cmd(6).Caption = "X" Then
    MsgBox ("X Wins")
    ResetGame
    ElseIf _
    Cmd(0).Caption = "O" And Cmd(3).Caption = "O" And Cmd(6).Caption = "O" Or _
    Cmd(1).Caption = "O" And Cmd(4).Caption = "O" And Cmd(7).Caption = "O" Or _
    Cmd(2).Caption = "O" And Cmd(5).Caption = "O" And Cmd(8).Caption = "O" Or _
    Cmd(0).Caption = "O" And Cmd(1).Caption = "O" And Cmd(2).Caption = "O" Or _
    Cmd(3).Caption = "O" And Cmd(4).Caption = "O" And Cmd(5).Caption = "O" Or _
    Cmd(6).Caption = "O" And Cmd(7).Caption = "O" And Cmd(8).Caption = "O" Or _
    Cmd(0).Caption = "O" And Cmd(4).Caption = "O" And Cmd(8).Caption = "O" Or _
    Cmd(2).Caption = "O" And Cmd(4).Caption = "O" And Cmd(6).Caption = "O" Then
    MsgBox ("O Wins")
    ResetGame
    ElseIf _
    (Cmd(0).Caption = "X" Or Cmd(0).Caption = "O") And (Cmd(3).Caption = "X" Or Cmd(3).Caption = "O") And _
    (Cmd(6).Caption = "X" Or Cmd(6).Caption = "O") And (Cmd(1).Caption = "X" Or Cmd(1).Caption = "O") And _
    (Cmd(4).Caption = "X" Or Cmd(4).Caption = "O") And (Cmd(7).Caption = "X" Or Cmd(7).Caption = "O") And _
    (Cmd(2).Caption = "X" Or Cmd(2).Caption = "O") And (Cmd(5).Caption = "X" Or Cmd(5).Caption = "O") And _
    (Cmd(8).Caption = "X" Or Cmd(8).Caption = "O") And (Cmd(0).Caption = "X" Or Cmd(0).Caption = "O") And _
    (Cmd(1).Caption = "X" Or Cmd(1).Caption = "O") And (Cmd(2).Caption = "X" Or Cmd(2).Caption = "O") And _
    (Cmd(3).Caption = "X" Or Cmd(3).Caption = "O") And (Cmd(4).Caption = "X" Or Cmd(4).Caption = "O") And _
    (Cmd(5).Caption = "X" Or Cmd(5).Caption = "O") And (Cmd(6).Caption = "X" Or Cmd(6).Caption = "O") And _
    (Cmd(7).Caption = "X" Or Cmd(7).Caption = "O") And (Cmd(8).Caption = "X" Or Cmd(8).Caption = "O") Then
    MsgBox ("Cats Game")
    ResetGame
    End If
End Function



Private Sub Exit_Click()
    End
End Sub

Private Sub Form_Load()
    Reset_Click

End Sub

