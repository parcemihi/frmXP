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
   Begin VB.Menu Reset 
      Caption         =   "Reset"
   End
End
Attribute VB_Name = "TicTacToe2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Take1 As Integer, Take2 As Integer, Take3 As Integer, _
    Take4 As Integer, Take5 As Integer, Take6 As Integer, Take7 As Integer
Dim theArray(0 To 8) As String
Dim pass As Integer
Dim j As Integer, k As Integer

Dim Trncnt As Integer
Dim Level As Integer
Dim WinGame As Integer
Dim C(2, 2) As String
Dim R1 As Integer, R2 As Integer, R3 As Integer
Dim R4 As Integer, R5 As Integer, R6 As Integer
Dim R7 As Integer, R8 As Integer, R9 As Integer

Dim P(2, 2) As String
Dim i As Integer
Dim WhosTurn As Boolean
Private Sub Cmd_Click(Index As Integer)
If Cmd(Index).Caption = "" Then
    Cmd(Index).Caption = "O"

    Check
    remov
    ComputerGoes
    Check
    remov
    End If

End Sub
Private Sub Form_Load()
For i = 0 To 8
    theArray(i) = 0
Next i

k = 0
    For i = 0 To 8
    Cmd(i).Caption = ""
    Next i

Level = 0
    Cmd(4).Caption = "X"
    For i = 0 To 2
        For j = 0 To 2
        C(j, i) = Cmd(k).Caption
        P(j, i) = Cmd(k).Caption
        k = k + 1
        Next j
    Next i
pass = 0
Trncnt = 0

Take1 = 0
Take2 = 0
Take3 = 0
Take4 = 0
Take5 = 0
Take6 = 0
Take7 = 0

remov
'Check

End Sub


Public Function remov()
For i = 0 To 8
pass = theArray(i) + pass
Next i

If theArray(0) = 0 And (Cmd(0).Caption = "X" Or Cmd(0).Caption = "O") Then
theArray(0) = 1
End If
If theArray(1) = 0 And (Cmd(1).Caption = "X" Or Cmd(1).Caption = "O") Then
theArray(1) = 2
End If
If theArray(2) = 0 And (Cmd(2).Caption = "X" Or Cmd(2).Caption = "O") Then
theArray(2) = 4
End If
If theArray(3) = 0 And (Cmd(3).Caption = "X" Or Cmd(3).Caption = "O") Then
theArray(3) = 8
End If
If theArray(4) = 0 And (Cmd(4).Caption = "X" Or Cmd(4).Caption = "O") Then
theArray(4) = 16
End If
If theArray(5) = 0 And (Cmd(5).Caption = "X" Or Cmd(5).Caption = "O") Then
theArray(5) = 32
End If
If theArray(6) = 0 And (Cmd(6).Caption = "X" Or Cmd(6).Caption = "O") Then
theArray(6) = 64
End If
If theArray(7) = 0 And (Cmd(7).Caption = "X" Or Cmd(7).Caption = "O") Then
theArray(7) = 128
End If
If theArray(8) = 0 And (Cmd(8).Caption = "X" Or Cmd(8).Caption = "O") Then
theArray(8) = 256
End If

Trncnt = Trncnt + 1

Select Case Trncnt
    Case 1
        For i = 0 To 8
        Take1 = theArray(i) + Take1
        Next i
        Take1 = Take1 - pass
        For i = 0 To 8
        If theArray(i) = Take2 Then
        Cmd(i).Caption = ""
        End If
        Next i
    Case 2
        For i = 0 To 8
        Take2 = theArray(i) + Take2
        Next i
        Take2 = Take2 - pass
        For i = 0 To 8
        If theArray(i) = Take3 Then
        Cmd(i).Caption = ""
        End If
        Next i
    Case 3
        For i = 0 To 8
        Take3 = theArray(i) + Take3
        Next i
        Take3 = Take3 - pass
        For i = 0 To 8
        If theArray(i) = Take4 Then
        Cmd(i).Caption = ""
        End If
        Next i
    Case 4
        For i = 0 To 8
        Take4 = theArray(i) + Take4
        Next i
        Take4 = Take4 - pass
        For i = 0 To 8
        If theArray(i) = Take5 Then
        Cmd(i).Caption = ""
        End If
        Next i
    Case 5
        For i = 0 To 8
        Take5 = theArray(i) + Take5
        Next i
        Take5 = Take5 - pass
        For i = 0 To 8
        If theArray(i) = Take6 Then
        Cmd(i).Caption = ""
        End If
        Next i
    Case 6
        For i = 0 To 8
        Take6 = theArray(i) + Take6
        Next i
        Take6 = Take6 - pass
        For i = 0 To 8
        If theArray(i) = Take7 Then
        Cmd(i).Caption = ""
        End If
        Next i
    Case 7
        For i = 0 To 8
        Take7 = theArray(i) + Take7
        Next i
        Take7 = Take7 - pass
        For i = 0 To 8
        If theArray(i) = Take1 Then
        Cmd(i).Caption = ""
        End If
        Next i
        Trncnt = 0
End Select
End Function
Public Function Check()

k = 0
    For i = 0 To 2
        For j = 0 To 2
        C(j, i) = Cmd(k).Caption
        P(j, i) = Cmd(k).Caption
        k = k + 1
        Next j
    Next i
    If _
    P(0, 0) = "X" And P(0, 1) = "X" And P(0, 2) = "X" Or _
    P(1, 0) = "X" And P(1, 1) = "X" And P(1, 2) = "X" Or _
    P(2, 0) = "X" And P(2, 1) = "X" And P(2, 2) = "X" Or _
    P(0, 0) = "X" And P(1, 0) = "X" And P(2, 0) = "X" Or _
    P(0, 1) = "X" And P(1, 1) = "X" And P(2, 1) = "X" Or _
    P(0, 2) = "X" And P(1, 2) = "X" And P(2, 2) = "X" Or _
    P(0, 0) = "X" And P(1, 1) = "X" And P(2, 2) = "X" Or _
    P(2, 0) = "X" And P(1, 1) = "X" And P(0, 2) = "X" _
    Then
    MsgBox ("X Wins")
    Form_Load
    ElseIf _
    P(0, 0) = "O" And P(0, 1) = "O" And P(0, 2) = "O" Or _
    P(1, 0) = "O" And P(1, 1) = "O" And P(1, 2) = "O" Or _
    P(2, 0) = "O" And P(2, 1) = "O" And P(2, 2) = "O" Or _
    P(0, 0) = "O" And P(1, 0) = "O" And P(2, 0) = "O" Or _
    P(0, 1) = "O" And P(1, 1) = "O" And P(2, 1) = "O" Or _
    P(0, 2) = "O" And P(1, 2) = "O" And P(2, 2) = "O" Or _
    P(0, 0) = "O" And P(1, 1) = "O" And P(2, 2) = "O" Or _
    P(2, 0) = "O" And P(1, 1) = "O" And P(0, 2) = "O" _
    Then
    MsgBox ("O Wins")
    Form_Load
    ElseIf _
    (P(0, 0) = "X" Or P(0, 0) = "O") And _
    (P(0, 1) = "X" Or P(0, 1) = "O") And _
    (P(0, 2) = "X" Or P(0, 2) = "O") And _
    (P(1, 0) = "X" Or P(1, 0) = "O") And _
    (P(1, 1) = "X" Or P(1, 1) = "O") And _
    (P(1, 2) = "X" Or P(1, 2) = "O") And _
    (P(2, 0) = "X" Or P(2, 0) = "O") And _
    (P(2, 1) = "X" Or P(2, 1) = "O") And _
    (P(2, 2) = "X" Or P(2, 2) = "O") And _
    (P(0, 0) = "X" Or P(0, 0) = "O") And _
    (P(1, 0) = "X" Or P(1, 0) = "O") And _
    (P(2, 0) = "X" Or P(2, 0) = "O") And _
    (P(0, 1) = "X" Or P(0, 1) = "O") And _
    (P(1, 1) = "X" Or P(1, 1) = "O") And _
    (P(2, 1) = "X" Or P(2, 1) = "O") And _
    (P(0, 2) = "X" Or P(0, 2) = "O") And _
    (P(1, 2) = "X" Or P(1, 2) = "O") And _
    (P(2, 2) = "X" Or P(2, 2) = "O") _
    Then
    MsgBox ("Cats Game")
    Form_Load
    End If
End Function
Private Sub Reset_Click()
    Form_Load
End Sub
Public Function ComputerGoes()

C(0, 0) = Cmd(0).Caption: C(1, 0) = Cmd(1).Caption: C(2, 0) = Cmd(2).Caption
C(0, 1) = Cmd(3).Caption: C(1, 1) = Cmd(4).Caption: C(2, 1) = Cmd(5).Caption
C(0, 2) = Cmd(6).Caption: C(1, 2) = Cmd(7).Caption: C(2, 2) = Cmd(8).Caption
    
    If C(0, 0) = "" And (C(1, 0) = "X" And C(2, 0) = "X" Or C(0, 1) = "X" And C(0, 2) = "X" Or C(1, 1) = "X" And C(2, 2) = "X") Then   '//win for Square 1
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
