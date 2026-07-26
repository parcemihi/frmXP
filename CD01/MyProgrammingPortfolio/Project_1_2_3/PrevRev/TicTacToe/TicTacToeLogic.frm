VERSION 5.00
Begin VB.Form TicTacToeLogic 
   BackColor       =   &H8000000A&
   Caption         =   "Tic Tac Toe Logic"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   7080
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   3000
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   3120
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   18
      Top             =   360
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   26
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   25
         Left            =   840
         TabIndex        =   26
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   24
         Left            =   1440
         TabIndex        =   25
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   23
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   840
         TabIndex        =   23
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   21
         Left            =   1440
         TabIndex        =   22
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   20
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   840
         TabIndex        =   20
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   1440
         TabIndex        =   19
         Top             =   1560
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Position Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   1440
         TabIndex        =   17
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   12
         Left            =   1440
         TabIndex        =   14
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   13
         Left            =   840
         TabIndex        =   13
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   14
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   15
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   16
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   17
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.ListBox List3 
      Height          =   1230
      ItemData        =   "TicTacToeLogic.frx":0000
      Left            =   5640
      List            =   "TicTacToeLogic.frx":0002
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1230
      ItemData        =   "TicTacToeLogic.frx":0004
      Left            =   4320
      List            =   "TicTacToeLogic.frx":0006
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Filter"
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "TicTacToeLogic.frx":0008
      Left            =   3000
      List            =   "TicTacToeLogic.frx":000A
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Combos = "
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Winning  Combos ="
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "TicTacToeLogic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer, b As Integer, c As Integer, d As Integer
Dim e As Integer, f As Integer, g As Integer, h As Integer
Dim cnt As Integer, entry As String, entri As String
Private Sub Command1_Click()
List1.Clear
For a = 1 To 9
  For b = 1 To 9
    For c = 1 To 9
    For d = 1 To 9
    For e = 1 To 9
    For f = 1 To 9
    For g = 1 To 9
    For h = 1 To 9
        If _
        a = b Or a = c Or a = d Or a = e Or a = f Or a = g Or a = h Or _
                 b = c Or b = d Or b = e Or b = f Or b = g Or b = h Or _
                          c = d Or c = e Or c = f Or c = g Or c = h Or _
                                   d = e Or d = f Or d = g Or d = h Or _
                                            e = f Or e = g Or e = h Or _
                                                     f = g Or f = h Or _
                                                              g = h Then
        Label1.Caption = a
        Else
        entri = entri & entry
        entry = h & g & f & e & d & c & b & a ' & "   " & cnt
        Text1.Text = entri & entry
        List1.AddItem entry
        Label1.Caption = "Combos = " & List1.ListCount
        End If
     Next h
     Next g
     Next f
     Next e
     Next d
     Next c
  Next b
Next a
cnt = 0
End Sub

Private Sub Command2_Click()
Dim i As Integer
'Const j = List1.List

For i = 0 To List1.ListCount
    If List1.List(i) = "123" Or _
       List1.List(i) = "456" Or _
       List1.List(i) = "789" Or _
       List1.List(i) = "147" Or _
       List1.List(i) = "258" Or _
       List1.List(i) = "369" Or _
       List1.List(i) = "159" Or _
       List1.List(i) = "357" _
    Then
    List2.AddItem List1.List(i)
    
    Else
    End If


Next i

Label3.Caption = "Winning  Combos = " & List2.ListCount
End Sub

Private Sub Timer1_Timer()
List1.Clear



a = a + 1
b = b + 1
c = c + 1
d = d + 1
e = e + 1
f = f + 1
g = g + 1
h = h + 1
        If _
        a = b Or a = c Or a = d Or a = e Or a = f Or a = g Or a = h Or _
                 b = c Or b = d Or b = e Or b = f Or b = g Or b = h Or _
                          c = d Or c = e Or c = f Or c = g Or c = h Or _
                                   d = e Or d = f Or d = g Or d = h Or _
                                            e = f Or e = g Or e = h Or _
                                                     f = g Or f = h Or _
                                                              g = h Then
Label1.Caption = a
Else
entri = entri & entry
entry = h & g & f & e & d & c & b & a ' & "   " & cnt
Text1.Text = entri & entry
List1.AddItem entry
Label1.Caption = "Combos = " & List1.ListCount
End If




End Sub
