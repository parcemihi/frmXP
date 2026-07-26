VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Special"
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dinner"
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lunch"
      Height          =   495
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Breakfast"
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sides"
      Height          =   495
      Index           =   3
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Appetizers"
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deserts"
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Drinks"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer, y As Integer, z As Integer

Private Sub Command1_Click(Index As Integer)
    Open App.Path & "\menu.dat" For Append As #1
    'Open "menu.dat" For Output As #1
    Write #1, Text1.Text
    Close #1



End Sub

Private Sub Form_Load()
Form1.WindowState = 2


For z = 0 To 7
    Command1(z).Height = Screen.Height / 9
    Command1(z).Width = Screen.Height / 9
    Command1(z).Left = 0
    Command1(z).Top = z * Screen.Height / 9
Next z

End Sub
