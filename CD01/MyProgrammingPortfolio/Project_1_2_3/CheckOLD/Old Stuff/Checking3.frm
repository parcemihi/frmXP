VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Check Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Next"
      Height          =   495
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Back"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   6135
      Left            =   3120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sort By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   5160
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CHKIndex As Integer
Dim EndOfCHK As Integer

Dim CHK() As CheckArray
Private Type CheckArray
    Number As Integer
    ChkDate As String
    Amount As Currency
    PaidTo As String
End Type

Private Sub Command2_Click()
If Not CHKIndex = EndOfCHK Then
    CHKIndex = CHKIndex + 1
    Text1(0).Text = CHK(CHKIndex).Number
    Text1(1).Text = CHK(CHKIndex).ChkDate
    Text1(2).Text = CHK(CHKIndex).Amount
    Text1(3).Text = CHK(CHKIndex).PaidTo
End If
End Sub
Private Sub Command1_Click()
If Not CHKIndex = 0 Then
    CHKIndex = CHKIndex - 1
    Text1(0).Text = CHK(CHKIndex).Number
    Text1(1).Text = CHK(CHKIndex).ChkDate
    Text1(2).Text = CHK(CHKIndex).Amount
    Text1(3).Text = CHK(CHKIndex).PaidTo
End If
End Sub

Private Sub Command3_Click()
    EndOfCHK = EndOfCHK + 1
    ReDim Preserve CHK(0 To EndOfCHK)
End Sub

Private Sub sort()

End Sub



Private Sub Command4_Click(Index As Integer)

End Sub

Private Sub Form_Load()
ReDim CHK(0 To 5000)
Open App.Path & "\03580946.k" For Input As #1
    Do Until EOF(1)
        Input #1, CHK(EndOfCHK).Number, CHK(EndOfCHK).ChkDate, CHK(EndOfCHK).Amount, CHK(EndOfCHK).PaidTo
        List1.AddItem CHK(EndOfCHK).Number & "   " & CHK(EndOfCHK).ChkDate & "   " & CHK(EndOfCHK).Amount & "   " & CHK(EndOfCHK).PaidTo
        EndOfCHK = EndOfCHK + 1
    Loop
Close #1
ReDim Preserve CHK(0 To EndOfCHK)
CHKIndex = 0
Text1(0).Text = CHK(0).Number
Text1(1).Text = CHK(0).ChkDate
Text1(2).Text = CHK(0).Amount
Text1(3).Text = CHK(0).PaidTo



End Sub

Private Sub List2_Click()

End Sub
