VERSION 5.00
Begin VB.Form Checking5 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton navi 
      Caption         =   "&Last"
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton navi 
      Caption         =   "N&ext"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton navi 
      Caption         =   "&Back"
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton navi 
      Caption         =   "&First"
      Height          =   495
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton NewEntry 
      Caption         =   "&New"
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton DepChk 
      Caption         =   "&Deposit"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton UpdateData 
      Caption         =   "&Update"
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   7
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   6
      Top             =   1290
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   930
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   570
      Width           =   555
   End
End
Attribute VB_Name = "Checking5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddDeposit As Boolean
Dim TotalChecks As Integer, ScrollChecks As Integer
Dim TotalDeposits As Integer, ScrollDeposits As Integer
Dim Checks(0 To 5000) As Banking_Checks
Dim Deposits(0 To 5000) As Banking_Deposits
Private Type Banking_Checks
    Number As Integer
    Date As String * 8
    Description As String * 40
    Amount As String * 11
End Type
Private Type Banking_Deposits
    Date As String * 8
    Description As String * 40
    Amount As String * 11
End Type

Private Sub DepChk_Click()
If AddDeposit = True Then
    AddDeposit = False
    DepChk.Caption = "&Deposit"
    Text1(0).Text = ""
    Text1(0).Enabled = True
    Text1(0).SetFocus
Else
    AddDeposit = True
    DepChk.Caption = "&Check"
    Text1(0).Text = "DEPOSIT"
    Text1(0).Enabled = False
    Text1(1).SetFocus
End If
End Sub


Private Sub navi_Click(Index As Integer)

Select Case Index
    Case 0
        ScrollChecks = 0
    Case 1
        If ScrollChecks = 0 Then
        ScrollChecks = TotalChecks
        Else
        ScrollChecks = ScrollChecks - 1
        End If
    Case 2
        If ScrollChecks = TotalChecks Then
        ScrollChecks = 0
        Else
        ScrollChecks = ScrollChecks + 1
        End If
    Case 3
        ScrollChecks = TotalChecks
End Select
Update_Text1
End Sub

Private Sub Update_Text1()
Text1(0).Text = Checks(ScrollChecks).Number
Text1(1).Text = Checks(ScrollChecks).Date
Text1(2).Text = Checks(ScrollChecks).Description
Text1(3).Text = Checks(ScrollChecks).Amount
End Sub



Private Sub NewEntry_Click()
TotalChecks = TotalChecks + 1
ScrollChecks = TotalChecks
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = ""
End Sub


Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Checks(ScrollChecks).Number = Val(Text1(0).Text)
Checks(ScrollChecks).Date = Text1(1).Text
Checks(ScrollChecks).Description = Text1(2).Text
Checks(ScrollChecks).Amount = Val(Text1(3).Text)

End Sub

Private Sub UpdateData_Click()
Dim i As Integer
Open App.Path & "\bank.dat" For Output As #1
    For i = 0 To TotalChecks
        Write #1, Checks(i).Number, Checks(i).Date, Checks(i).Description, Checks(i).Amount
    Next i
Close #1
End Sub
