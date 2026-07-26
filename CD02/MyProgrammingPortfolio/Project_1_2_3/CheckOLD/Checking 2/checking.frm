VERSION 5.00
Begin VB.Form frmChecking 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cnmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdWithdrawl 
      Caption         =   "&Withdrawls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeposit 
      Caption         =   "&Deposits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "L&ast"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "F&irst"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4155
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   6495
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackward 
      Caption         =   "&Backward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "&Forward"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   3600
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   885
      TabIndex        =   11
      Top             =   1770
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   10
      Top             =   1410
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Date"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   9
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   8
      Top             =   690
      Width           =   1065
   End
End
Attribute VB_Name = "frmChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim isCheck As Boolean

Private Sub cnmdNew_Click()
Select Case isCheck
    Case True
    frmWithdrawls.Show
    Case False
    frmDeposits.Show
End Select
End Sub

Private Sub Form_Load()
isCheck = True
totChks = 0
totDeps = 0
CheckIndx = 0
DepositIndx = 0
    Open App.Path & "\checks.k" For Input As #1
        Do Until EOF(1)
            Input #1, Checks(totChks).c1Numb, Checks(totChks).c2Spec, Checks(totChks).c3Date, Checks(totChks).c4Desr, Checks(totChks).c5Amnt
            totChks = totChks + 1
        Loop
    Close #1
    Open App.Path & "\deposit.k" For Input As #1
        Do Until EOF(1)
            Input #1, Deposits(totDeps).d1Indx, Deposits(totDeps).d2Date, Deposits(totDeps).d3Desr, Deposits(totDeps).d4Amnt
            totDeps = totDeps + 1
        Loop
    Close #1
Deposit_Withdrawl
End Sub

Private Sub Deposit_Withdrawl()
Dim retVal
    Select Case isCheck
        Case True
            retVal = navigate_Entry _
            (Checks(CheckIndx).c1Numb, _
            Checks(CheckIndx).c2Spec, _
            Checks(CheckIndx).c3Date, _
            Checks(CheckIndx).c4Desr, _
            Checks(CheckIndx).c5Amnt)
        Case False
'            retVal = navigate_Entry _
'            (Deposits(DepositIndx).d1Indx, , _
'            Deposits(DepositIndx).d2Date, _
'            Deposits(DepositIndx).d3Desr, _
'            Deposits(DepositIndx).d4Amnt)
    End Select

End Sub

Function navigate_Entry(Index As Integer, Optional Special As Integer, Optional ChkDate As Date, Optional Description As String, Optional Amount As Currency)
'    Text1(0).Text = Index
'    Text1(1).Text = Special
'    Text1(2).Text = F3
'    Text1(3).Text = F4
End Function

Private Sub cmdFirst_Click()
Select Case isCheck
    Case True
        CheckIndx = 0
    Case False
        DepositIndx = 0
End Select
Deposit_Withdrawl
End Sub


Private Sub cmdForward_Click()
Select Case isCheck
    Case True
        If CheckIndx = totChks - 1 Then
            CheckIndx = 0
        Else
            CheckIndx = CheckIndx + 1
        End If
    Case False
        If DepositIndx = totDeps - 1 Then
            DepositIndx = 0
        Else
            DepositIndx = DepositIndx + 1
        End If
End Select
Deposit_Withdrawl
End Sub

Private Sub cmdBackward_Click()
Select Case isCheck
    Case True
        If CheckIndx = 0 Then
            CheckIndx = totChks - 1
        Else
            CheckIndx = CheckIndx - 1
        End If
    Case False
        If DepositIndx = 0 Then
            DepositIndx = totDeps - 1
        Else
            DepositIndx = DepositIndx - 1
        End If
End Select
Deposit_Withdrawl
End Sub
Private Sub cmdLast_Click()
Select Case isCheck
    Case True
        CheckIndx = totChks - 1
    Case False
        DepositIndx = totDeps - 1
End Select
Deposit_Withdrawl
End Sub
Private Sub cmdDeposit_Click()
    If isCheck = True Then
        isCheck = False
    End If
    Deposit_Withdrawl
End Sub
Private Sub cmdWithdrawl_Click()
    If isCheck = False Then
        isCheck = True
    End If
    Deposit_Withdrawl
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub
Private Sub Text1_LostFocus(Index As Integer)
Select Case isCheck
    Case True
        Checks(CheckIndx).c1Numb = Text1(0).Text
        Checks(CheckIndx).c2Date = Text1(1).Text
        Checks(CheckIndx).c3Desr = Text1(2).Text
        Checks(CheckIndx).c4Amnt = Text1(3).Text
    Case False
        Deposits(DepositIndx).d1Indx = Text1(0).Text
        Deposits(DepositIndx).d2Date = Text1(1).Text
        Deposits(DepositIndx).d3Desr = Text1(2).Text
        Deposits(DepositIndx).d4Amnt = Text1(3).Text
End Select
End Sub
Private Sub Fill_ListBox()
    For i = 0 To totChks
        List1.AddItem Checks(i).c1Numb & Checks(i).c2Date & Checks(i).c3Desr & Checks(i).c4Amnt & Checks(i).c5Depo
    Next i
End Sub
Private Sub cmdExit_Click()
    End
End Sub
'        Label1(0).Visible = True
'        Text1(0).Visible = True
'        Label1(0).Visible = False
'        Text1(0).Visible = False

