VERSION 5.00
Begin VB.Form frmChecking 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton NaviScroll 
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
      Index           =   3
      Left            =   4320
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton NaviScroll 
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
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton NaviScroll 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton NaviScroll 
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
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
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
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   13
      Top             =   3840
      Width           =   10935
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
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Out of Seq"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   19
      Top             =   600
      Width           =   765
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
      TabIndex        =   18
      Top             =   3600
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   765
      TabIndex        =   17
      Top             =   1650
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   510
      TabIndex        =   16
      Top             =   1290
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Date"
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   14
      Top             =   600
      Width           =   1065
   End
End
Attribute VB_Name = "frmChecking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EditCheck As Boolean
Dim NumOfChecks As Integer
Dim Scroll_CHK As Integer
Dim Checks_Written(0 To 2000) As CW
Private Type CW
    CheckNum As Integer
    OutOfSeq As Integer
    ChekDate As Date
    Descript As String
    Amount As Currency
End Type

Private Sub Form_Load()
NumOfChecks = -1
Open App.Path & "\checks.k" For Input As #1
    Do Until EOF(1)
        NumOfChecks = NumOfChecks + 1
        Input #1, Checks_Written(NumOfChecks).CheckNum, _
                  Checks_Written(NumOfChecks).OutOfSeq, _
                  Checks_Written(NumOfChecks).ChekDate, _
                  Checks_Written(NumOfChecks).Descript, _
                  Checks_Written(NumOfChecks).Amount
        Update_List
    Loop
Close #1
UpdateText
End Sub
Private Sub Update_List()
Dim ChkField(0 To 4) As String
Dim i As Integer
Dim LengthOK As Boolean
ChkField(0) = Checks_Written(NumOfChecks).CheckNum
ChkField(1) = Checks_Written(NumOfChecks).OutOfSeq
ChkField(2) = Checks_Written(NumOfChecks).ChekDate
ChkField(3) = Checks_Written(NumOfChecks).Descript
ChkField(4) = Checks_Written(NumOfChecks).Amount

LengthOK = False
Do While Not LengthOK
LengthOK = True
    If Len(ChkField(0)) < 5 Then
        ChkField(0) = " " & ChkField(0):        LengthOK = False
    End If
    If Len(ChkField(1)) < 1 Then
        ChkField(1) = " " & ChkField(1):        LengthOK = False
    End If
    If Len(ChkField(2)) < 8 Then
        ChkField(2) = " " & ChkField(2):        LengthOK = False
    End If
    If Len(ChkField(3)) < 20 Then
        ChkField(3) = " " & ChkField(3):        LengthOK = False
    End If
    If Len(ChkField(4)) < 9 Then
        ChkField(4) = " " & ChkField(4):        LengthOK = False
    End If
Loop



List1.AddItem ChkField(0) & "." & ChkField(1) & " " & ChkField(2) & " " & ChkField(3) & " " & ChkField(4)


End Sub
Private Sub cnmdNew_Click()
Dim i As Integer
For i = 0 To 4
    Text1(i).Text = ""
Next i
Text1(0).SetFocus
NumOfChecks = NumOfChecks + 1
End Sub

Private Sub cmdUpdate_Click()
Dim msg As Integer
Dim IsCheckOK As Boolean
Dim CheckForDuplicate As Integer
IsCheckOK = False

Do Until IsCheckOK = True
    If Checks_Written(CheckForDuplicate).CheckNum = Val(Text1(0).Text) Then
        If Checks_Written(CheckForDuplicate).OutOfSeq = Val(Text1(1).Text) Then
            IsCheckOK = False
            msg = MsgBox("Error, Duplicate check. " & Chr(13) & "If this is a bank check, increment the 'Out Of Seq' text box", vbOKOnly, "Error, Duplicate Check")
        End If
    End If
    CheckForDuplicate = CheckForDuplicate + 1
    
    If CheckForDuplicate = NumOfChecks Then
        IsCheckOK = True
        NumOfChecks = NumOfChecks + 1
    End If
    
Loop
End Sub

Private Sub Update_Checking()
    Checks_Written(NumOfChecks).CheckNum = Text1(0).Text
    Checks_Written(NumOfChecks).OutOfSeq = Text1(1).Text
    Checks_Written(NumOfChecks).ChekDate = Text1(2).Text
    Checks_Written(NumOfChecks).Descript = Text1(3).Text
    Checks_Written(NumOfChecks).Amount = Text1(4).Text
End Sub

Private Sub UpdateText()
    Text1(0).Text = Checks_Written(Scroll_CHK).CheckNum
    Text1(1).Text = Checks_Written(Scroll_CHK).OutOfSeq
    Text1(2).Text = Checks_Written(Scroll_CHK).ChekDate
    Text1(3).Text = Checks_Written(Scroll_CHK).Descript
    Text1(4).Text = Format(Checks_Written(Scroll_CHK).Amount, "Currency")
End Sub

Private Sub NaviScroll_Click(Index As Integer)
Select Case NaviScroll(Index).Caption
    Case "F&irst"
        Scroll_CHK = 0
    Case "&Backward"
        If Scroll_CHK = 0 Then
            Scroll_CHK = NumOfChecks
        Else
            Scroll_CHK = Scroll_CHK - 1
        End If
    Case "&Forward"
        If Scroll_CHK = NumOfChecks Then
            Scroll_CHK = 0
        Else
            Scroll_CHK = Scroll_CHK + 1
        End If
    Case "L&ast"
        Scroll_CHK = NumOfChecks
End Select
UpdateText
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub
