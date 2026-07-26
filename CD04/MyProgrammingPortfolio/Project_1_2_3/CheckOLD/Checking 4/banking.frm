VERSION 5.00
Begin VB.Form Banking 
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   9075
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton AddCheck 
      Caption         =   "New &Check"
      Height          =   495
      Left            =   6960
      TabIndex        =   24
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton DataUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   5640
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox DepTotal 
      Height          =   285
      Left            =   7440
      TabIndex        =   22
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox DepAmount 
      Height          =   315
      Left            =   7440
      TabIndex        =   21
      Text            =   "Amount"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox DepDate 
      Height          =   285
      Left            =   7440
      TabIndex        =   19
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox DepDescription 
      Height          =   315
      Left            =   7440
      TabIndex        =   18
      Text            =   "Description"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton NewDeposit 
      Caption         =   "&Deposit"
      Height          =   495
      Left            =   8280
      TabIndex        =   17
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton NaviScroll 
      Caption         =   "&Last"
      Height          =   495
      Index           =   3
      Left            =   4200
      TabIndex        =   13
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton NaviScroll 
      Caption         =   "&Next"
      Height          =   495
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton NaviScroll 
      Caption         =   "&Back"
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton NaviScroll 
      Caption         =   "&First"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   600
      TabIndex        =   9
      Top             =   2880
      Width           =   8895
   End
   Begin VB.TextBox CheckField 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox CheckField 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox CheckField 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox CheckField 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox CheckField 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deposit Total"
      Height          =   195
      Index           =   7
      Left            =   6240
      TabIndex        =   20
      Top             =   1560
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deposit Date"
      Height          =   195
      Index           =   4
      Left            =   6345
      TabIndex        =   16
      Top             =   450
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deposit Amount"
      Height          =   195
      Index           =   6
      Left            =   6240
      TabIndex        =   15
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deposit Description"
      Height          =   195
      Index           =   5
      Left            =   5955
      TabIndex        =   14
      Top             =   810
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   1545
      TabIndex        =   8
      Top             =   1530
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   1290
      TabIndex        =   7
      Top             =   1170
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Date"
      Height          =   195
      Index           =   1
      Left            =   1230
      TabIndex        =   6
      Top             =   810
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Number / Seq"
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   5
      Top             =   450
      Width           =   1515
   End
End
Attribute VB_Name = "Banking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddNewCheck As Boolean

Private Sub Form_Load()
    basBanking.Input_Data
    UpdateCheckFields
End Sub

Private Sub CheckField_LostFocus(Index As Integer)
Dim i As Integer
    Select Case Index
        Case 0
            Checking(CurrentCheck).Number = Banking.CheckField(Index).Text
        Case 1
            Checking(CurrentCheck).Offset = Val(CheckField(Index).Text)
        Case 2
            Checking(CurrentCheck).Date = CheckField(Index).Text
        Case 3
            Checking(CurrentCheck).Description = CheckField(Index).Text
        Case 4
            CheckField(Index).Text = Format(CheckField(Index).Text, "##,##0.00")
            For i = 1 To 10 - Len(CheckField(Index).Text)
                CheckField(Index).Text = " " & CheckField(Index).Text
            Next i
            CheckField(Index).Text = "$" & CheckField(Index).Text
            Checking(CurrentCheck).Amount = CheckField(Index).Text
    End Select

End Sub

Private Sub AddCheck_Click()
Dim i As Integer
TotalChecks = TotalChecks + 1
CurrentCheck = TotalChecks
For i = 0 To 4
    CheckField(i).Text = ""
Next i
CheckField(0).SetFocus
End Sub

Private Sub UpdateDataArray()
    Checking(CurrentCheck).Number = CheckField(0).Text
    Checking(CurrentCheck).Offset = Val(CheckField(1).Text)
    Checking(CurrentCheck).Date = CheckField(2).Text
    Checking(CurrentCheck).Description = CheckField(3).Text
    Checking(CurrentCheck).Amount = CheckField(4).Text
End Sub

Private Sub UpdateCheckFields()
Dim i As Integer
    CheckField(0).Text = Checking(CurrentCheck).Number
    CheckField(1).Text = Checking(CurrentCheck).Offset
    CheckField(2).Text = Checking(CurrentCheck).Date
    CheckField(3).Text = Checking(CurrentCheck).Description
    CheckField(4).Text = Checking(CurrentCheck).Amount
    
    CheckField(4).Text = Format(CheckField(4).Text, "##,##0.00")
    For i = 1 To 10 - Len(CheckField(4).Text)
        CheckField(4).Text = " " & CheckField(4).Text
    Next i
    CheckField(4).Text = "$" & CheckField(4).Text
End Sub

Private Sub NaviScroll_Click(Index As Integer)
Select Case Index
    Case 0
        CurrentCheck = 0
    Case 1
        If CurrentCheck = 0 Then
            CurrentCheck = TotalChecks
        Else
            CurrentCheck = CurrentCheck - 1
        End If
    Case 2
        If CurrentCheck = TotalChecks Then
            CurrentCheck = 0
        Else
            CurrentCheck = CurrentCheck + 1
        End If
    Case 3
        CurrentCheck = TotalChecks
End Select
UpdateCheckFields
End Sub

Private Sub CheckField_GotFocus(Index As Integer)
    CheckField(Index).SelStart = 0
    CheckField(Index).SelLength = Len(CheckField(Index))
End Sub

Private Sub NewDeposit_Click()
DepInput.Show
End Sub

Private Sub DataUpdate_Click()
basBanking.Output_Data
End Sub






'If AddNewCheck = False Then
'    AddCheck.Default = True
'    AddNewCheck = True
'    AddCheck.Caption = "&Done"
'    TotalChecks = TotalChecks + 1
'    CurrentCheck = TotalChecks
'    For i = 0 To 4
'    CheckField(i).Text = ""
'    Next i
'    For i = 0 To 3
'    NaviScroll(i).Enabled = False
'    Next i
'    DataUpdate.Enabled = False
'    NewDeposit.Enabled = False
'    DepAmount.Enabled = False
'    DepDate.Enabled = False
'    DepDescription.Enabled = False
'    DepTotal.Enabled = False
'Else
'    If CheckField(0).Text = "" Then
'        msg = MsgBox("Check Number can not be left blank." & Chr(13) & _
'        "Would you like to cancel this action?", vbYesNo)
'
'        If msg = 6 Then
'            TotalChecks = TotalChecks - 1
'            CurrentCheck = TotalChecks
'            AddNewCheck = False
'            AddCheck.Caption = "New &Check"
'            For i = 0 To 3
'            NaviScroll(i).Enabled = True
'            Next i
'            DataUpdate.Enabled = True
'            NewDeposit.Enabled = True
'            DepAmount.Enabled = True
'            DepDate.Enabled = True
'            DepDescription.Enabled = True
'            DepTotal.Enabled = True
'            UpdateCheckFields
'        End If
'    Else
'        AddNewCheck = False
'        AddCheck.Caption = "New &Check"
'        For i = 0 To 3
'        NaviScroll(i).Enabled = True
'        Next i
'        DataUpdate.Enabled = True
'        NewDeposit.Enabled = True
'        DepAmount.Enabled = True
'        DepDate.Enabled = True
'        DepDescription.Enabled = True
'        DepTotal.Enabled = True
'        UpdateDataArray
'        UpdateCheckFields
'    End If
'End If
'CheckField(0).SetFocus

'        For i = 0 To TotalChecks
'            If Checking(i).Number = Checking(CurrentCheck).Number Then
'                msg = MsgBox("Duplicate Check Number", vbYesNoCancel, "Duplicate Check Number")
'            End If
'        Next i



'Dim x As String
'Dim i As Integer
'Dim msg As Integer
'
'If AddNewCheck = False Then
'    Select Case Index
'        Case 0
'            Checking(CurrentCheck).Number = Banking.CheckField(Index).Text
'        Case 1
'            Checking(CurrentCheck).Offset = CheckField(Index).Text
'        Case 2
'            Checking(CurrentCheck).Date = CheckField(Index).Text
'        Case 3
'            Checking(CurrentCheck).Description = CheckField(Index).Text
'        Case 4
'            CheckField(Index).Text = Format(CheckField(Index).Text, "##,##0.00")
'            For i = 1 To 10 - Len(CheckField(Index).Text)
'                CheckField(Index).Text = " " & CheckField(Index).Text
'            Next i
'            CheckField(Index).Text = "$" & CheckField(Index).Text
'            Checking(CurrentCheck).Amount = CheckField(Index).Text
'    End Select
'    UpdateCheckFields
'
'End If

