VERSION 5.00
Begin VB.Form account 
   Caption         =   "Form1"
   ClientHeight    =   6600
   ClientLeft      =   -1920
   ClientTop       =   645
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.CommandButton listView 
      Caption         =   "Taxable"
      Height          =   255
      Index           =   5
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton listView 
      Caption         =   "Cleared"
      Height          =   255
      Index           =   4
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton NewEntry 
      Caption         =   "New"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1335
   End
   Begin VB.OptionButton ChkDep_OPbttn 
      Caption         =   "&Deposit"
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton DeleteEntry 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton CommitData 
      Caption         =   "Com&mit"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton UpdateData 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtCheckEdit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
   Begin VB.OptionButton ChkDep_OPbttn 
      Caption         =   "&VISA"
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "&Last"
      Height          =   495
      Index           =   3
      Left            =   2400
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "&Next"
      Height          =   495
      Index           =   2
      Left            =   1635
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5400
      Width           =   735
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "&Back"
      Height          =   495
      Index           =   1
      Left            =   885
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5400
      Width           =   735
   End
   Begin VB.OptionButton ChkDep_OPbttn 
      Caption         =   "&Check"
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5940
      Left            =   3225
      TabIndex        =   8
      Top             =   240
      Width           =   6135
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   315
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CheckBox checkboxCleared 
      Alignment       =   1  'Right Justify
      Caption         =   "Cl&eared"
      Height          =   285
      Left            =   240
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtCheckEdit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox txtCheckEdit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtCheckEdit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtCheckEdit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtCheckEdit 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "&First"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton listView 
      Caption         =   "Date"
      Height          =   255
      Index           =   3
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton listView 
      Caption         =   "Deposits"
      Height          =   255
      Index           =   2
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton listView 
      Caption         =   "VISA"
      Height          =   255
      Index           =   1
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton listView 
      Caption         =   "Checks"
      Height          =   255
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "List By"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   40
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Available"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   39
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   38
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PT"
      Height          =   195
      Index           =   3
      Left            =   2160
      TabIndex        =   37
      Top             =   3045
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Amount"
      Height          =   195
      Index           =   2
      Left            =   1320
      TabIndex        =   36
      Top             =   2400
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transaction Amount"
      Height          =   195
      Index           =   5
      Left            =   1590
      TabIndex        =   35
      Top             =   2685
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   34
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   3405
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check #"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   3045
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Credits"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   31
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Debits"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   30
      Top             =   690
      Width           =   450
   End
   Begin VB.Label totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   29
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   28
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Balance"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   585
   End
   Begin VB.Label totals 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   26
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Update_Array
Update_Totals
Update_AccountEditor
Update_ListBox
End Sub

Private Sub Update_Array()
EntryPointer = 0
Open App.Path & "\44219512.dat" For Input As #1
    Do While Not EOF(1)
                  EntryPointer = EntryPointer + 1
        Input #1, checking.account(EntryPointer).ENT0_CVD_TYPE, _
                  checking.account(EntryPointer).ENT1_CLEARED, _
                  checking.account(EntryPointer).ENT2_NUM, _
                  checking.account(EntryPointer).ENT3_FAMILY, _
                  checking.account(EntryPointer).ENT4_MONTH, _
                  checking.account(EntryPointer).ENT5_DAY, _
                  checking.account(EntryPointer).ENT6_YEAR, _
                  checking.account(EntryPointer).ENT7_DESCRIPTION, _
                  checking.account(EntryPointer).ENT8_AMOUNT
    Loop
Close #1
TotalEntries = EntryPointer
'EntryPointer = 0
End Sub

Private Sub Update_Totals()
Dim AccountTotal As Single
Dim AccountDebits As Single
Dim AccountCredits  As Single
Dim i As Integer
AccountTotal = 0
AccountDebits = 0
AccountCredits = 0


'Current Balance Switch And Option ADD HERE
For i = 1 To TotalEntries
    AccountTotal = AccountTotal + checking.account(i).ENT8_AMOUNT
    If checking.account(i).ENT8_AMOUNT < 0 Then
        AccountDebits = AccountDebits + checking.account(i).ENT8_AMOUNT
        Else
        AccountCredits = AccountCredits + checking.account(i).ENT8_AMOUNT
        End If
Next i
totals(0).Caption = Format(AccountTotal, "Currency")
totals(1).Caption = Format(AccountDebits, "Currency")
totals(2).Caption = Format(AccountCredits, "Currency")
End Sub

Private Sub Update_AccountEditor()
If TotalEntries = 0 Then
    List1.Clear
Else
    Select Case checking.account(EntryPointer).ENT0_CVD_TYPE
        Case 0 '= Check
        ChkDep_OPbttn(0).Value = True
            txtCheckEdit(0).Text = checking.account(EntryPointer).ENT2_NUM
            txtCheckEdit(0).Enabled = True
        Case 1 '= VISA
        ChkDep_OPbttn(1).Value = True
            txtCheckEdit(0).Text = "VISA"
            txtCheckEdit(0).Enabled = False
        Case 2 '= Deposit
        ChkDep_OPbttn(2).Value = True
            txtCheckEdit(0).Text = "DEPOSIT"
            txtCheckEdit(0).Enabled = False
        End Select
    checkboxCleared.Value = checking.account(EntryPointer).ENT1_CLEARED
    txtCheckEdit(1).Text = checking.account(EntryPointer).ENT3_FAMILY
txtCheckEdit(2).Text = checking.account(EntryPointer).ENT4_MONTH
txtCheckEdit(3).Text = checking.account(EntryPointer).ENT5_DAY
txtCheckEdit(4).Text = checking.account(EntryPointer).ENT6_YEAR
txtCheckEdit(5).Text = checking.account(EntryPointer).ENT7_DESCRIPTION
txtCheckEdit(6).Text = Format(checking.account(EntryPointer).ENT8_AMOUNT, "Currency")
End If
End Sub

Private Sub Update_ListBox()
Dim i As Integer, x As Single
Dim chkDATE As String
Dim cNUMBER As String, cMONTH As String, cDAY As String, cYEAR As String, cDESCRIPTION As String, cSPACER As Integer, cAMOUNT As String
List1.Clear
For i = 1 To TotalEntries
    Select Case checking.account(i).ENT0_CVD_TYPE
        Case 0
        cNUMBER = " " & checking.account(i).ENT2_NUM
        Do While Len(cNUMBER) < 6
        cNUMBER = " " & cNUMBER
        Loop
        cNUMBER = cNUMBER & "-" & checking.account(i).ENT3_FAMILY
        Case 1
        cNUMBER = "VISA    "
        Case 2
        cNUMBER = checking.account(i).ENT3_FAMILY
            Do While Len(cNUMBER) < 4
                cNUMBER = "0" & cNUMBER
            Loop
            cNUMBER = "DEP " & cNUMBER
    End Select
    If checking.account(i).ENT4_MONTH < 10 Then
        cMONTH = "0" & checking.account(i).ENT4_MONTH
        Else
        cMONTH = checking.account(i).ENT4_MONTH
        End If
    If checking.account(i).ENT5_DAY < 10 Then
        cDAY = "0" & checking.account(i).ENT5_DAY
        Else
        cDAY = checking.account(i).ENT5_DAY
        End If
    cYEAR = Mid(checking.account(i).ENT6_YEAR, 3, 2)
    chkDATE = cMONTH & "/" & cDAY ' & "/" & cYEAR
    cDESCRIPTION = checking.account(i).ENT7_DESCRIPTION
    cAMOUNT = checking.account(i).ENT8_AMOUNT * 100
    x = Fix(checking.account(i).ENT8_AMOUNT) * 100
    x = cAMOUNT - x
    cAMOUNT = Abs(x)
    If Len(cAMOUNT) = 1 Then
        cAMOUNT = ".0" & cAMOUNT
        Else
        cAMOUNT = "." & cAMOUNT
    End If
    If checking.account(i).ENT0_CVD_TYPE < 2 Then
        cAMOUNT = Fix(checking.account(i).ENT8_AMOUNT) & cAMOUNT
        If Mid(cAMOUNT, 1, 1) = "-" Then
        cAMOUNT = "(" & Mid(cAMOUNT, 2, Len(cAMOUNT)) & ")"
        Else
        cAMOUNT = "(" & cAMOUNT & ")"
        End If
        Do While Len(cAMOUNT) < 9
            cAMOUNT = " " & cAMOUNT
        Loop
        cAMOUNT = "          $" & cAMOUNT
    Else
        cAMOUNT = Fix(checking.account(i).ENT8_AMOUNT) & cAMOUNT
        cAMOUNT = " " & cAMOUNT & " "
        Do While Len(cAMOUNT) < 9
            cAMOUNT = " " & cAMOUNT
        Loop
        cAMOUNT = "$" & cAMOUNT
    End If
    List1.AddItem cNUMBER & " " & chkDATE & " " & cDESCRIPTION & " " & cAMOUNT
Next i
End Sub

Private Sub NaviCheck_Click(Index As Integer)
'Possible Option Switch
UpdateData_Click

Select Case Index
    Case 0
        EntryPointer = 1
    Case 1
        If EntryPointer = 1 Then
            EntryPointer = TotalEntries
        Else
            EntryPointer = EntryPointer - 1
        End If
    Case 2
        If EntryPointer = TotalEntries Then
            EntryPointer = 1
        Else
            EntryPointer = EntryPointer + 1
        End If
    Case 3
        EntryPointer = TotalEntries
End Select
Update_AccountEditor
End Sub

Public Sub UpdateData_Click()
Dim i As Integer
Dim ff As Integer
Dim fixFormat As String

For i = 0 To 2
    If ChkDep_OPbttn(i).Value = True Then
        checking.account(EntryPointer).ENT0_CVD_TYPE = i
    End If
Next i
checking.account(EntryPointer).ENT1_CLEARED = checkboxCleared.Value
fixFormat = txtCheckEdit(6).Text
txtCheckEdit(6).Text = ""
For ff = 1 To Len(fixFormat)
    Select Case Mid(fixFormat, ff, 1)
        Case 0 To 9
        txtCheckEdit(6).Text = txtCheckEdit(6).Text & Mid(fixFormat, ff, 1)
        Case "."
        txtCheckEdit(6).Text = txtCheckEdit(6).Text & Mid(fixFormat, ff, 1)
    End Select
Next ff
Select Case checking.account(EntryPointer).ENT0_CVD_TYPE
    Case 0
    checking.account(EntryPointer).ENT8_AMOUNT = -1 * Abs(Val(txtCheckEdit(6).Text))
    checking.account(EntryPointer).ENT2_NUM = Val(txtCheckEdit(0).Text)
    Case 1
    checking.account(EntryPointer).ENT8_AMOUNT = -1 * Abs(Val(txtCheckEdit(6).Text))
    Case 2
    checking.account(EntryPointer).ENT8_AMOUNT = Abs(Val(txtCheckEdit(6).Text))
End Select
checking.account(EntryPointer).ENT3_FAMILY = Val(txtCheckEdit(1).Text)
checking.account(EntryPointer).ENT4_MONTH = Val(txtCheckEdit(2).Text)
checking.account(EntryPointer).ENT5_DAY = Val(txtCheckEdit(3).Text)
checking.account(EntryPointer).ENT6_YEAR = Val(txtCheckEdit(4).Text)
checking.account(EntryPointer).ENT7_DESCRIPTION = txtCheckEdit(5).Text
Update_Totals
Update_ListBox
Update_AccountEditor
End Sub

Private Sub CommitData_Click()
Dim i As Integer
Sort_Data 0, 1
UpdateData_Click
Open App.Path & "\44219512.dat" For Output As #1
    For i = 1 To TotalEntries
        Write #1, _
        checking.account(i).ENT0_CVD_TYPE, _
        checking.account(i).ENT1_CLEARED, _
        checking.account(i).ENT2_NUM, _
        checking.account(i).ENT3_FAMILY, _
        checking.account(i).ENT4_MONTH, _
        checking.account(i).ENT5_DAY, _
        checking.account(i).ENT6_YEAR, _
        checking.account(i).ENT7_DESCRIPTION, _
        checking.account(i).ENT8_AMOUNT
    Next i
Close #1
End Sub

Private Sub mnuDelete_Click()
DeleteEntry_Click
End Sub

Private Sub DeleteEntry_Click()
Dim msg As Integer
Dim i As Integer
msg = MsgBox("Are you certain you want to delete this entry?", vbYesNo)
If msg = 6 Then
    If TotalEntries <= 1 Then
        With checking.account(1)
            .ENT0_CVD_TYPE = 0: .ENT1_CLEARED = 0: .ENT2_NUM = 0: .ENT3_FAMILY = 0: .ENT4_MONTH = 0: .ENT5_DAY = 0: .ENT6_YEAR = 0: .ENT7_DESCRIPTION = "": .ENT8_AMOUNT = 0
        End With
        TotalEntries = 0
        For i = 0 To 6
            txtCheckEdit(i).Text = ""
        Next i
        ChkDep_OPbttn(0).Value = False: ChkDep_OPbttn(1).Value = False: ChkDep_OPbttn(2).Value = False: checkboxCleared.Value = 0
    Else
        If EntryPointer = TotalEntries Then
            TotalEntries = TotalEntries - 1
            EntryPointer = TotalEntries
        Else
            For i = EntryPointer To TotalEntries
                With checking.account(i)
                    .ENT0_CVD_TYPE = checking.account(i + 1).ENT0_CVD_TYPE
                    .ENT1_CLEARED = checking.account(i + 1).ENT1_CLEARED
                    .ENT2_NUM = checking.account(i + 1).ENT2_NUM
                    .ENT3_FAMILY = checking.account(i + 1).ENT3_FAMILY
                    .ENT4_MONTH = checking.account(i + 1).ENT4_MONTH
                    .ENT5_DAY = checking.account(i + 1).ENT5_DAY
                    .ENT6_YEAR = checking.account(i + 1).ENT6_YEAR
                    .ENT7_DESCRIPTION = checking.account(i + 1).ENT7_DESCRIPTION
                    .ENT8_AMOUNT = checking.account(i + 1).ENT8_AMOUNT
                End With
            Next i
            TotalEntries = TotalEntries - 1
        End If
    End If
    Update_AccountEditor
    Update_ListBox
    Update_Totals
End If
End Sub

Private Sub Sort_Data(SortPriority As Integer, UP_DOWN As Integer)
Dim IsArraySorted As Boolean, AP As Integer, aENT(0 To 8), bENT(0 To 8)
IsArraySorted = False

Do Until IsArraySorted = True
IsArraySorted = True
    For AP = 1 To checking.TotalEntries
        If UP_DOWN = 0 Then
            With checking.account(AP + 1)
                aENT(0) = .ENT0_CVD_TYPE: aENT(1) = .ENT1_CLEARED: aENT(2) = .ENT2_NUM: aENT(3) = .ENT3_FAMILY: aENT(4) = .ENT4_MONTH: aENT(5) = .ENT5_DAY: aENT(6) = .ENT6_YEAR: aENT(7) = .ENT7_DESCRIPTION: aENT(8) = .ENT8_AMOUNT
            End With
            With checking.account(AP)
                bENT(0) = .ENT0_CVD_TYPE: bENT(1) = .ENT1_CLEARED: bENT(2) = .ENT2_NUM: bENT(3) = .ENT3_FAMILY: bENT(4) = .ENT4_MONTH: bENT(5) = .ENT5_DAY: bENT(6) = .ENT6_YEAR: bENT(7) = .ENT7_DESCRIPTION: bENT(8) = .ENT8_AMOUNT
            End With
        ElseIf UP_DOWN = 1 Then
            With checking.account(AP)
                aENT(0) = .ENT0_CVD_TYPE: aENT(1) = .ENT1_CLEARED: aENT(2) = .ENT2_NUM: aENT(3) = .ENT3_FAMILY: aENT(4) = .ENT4_MONTH: aENT(5) = .ENT5_DAY: aENT(6) = .ENT6_YEAR: aENT(7) = .ENT7_DESCRIPTION: aENT(8) = .ENT8_AMOUNT
            End With
            With checking.account(AP + 1)
                bENT(0) = .ENT0_CVD_TYPE: bENT(1) = .ENT1_CLEARED: bENT(2) = .ENT2_NUM: bENT(3) = .ENT3_FAMILY: bENT(4) = .ENT4_MONTH: bENT(5) = .ENT5_DAY: bENT(6) = .ENT6_YEAR: bENT(7) = .ENT7_DESCRIPTION: bENT(8) = .ENT8_AMOUNT
            End With
        End If
        
            Select Case SortPriority
            Case 0 'ENT0
            Case 1 'ENT1
            Case 2 'ENT2
            Case 3 'ENT3
            Case 4 'ENT4
            Case 5 'ENT5
            Case 6 'ENT6
            Case 7 'ENT7
            Case 8 'ENT8
            End Select

'        For AP = 1 To checking.TotalEntries
'        With checking.account(AP)
'            If .ENT6_YEAR > checking.account(AP + 1).ENT6_YEAR Then
'            SwapArray AP, AP + 1
'            IsArraySorted = False
'        ElseIf .ENT6_YEAR = checking.account(AP + 1).ENT6_YEAR Then
'            If .ENT4_MONTH > checking.account(AP + 1).ENT4_MONTH Then
'            SwapArray AP, AP + 1
'            IsArraySorted = False
'            ElseIf .ENT4_MONTH = checking.account(AP + 1).ENT4_MONTH Then
'                If .ENT5_DAY > checking.account(AP + 1).ENT5_DAY Then
'                SwapArray AP, AP + 1
'                IsArraySorted = False
'                End If
'            End If
'        End If
'    End With
    Next AP
    Loop
'
End Sub

Function SwapArray(Epoint As Integer, Rpoint As Integer)
With checking.account(Epoint)
    checking.account(0).ENT0_CVD_TYPE = .ENT0_CVD_TYPE
    checking.account(0).ENT1_CLEARED = .ENT1_CLEARED
    checking.account(0).ENT2_NUM = .ENT2_NUM
    checking.account(0).ENT3_FAMILY = .ENT3_FAMILY
    checking.account(0).ENT4_MONTH = .ENT4_MONTH
    checking.account(0).ENT5_DAY = .ENT5_DAY
    checking.account(0).ENT6_YEAR = .ENT6_YEAR
    checking.account(0).ENT7_DESCRIPTION = .ENT7_DESCRIPTION
    checking.account(0).ENT8_AMOUNT = .ENT8_AMOUNT
End With
With checking.account(Epoint)
    .ENT0_CVD_TYPE = checking.account(Rpoint).ENT0_CVD_TYPE
    .ENT1_CLEARED = checking.account(Rpoint).ENT1_CLEARED
    .ENT2_NUM = checking.account(Rpoint).ENT2_NUM
    .ENT3_FAMILY = checking.account(Rpoint).ENT3_FAMILY
    .ENT4_MONTH = checking.account(Rpoint).ENT4_MONTH
    .ENT5_DAY = checking.account(Rpoint).ENT5_DAY
    .ENT6_YEAR = checking.account(Rpoint).ENT6_YEAR
    .ENT7_DESCRIPTION = checking.account(Rpoint).ENT7_DESCRIPTION
    .ENT8_AMOUNT = checking.account(Rpoint).ENT8_AMOUNT
End With
With checking.account(Rpoint)
    .ENT0_CVD_TYPE = checking.account(0).ENT0_CVD_TYPE
    .ENT1_CLEARED = checking.account(0).ENT1_CLEARED
    .ENT2_NUM = checking.account(0).ENT2_NUM
    .ENT3_FAMILY = checking.account(0).ENT3_FAMILY
    .ENT4_MONTH = checking.account(0).ENT4_MONTH
    .ENT5_DAY = checking.account(0).ENT5_DAY
    .ENT6_YEAR = checking.account(0).ENT6_YEAR
    .ENT7_DESCRIPTION = checking.account(0).ENT7_DESCRIPTION
    .ENT8_AMOUNT = checking.account(0).ENT8_AMOUNT
End With
End Function

Private Sub ChkDep_OPbttn_Click(Index As Integer)
checking.account(EntryPointer).ENT0_CVD_TYPE = Index
End Sub

Private Sub Form_Resize()
List1.Width = account.Width - List1.Left - 400
End Sub

Private Sub listView_Click(Index As Integer)
If listView(Index).BackColor = vb3DFace Then
    listView(Index).BackColor = vb3DHighlight
    Else
    listView(Index).BackColor = vb3DFace
End If
End Sub

Private Sub List1_Click()
EntryPointer = List1.ListIndex + 1
Update_AccountEditor
txtCheckEdit(6).SetFocus
End Sub

Private Sub mnuFile_Click()
Form_Load
End Sub

Private Sub mnuNew_Click()
NewEntry_Click
End Sub

Private Sub NewEntry_Click()
New_CVD_Form.Show
New_CVD_Form.SetFocus
account.Enabled = False
End Sub

Private Sub txtCheckEdit_GotFocus(Index As Integer)
txtCheckEdit(Index).SelStart = 0
txtCheckEdit(Index).SelLength = Len(txtCheckEdit(Index).Text)
End Sub

Private Sub checkboxCleared_Click()
checking.account(EntryPointer).ENT1_CLEARED = checking.account(EntryPointer).ENT1_CLEARED
End Sub
