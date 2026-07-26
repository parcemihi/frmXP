VERSION 5.00
Begin VB.Form Checking 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton UpdateDatabase 
      Caption         =   "&Update"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton EnterData 
      Caption         =   "&Enter"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   750
      TabIndex        =   3
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1050
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check To the order of"
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
      Index           =   3
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Amount"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   7
      Top             =   960
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Date"
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
      Index           =   1
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   240
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer 'My Universal Integer Counter for 'For Loops' and 'Do Loops'
Dim EntryPoint As Integer
Dim ID As Integer
Dim CDB(500) As chkacct
Private Type chkacct
    Chk1Num As String
    Chk2Date As String
    Chk3Amt As Single
    Chk4PayTo As String
End Type

Private Sub EnterData_Click()
Dim msg As Integer
msg = 0
i = 0
Do Until i = ID
    If Text1(0).Text = CDB(i).Chk1Num Then
        msg = MsgBox("Duplicate check number. Do you want to overwrite?", vbYesNoCancel, "Duplicate Entry")
        EntryPoint = i
        i = ID
    Else
        i = i + 1
    End If
Loop

Select Case msg
    Case 0
            CDB(ID).Chk1Num = Text1(0).Text
            CDB(ID).Chk2Date = Text1(1).Text
            CDB(ID).Chk3Amt = Text1(2).Text
            CDB(ID).Chk4PayTo = Text1(3).Text
            ID = ID + 1
    Case 6
            CDB(EntryPoint).Chk1Num = Text1(0).Text
            CDB(EntryPoint).Chk2Date = Text1(1).Text
            CDB(EntryPoint).Chk3Amt = Text1(2).Text
            CDB(EntryPoint).Chk4PayTo = Text1(3).Text

    Case Else
End Select

'If ID = 99 Then
Open App.Path & "\03580946.k" For Output As #1
    For i = 0 To ID - 1
        Write #1, CDB(i).Chk1Num, CDB(i).Chk2Date, CDB(i).Chk3Amt, CDB(i).Chk4PayTo
    Next i
Close #1
ID = 0
'End If






End Sub

Private Sub Form_Load()
'ID = 0
'Open App.Path & "\03580946.k" For Input As #1
'    Do Until EOF(1)
'        Input #1, CDB(ID).Chk1Num, CDB(ID).Chk2Date, CDB(ID).Chk3Amt, CDB(ID).Chk4PayTo
'        ID = ID + 1
'    Loop
'Close #1
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub UpdateDatabase_Click()
Open App.Path & "\03580946.k" For Append As #1
    Write #1, CDB(ID).Chk1Num, CDB(ID).Chk2Date, CDB(ID).Chk3Amt, CDB(ID).Chk4PayTo
Close #1
End Sub
