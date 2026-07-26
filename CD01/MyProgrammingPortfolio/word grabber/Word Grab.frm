VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCompact 
      Caption         =   "Compact &List"
      Height          =   495
      Left            =   2760
      TabIndex        =   28
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "P&rocess File"
      Height          =   495
      Left            =   1440
      TabIndex        =   27
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "&Open File"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton alpha 
      Caption         =   "Z"
      Height          =   375
      Index           =   25
      Left            =   9120
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "Y"
      Height          =   375
      Index           =   24
      Left            =   8760
      TabIndex        =   24
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "X"
      Height          =   375
      Index           =   23
      Left            =   8400
      TabIndex        =   23
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "W"
      Height          =   375
      Index           =   22
      Left            =   8040
      TabIndex        =   22
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "V"
      Height          =   375
      Index           =   21
      Left            =   7680
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "U"
      Height          =   375
      Index           =   20
      Left            =   7320
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "T"
      Height          =   375
      Index           =   19
      Left            =   6960
      TabIndex        =   19
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "S"
      Height          =   375
      Index           =   18
      Left            =   6600
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "R"
      Height          =   375
      Index           =   17
      Left            =   6240
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "Q"
      Height          =   375
      Index           =   16
      Left            =   5880
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "P"
      Height          =   375
      Index           =   15
      Left            =   5520
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "O"
      Height          =   375
      Index           =   14
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "N"
      Height          =   375
      Index           =   13
      Left            =   4800
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "M"
      Height          =   375
      Index           =   12
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "L"
      Height          =   375
      Index           =   11
      Left            =   4080
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "K"
      Height          =   375
      Index           =   10
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "J"
      Height          =   375
      Index           =   9
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "I"
      Height          =   375
      Index           =   8
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "H"
      Height          =   375
      Index           =   7
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "G"
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "F"
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "E"
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "D"
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "C"
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "B"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton alpha 
      Caption         =   "A"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim FileName As String
Dim FileData(5000) As String

Private Sub cmdCompact_Click()
Dim DataNum As Integer, Complete As Boolean, i As Integer, x As String, Duplicate(5000) As Boolean
Dim stepper As Integer

Open App.Path & "\worddata.dat" For Input As #1
    Do Until EOF(1)
        DataNum = DataNum + 1
        Input #1, FileData(DataNum)
    Loop
Close #1
Complete = False
Do Until Complete = True
Complete = True
i = 0
    Do While i < DataNum
        i = i + 1
        If FileData(i) > FileData(i + 1) Then
            FileData(0) = FileData(i)
            FileData(i) = FileData(i + 1)
            FileData(i + 1) = FileData(0)
            Complete = False
        End If
        
        
    Loop
Loop

For i = 1 To DataNum - 1
    If FileData(i) = FileData(i + 1) Then
        Duplicate(i + 1) = True
    End If
Next i

i = 0

Do Until i = DataNum
    i = i + 1
        If Duplicate(i) = True Then
            Do Until Duplicate(stepper) = False
                stepper = stepper + 1
            Loop
            FileData(i) = FileData(stepper)
        End If
Next i



x = FileData(1): x = FileData(2): x = FileData(3): x = FileData(4): x = FileData(5): x = FileData(6): x = FileData(7): x = FileData(8): x = FileData(9): x = FileData(10): x = FileData(11): x = FileData(12)
End Sub





Private Sub cmdProcess_Click()
Dim Buff As String
Dim OneChar As String * 1
Dim DataNum As Integer

Open FileName For Input As #1
    Do Until EOF(1)
        Do While Not EOF(1)
            OneChar = Input(1, #1)
            If OneChar = " " Or OneChar = "." Or EOF(1) Then
                If EOF(1) Then
                    Buff = Buff & OneChar
                End If
                DataNum = DataNum + 1
                FileData(DataNum) = Buff
                Buff = ""
                If DataNum = 5 Then
                    Open App.Path & "\worddata.dat" For Append As #2
                        Do Until DataNum = 0
                            Write #2, FileData(DataNum)
                            DataNum = DataNum - 1
                        Loop
                    Close #2
                End If
            Else
                Buff = Buff & OneChar
            End If
        Loop
    Loop
Close #1
End Sub



Private Sub cmdOpenFile_Click()
ChDir App.Path
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
FileName = CommonDialog1.FileName
Exit Sub
ErrHandler:
    Exit Sub
End Sub

