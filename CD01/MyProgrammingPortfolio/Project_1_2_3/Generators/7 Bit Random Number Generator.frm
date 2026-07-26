VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Text            =   "1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Text            =   "1000"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Record Log"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7 Bit Binary"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Depth"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Interval In Sec"
      Height          =   195
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   1035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gen7bit As String
Dim RecLog As Boolean
Private Sub Command1_Click()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
Else
    Timer1.Enabled = True
End If
End Sub

Private Sub Command2_Click()
If RecLog = True Then
    RecLog = False
Else
    RecLog = True
End If
End Sub

Private Sub Text1_Change()
Timer1.Interval = Val(Text1.Text)
End Sub

Private Sub Timer1_Timer()
Randomize
Dim i As Integer
Dim Pulse As Integer
Pulse = Val(Text2.Text)
List1.Clear
Do Until Pulse <= -0
    
    gen7bit = ""
    Pulse = Pulse - 1
    gen7bit = Chr(Int((256 * Rnd)))
    List1.AddItem gen7bit
    If RecLog = True Then
        Open App.Path & "\gen7bit.log" For Append As #1
        Print #1, gen7bit
        Close #1
    End If
Loop
If RecLog = True Then
    Open App.Path & "\gen7bit.log" For Append As #1
        Print #1,
    Close #1
End If
List1.AddItem ""
End Sub
