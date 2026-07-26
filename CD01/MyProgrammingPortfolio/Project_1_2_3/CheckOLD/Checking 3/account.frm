VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   9090
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "View &Checking"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   4418
      TabIndex        =   11
      Top             =   1433
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3338
      TabIndex        =   10
      Top             =   1433
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2498
      TabIndex        =   9
      Top             =   1433
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1658
      TabIndex        =   8
      Top             =   1433
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cleared"
      Height          =   375
      Left            =   4418
      TabIndex        =   7
      Top             =   953
      Width           =   855
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
      Height          =   6300
      Left            =   818
      TabIndex        =   6
      Top             =   2280
      Width           =   9735
   End
   Begin VB.TextBox checks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   7898
      TabIndex        =   5
      Top             =   473
      Width           =   2775
   End
   Begin VB.TextBox checks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   4418
      TabIndex        =   4
      Top             =   473
      Width           =   1815
   End
   Begin VB.TextBox checks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   818
      TabIndex        =   3
      Top             =   953
      Width           =   3255
   End
   Begin VB.TextBox checks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2258
      TabIndex        =   2
      Top             =   473
      Width           =   1815
   End
   Begin VB.TextBox checks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   818
      TabIndex        =   1
      Top             =   473
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   818
      TabIndex        =   0
      Top             =   1433
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ACT(0 To 2000) As account
Private Type account
    cleared As String * 1
    date As String * 8
    checkNum As Integer
    description As String * 30
    min As Single
    add As Single
    
End Type
Dim TotalEntries As Integer

Dim AccountBalance As Double


Private Sub Form_Load()
TotalEntries = -1
Dim VariLen(0 To 5) As String
Dim i As Integer
Open App.Path & "\03580946.txt" For Input As #1
    Do Until EOF(1)
    TotalEntries = TotalEntries + 1
        Input #1, ACT(TotalEntries).cleared, _
                  ACT(TotalEntries).date, _
                  ACT(TotalEntries).checkNum, _
                  ACT(TotalEntries).description, _
                  ACT(TotalEntries).min, _
                  ACT(TotalEntries).add
                  
                  
        
        List1.AddItem ACT(TotalEntries).date & "   " & ACT(TotalEntries).checkNum & ACT(TotalEntries).description & ACT(TotalEntries).min & ACT(TotalEntries).add
    Loop
Close #1
End Sub


Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
    Case 1
    Case 2
    Case 3
    Case 4
        
End Select

End Sub

Private Sub List1_DblClick()
List1.RemoveItem List1.ListIndex
End Sub
