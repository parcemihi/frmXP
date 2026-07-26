VERSION 5.00
Begin VB.Form Solitare 
   Caption         =   "Solitare Score"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   338
      TabIndex        =   2
      Top             =   893
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   338
      TabIndex        =   0
      Top             =   413
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Left            =   1658
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   413
      Width           =   1215
   End
End
Attribute VB_Name = "Solitare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Text2.Text = Val(Text1.Text) + Val(Text2.Text)

Open App.Path & "\sol.k" For Append As #1
    Write #1, Val(Text2.Text)
Close #1

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Form_Load()
Dim loadList As Integer
Open App.Path & "\sol.k" For Input As #1
    Do Until EOF(1)
        Input #1, loadList
    Loop
Close #1
Text2.Text = loadList
End Sub
