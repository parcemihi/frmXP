VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "User Password"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu password 
         Caption         =   "Get &Password"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim midStart As String, result As String, pass As String
midStart = 0

Do
    midStart = midStart + 1
        result = Asc(Mid(Form1.Text1.Text, midStart, 1))
        pass = pass & result

Loop Until midStart = Len(Form1.Text1.Text)

If pass = Form1.Text2.Text Then
    Form2.Show
Else: result = MsgBox("That is not the correct password", vbCritical)
End If



End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub password_Click()
    Form1.Hide
    Form3.Show
End Sub
