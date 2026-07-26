VERSION 5.00
Begin VB.Form deposits 
   Caption         =   "Deposit"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form2"
   ScaleHeight     =   2325
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox depDate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1575
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
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
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
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox depCashBack 
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
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox depAmount 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Left            =   2220
      TabIndex        =   8
      Top             =   180
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   120
      Picture         =   "deposits.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cash Back"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
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
      Left            =   1980
      TabIndex        =   6
      Top             =   660
      Width           =   675
   End
End
Attribute VB_Name = "deposits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEnter_Click()

'Enter Data Into Database

    depDate = Date
    depAmount.Text = ""
    depCashBack.Text = ""
    depDate.SetFocus
End Sub




Private Sub depdate_GotFocus()
    depDate.SelStart = 0
    depDate.SelLength = Len(depDate)
End Sub
Private Sub depAmount_GotFocus()
    depAmount.SelStart = 0
    depAmount.SelLength = Len(depAmount)
End Sub
Private Sub depCashBack_GotFocus()
    depCashBack.SelStart = 0
    depCashBack.SelLength = Len(depCashBack)
End Sub




Private Sub cmdExit_Click()
    deposits.Hide
    main.Show
End Sub
Private Sub Form_Unload(Cancel As Integer)
    main.Show
End Sub

