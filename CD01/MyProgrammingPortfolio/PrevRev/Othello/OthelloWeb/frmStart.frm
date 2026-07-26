VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Othello - Spel starten"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2130
      TabIndex        =   5
      Top             =   1800
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   270
      TabIndex        =   4
      Top             =   1800
      Width           =   1485
   End
   Begin VB.OptionButton optColor 
      Caption         =   "&White"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   2370
      TabIndex        =   3
      Top             =   1260
      Width           =   1425
   End
   Begin VB.OptionButton optColor 
      Caption         =   "&Black"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   870
      TabIndex        =   2
      Top             =   1260
      Value           =   -1  'True
      Width           =   1425
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "With wich color do you want to play ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   645
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   3075
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Black always starts the game."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   870
      Width           =   3075
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmStart.frx":0000
      Top             =   630
      Width           =   480
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ChosenColor As Integer
Public OK As Boolean

Private Sub cmdCancel_Click()
   OK = False: Hide
End Sub

Private Sub cmdOK_Click()
   OK = True: Hide
End Sub

Private Sub Form_Load()
   ChosenColor = 0
End Sub

Private Sub optColor_Click(Index As Integer)
   ChosenColor = Index
End Sub

