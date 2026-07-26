VERSION 5.00
Begin VB.Form frmBookEntry 
   Caption         =   "Book Title Entry"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add new title"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox cboBooks 
      Height          =   1740
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   1800
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Enter and View Book Titles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "frmBookEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   cboBooks.AddItem cboBooks.Text  ' Adds user's value to the box
   cboBooks.SetFocus
End Sub

Private Sub Command2_Click()
  Unload Me
  End
End Sub
