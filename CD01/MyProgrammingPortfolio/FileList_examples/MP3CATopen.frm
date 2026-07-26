VERSION 5.00
Begin VB.Form MP3CATopen 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Text            =   " *.dat"
      Top             =   2400
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3135
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Load File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "MP3CATopen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
mp3_cat_dat = Text1.Text

End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    File1.Pattern = "*.dat"
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    Text1.Text = File1.FileName
End Sub

Private Sub Form_Load()
    Dir1.Path = App.Path & "mp3_cat"
End Sub
