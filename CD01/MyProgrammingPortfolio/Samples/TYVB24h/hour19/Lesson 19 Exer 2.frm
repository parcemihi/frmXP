VERSION 5.00
Begin VB.Form frmShape 
   Caption         =   "Special Shape"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   3855
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         FillColor       =   &H000000FF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   2535
         Left            =   0
         Top             =   0
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "(Hint: The key to the green background is a green frame.)"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   4455
   End
End
Attribute VB_Name = "frmShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  Unload Me
  End
End Sub
