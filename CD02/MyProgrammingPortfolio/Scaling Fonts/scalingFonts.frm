VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KGCommand1 As ButtonHead

Private Type ButtonHead
    height As Integer
    width As Integer
    left As Integer
    top As Integer
End Type

Private Sub Form_Load()
KGCommand1.top = Command1.top
KGCommand1.height = Command1.height
KGCommand1.width = Command1.width
KGCommand1.left = Command1.left
End Sub

Private Sub Form_Resize()
If Form1.MaxButton = False Then
    Form1.width = Form1.height * 1.5
End If

'Command Button Size
Command1.height = Form1.width * 0.1
Command1.width = Command1.height * 2

'Command Button Fontsize
x = Command1.FontSize
x = Command1.height * 0.015
Command1.FontSize = x

'Command Button Centering
Command1.left = Form1.width / 2 - Command1.width / 2
Command1.top = Form1.height / 2 - Command1.height / 2

End Sub
