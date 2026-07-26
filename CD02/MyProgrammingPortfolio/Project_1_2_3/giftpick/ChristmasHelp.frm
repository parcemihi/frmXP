VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Christmas Gift Giver Picker Help File"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   495
      TabIndex        =   0
      Top             =   735
      Width           =   720
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim msg As String
Private Sub Form_Load()
msg = ""
msg = msg & "For help on a certain button or list," & Chr(13)
msg = msg & "hold the mouse over the object for a " & Chr(13)
msg = msg & "short while and a tool tip will give " & Chr(13)
msg = msg & "you a clue as to what the object does." & Chr(13)
Label1.Caption = msg

End Sub

Private Sub Form_Resize()
With Form2
    .Width = Label1.Width * 1.5
    .Height = Label1.Height * 2
    .Left = (Screen.Width - Form2.Width) / 2
    .Top = (Screen.Height - Form2.Height) / 2
End With
With Label1
    .Top = (Form2.Height - Label1.Height) / 2
    .Left = (Form2.Width - Label1.Width) / 2
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Show
End Sub
