VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "1 To 10"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Click to test"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Type a number from 1 to 10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  ' Beep or issue error
  Dim intCtr As Integer
  Dim msg As Variant  ' MsgBox() return
  If Val(txtNumber.Text) >= 1 And _
     Val(txtNumber.Text) <= 10 Then
       For intCtr = 1 To Val(txtNumber.Text)
         ' If your computer is fast, you may
         ' not hear a lot of beeps
         Beep
       Next intCtr
       ' Let the user enter another
       txtNumber.Text = ""
  Else
    msg = MsgBox("You did not type a number from 1 to 10", _
        vbExclamation, "Error")
    txtNumber.SetFocus
  End If
  
End Sub

Private Sub Command2_Click()
  End
End Sub
