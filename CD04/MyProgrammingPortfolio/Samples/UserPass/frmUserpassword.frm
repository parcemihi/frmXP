VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form3"
   ScaleHeight     =   2055
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Password"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "User Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim midStart As String, result As String, pass As String
midStart = 0

Text1.Text = ""
Text2.Text = ""


If Len(Form3.Text1.Text) > 0 Then

Do
    midStart = midStart + 1
    result = Asc(Mid(Form3.Text1.Text, midStart, 1))
        If result < 65 Or result > 122 Then
        result = 7
        End If
    pass = pass & result
Loop Until midStart = Len(Form3.Text1.Text)

    Form3.Text2.Text = pass

Else: result = MsgBox("You have not entered a user name.", vbCritical)

End If

End Sub

Private Sub Command2_Click()
    Form3.Hide
    Form1.Show
End Sub
