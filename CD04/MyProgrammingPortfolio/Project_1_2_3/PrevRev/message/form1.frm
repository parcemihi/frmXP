VERSION 5.00
Begin VB.Form Message 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   8355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "New Message"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Messages"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   7815
   End
   Begin VB.Timer Timer1 
      Left            =   7680
      Top             =   8160
   End
End
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Variant

Private Sub Command1_Click()

Open "\\Valued_2\c\Keiths Stuff\message.dat" For Input As #1
  Do Until EOF(1)
  Input #1, x
  Text1.Text = x
  Loop
Close #1


End Sub

Private Sub Form_Load()
Command1.Caption = "Check" & vbCrLf & "Messages"
Command2.Caption = "New" & vbCrLf & "Message"

Text1.Width = .8 *











End Sub
