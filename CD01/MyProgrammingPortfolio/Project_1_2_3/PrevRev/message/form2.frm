VERSION 5.00
Begin VB.Form addComputer 
   Caption         =   "Form2"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&Done"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   5760
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add To List"
      Default         =   -1  'True
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   8445
   End
End
Attribute VB_Name = "addComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Variant
Dim msg As Variant
Dim MemberOk As Boolean

Private Sub Command1_Click()
Open App.Path & "\members.dat" For Input As #1
Do Until EOF(1)
Input #1, x
    If Text1.Text = x Then
    msg = MsgBox("That computer is already on the list", vbOKOnly)
    MemberOk = False
    Exit Do
    ElseIf Text1.Text = "" Then
    msg = MsgBox("Please enter a valid name", vbOKOnly)
    MemberOk = False
    Exit Do
    
    End If
Loop
Close #1
    If MemberOk = True Then
    Open App.Path & "\members.dat" For Append As #1
    Write #1, Text1.Text
    Close #1
    End If
MemberOk = True
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
End Sub
Private Sub Command3_Click()
Unload addComputer
Message.Show
End Sub
Private Sub Form_Load()
Label1.Caption = "Type the name of the computer you would like to add to the message center in the box below." & vbCrLf & vbCrLf & _
"If you are uncertain of the name of the computer you would like to add, open the Network Neighborhood icon located on the desktop. There you will find the names of the computer's within your network."
MemberOk = True
Open App.Path & "\members.dat" For Append As #1
Close #1
End Sub
