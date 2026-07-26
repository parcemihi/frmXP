VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Message Members"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "&Help"
      Height          =   495
      Left            =   7680
      TabIndex        =   6
      ToolTipText     =   "Displays the help file."
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      ToolTipText     =   "Closes the add member program and starts the message center."
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      ToolTipText     =   "Delete the highlited list item."
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add To List"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Add computer name to the list."
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Computer Name"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   4320
      TabIndex        =   0
      ToolTipText     =   "Click to hide help."
      Top             =   360
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   2295
      Left            =   4200
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Line Line4 
      Visible         =   0   'False
      X1              =   3835
      X2              =   3835
      Y1              =   600
      Y2              =   360
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   4200
      X2              =   3840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   3600
      X2              =   3840
      Y1              =   480
      Y2              =   600
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   3600
      X2              =   3840
      Y1              =   480
      Y2              =   360
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Variant
Dim msg As Variant
Dim MemberOk As Boolean
Dim i As Integer

Private Sub Command1_Click()
Open App.Path & "\members.dat" For Input As #1
Do Until EOF(1)
Input #1, x
    If Text1.Text = x Then
    msg = MsgBox("That computer is already on the list", vbOKOnly)
    Exit Do
    ElseIf Text1.Text = "" Then
    msg = MsgBox("Please enter a valid name", vbOKOnly)
    Exit Do
    
    End If
Loop
Close #1
'    If Text1.Text = "" Then
'    msg = MsgBox("Please enter a valid name", vbOKOnly)
'Else
If Not Text1.Text = "" Then
    Open App.Path & "\members.dat" For Append As #1
    Write #1, Text1.Text
    Close #1
    List1.AddItem Text1.Text
    End If
Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Command2_Click()

On Error GoTo ErrorHandle

If Not List1.ListCount = 0 Then
List1.RemoveItem List1.ListIndex
Text1.Text = ""
Text1.SetFocus

    Open App.Path & "\members.dat" For Output As #1
    For i = 0 To List1.ListCount - 1
    Write #1, List1.List(i)
    Next i
    Close #1
Exit Sub

End If
ErrorHandle:
msg = MsgBox("You must select a member name to delete.", vbCritical)

End Sub
Private Sub Command3_Click()
Unload Form2
Form1.Show
End Sub

Private Sub Command4_Click()
If Label1.Visible = False Then
Label1.Visible = True
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
Shape1.Visible = True
Else
Label1.Visible = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Visible = False
End If
End Sub

Private Sub Form_Load()
Label1.Caption = "Type the name of the computer you would like to add." & vbCrLf & vbCrLf & _
"If you are uncertain of the name of the computer , open the Network Neighborhood icon located on the desktop. There you will find the names of the computer's within your network."
MemberOk = False

Open App.Path & "\members.dat" For Input As #1
Do Until EOF(1)
Input #1, x
List1.AddItem x
Loop
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
End Sub

Private Sub Label1_Click()
Label1.Visible = False
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Shape1.Visible = False

End Sub

