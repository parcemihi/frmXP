VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Keith's Message Center"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10470
   Icon            =   "Message.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10470
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9480
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   10
      Text            =   "Select Recipient"
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Message"
      Height          =   615
      Index           =   6
      Left            =   8640
      TabIndex        =   8
      Top             =   5520
      Width           =   1335
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
      Height          =   5080
      Left            =   470
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Message.frx":030A
      Top             =   470
      Width           =   9531
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Previous Message"
      Height          =   615
      Index           =   5
      Left            =   7280
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Message"
      Height          =   615
      Index           =   4
      Left            =   5920
      TabIndex        =   6
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Message"
      Height          =   615
      Index           =   3
      Left            =   4560
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check Messages"
      Height          =   615
      Index           =   2
      Left            =   3200
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Message Timer"
      Height          =   615
      Index           =   1
      Left            =   1840
      TabIndex        =   3
      Top             =   5520
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Message"
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   6240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8520
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.Menu addMember 
      Caption         =   "&Add Members"
   End
   Begin VB.Menu exit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As String, y As String, z As Integer
Dim i As Integer, j As Integer, k As Integer
Dim msgEnd As String
Dim MessageTimer As Integer
Dim TimeStart As Long
Dim TimeEnd As Long
Dim TimeInterval As Long
Dim TimeDay As Integer, TimeMonth As Integer
Private Sub Form_Load()
    Command1(0).Caption = "Delete" & vbCrLf & "Message"
    Command1(1).Caption = "Set Message" & vbCrLf & "Timer"
    Command1(2).Caption = "Check " & vbCrLf & "Messages"
    Command1(3).Caption = "New" & vbCrLf & "Message"
    Command1(4).Caption = "Send" & vbCrLf & "Message"
    Command1(5).Caption = "Previous" & vbCrLf & "Message"
    Command1(6).Caption = "Next" & vbCrLf & "Message"
    Open App.Path & "\members.dat" For Input As #1
    Do Until EOF(1)
    Input #1, x
        Combo1.AddItem x
    Loop
    Close #1
    Open App.Path & "\message.dat" For Input As #1
    Do Until EOF(1)
    Input #1, x
        List1.AddItem x
    Loop
    Close #1
    Text1.Text = List1.Text
    place
    'LoadMessages
End Sub
Private Sub Command2_Click()
    On Error GoTo ErrHandler
    ChDir "C:\WINDOWS\Desktop"
    CommonDialog1.Filter = "KEG Message Files (*.keg)|*.keg|"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.DefaultExt = "c:\windows\desktop"
    CommonDialog1.ShowOpen
    x = (CommonDialog1.FileName)
    Exit Sub
ErrHandler:
   Exit Sub
End Sub
Public Sub place()
    Label2.Caption = "Message " & List1.ListIndex + 1 & " of " & List1.ListCount
End Sub
Private Sub Update_Messages()
    Open App.Path & "\message.dat" For Output As #1
    For i = 0 To List1.ListCount - 1
    Write #1, List1.List(i)
    Next i
    Close #1
    LoadMessages
End Sub
Private Sub addMember_Click()
    Combo1.Clear
    Form1.Hide
    Form2.Show
    Open App.Path & "\members.dat" For Input As #1
    Do Until EOF(1)
    Input #1, x
        Combo1.AddItem x
    Loop
Close #1
End Sub
Private Sub exit_Click()
    Unload Form1
    Unload Form2
End Sub

Private Sub Timer1_Timer()
    If TimeMonth = Month(Date) And TimeDay = Day(Date) Then
    If TimeEnd <= 60 * Hour(Time) + Minute(Time) Then
    TimeEnd = TimeEnd + TimeInterval
    LoadMessages
    End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    Select Case Index
    Case 0
    If Not List1.ListCount = 0 Then
    List1.RemoveItem List1.ListIndex
    Text1.Text = List1.Text
    End If
    Update_Messages
    Case 1
    TimeInterval = InputBox("Enter the number of minutes you would like the message timer to wait before checking for new messages. To disable the timer enter 0.", "Message Timer", 10)
    If TimeInterval = 0 Then
    Timer1.Enabled = False
    Else
    Timer1.Enabled = True
    Timer1.Interval = 60000
    End If
    TimeMonth = Month(Date)
    TimeDay = Day(Date)
    TimeStart = 60 * Hour(Time) + Minute(Time)
    TimeEnd = TimeStart + TimeInterval
    Case 2
    LoadMessages
'    place
    Case 3
    Text1.Text = ""
    Text1.SetFocus
    Case 4
    On Error GoTo ErrorHandle
    Open "\\" & Combo1.Text & "\c\message\" & Combo1.Text & ".dat" For Append As #1
    Write #1, Text1.Text
    Close #1
    Case 5
    If List1.ListIndex > 0 Then
    List1.ListIndex = List1.ListIndex - 1
    End If
    Text1.Text = List1.Text
    Case 6
    If List1.ListIndex < List1.ListCount - 1 Then
    List1.ListIndex = List1.ListIndex + 1
    End If
    Text1.Text = List1.Text
    End Select
'    place
Exit Sub
ErrorHandle:
    msgEnd = MsgBox("You have chosen an invalid member name.", vbCritical)
End Sub
Private Sub LoadMessages()
    z = 0
    Open App.Path & "\message.dat" For Input As #1
        Do Until EOF(1)
        Input #1, x
        z = z + 1
        Loop
    Close #1
    If Not z = List1.ListCount Then
    Beep
    msgEnd = MsgBox("You have a message", vbOKOnly, "Message Alarm")
    List1.Clear
    Open App.Path & "\message.dat" For Input As #1
    Do Until EOF(1)
            Input #1, x
            List1.AddItem x
    Loop
    Close #1
    End If
    List1.ListIndex = List1.ListCount - 1
    Text1.Text = List1.List(List1.ListCount - 1)
    place
End Sub
Private Sub Form_Resize()
    Text1.Width = Form1.Width * 0.9
    Text1.Left = (Form1.Width - Text1.Width) / 2
    Text1.Top = Text1.Left
    Text1.Height = Form1.Height - Text1.Top * 4 - Command1(0)
    Combo1.Top = Text1.Top - Combo1.Height * 1.2
    Combo1.Left = Text1.Left
    For i = 0 To 6
    Command1(i).Top = Text1.Height + Text1.Top
    Command1(i).Width = Text1.Width / 7
    Command1(i).Left = i * Text1.Width / 7 + Text1.Left
    Command1(i).Height = 0.5 * Command1(i).Width
    Next i
    Label2.Left = Text1.Left + Text1.Width - Label2.Width
    Label2.Top = Text1.Top - Label2.Height * 1.2
End Sub
