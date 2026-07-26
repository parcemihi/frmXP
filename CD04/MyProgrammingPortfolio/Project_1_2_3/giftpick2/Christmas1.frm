VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Christmas Gift Giver Picker"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Stealth Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "Off"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "On"
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
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear &List"
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
      Left            =   7920
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Create List"
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
      Left            =   6360
      TabIndex        =   4
      ToolTipText     =   "Click here to find out who you will give your gift to."
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      ItemData        =   "Christmas1.frx":0000
      Left            =   3960
      List            =   "Christmas1.frx":0002
      TabIndex        =   3
      ToolTipText     =   "This box displays the list of people and to whom they will be awarded the pleasure of giving a gift to."
      Top             =   960
      Width           =   5415
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
      Height          =   360
      Left            =   600
      MaxLength       =   21
      TabIndex        =   0
      ToolTipText     =   "Enter a name to be added to the Christmas gift giver list."
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add to List"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5580
      ItemData        =   "Christmas1.frx":0004
      Left            =   600
      List            =   "Christmas1.frx":0006
      TabIndex        =   1
      ToolTipText     =   "To remove a name, 'Click' the name and press the 'Delete' key."
      Top             =   960
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewList 
         Caption         =   "&New List"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenList 
         Caption         =   "&Open List"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveList 
         Caption         =   "&Save List"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveListAs 
         Caption         =   "&Save List As"
         Shortcut        =   ^L
      End
      Begin VB.Menu Space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu Space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ff As String

Dim CL() As ChristmasList
Private Type ChristmasList
    NameNo As Integer
    NameList As String
End Type

Dim RndNameNum As Integer
Dim NameOK As Boolean
Dim RetryNo As Integer
Dim chrList As String
Dim msg As Integer
Dim i As Integer, x As Integer

Private Sub Command2_Click()
Dim FNL As String
ReDim CL(0 To List1.ListCount)
Randomize
i = 0
List2.Clear
If Not List1.ListCount <= 1 Then
    Do Until i = List1.ListCount
        NameOK = True
        RndNameNum = Int((List1.ListCount * Rnd) + 1)
        For x = 0 To i
            If CL(i).NameNo = RndNameNum Or CL(i).NameNo - 1 = i Then
                NameOK = False
            End If
        Next x
    
        If NameOK = True Then
            CL(i).NameNo = RndNameNum
            FNL = List1.List(i)
            Do Until Len(FNL) = 21
                FNL = FNL & " "
            Loop
            CL(i).NameList = FNL
            List2.AddItem CL(i).NameList & FNL & CL(i).NameNo

'            List2.AddItem FNL & List1.List(RndNameNum - 1)
            i = i + 1
        End If
        RetryNo = RetryNo + 1
    
        If RetryNo >= List1.ListCount * 4 Then
            ReDim Names(0 To List1.ListCount)
            i = 0
            List2.Clear
            RetryNo = 0
        End If
    Loop
    Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
List1.AddItem Text1.Text
Text1_GotFocus
End Sub

Private Sub Form_Load()
ChDir App.Path
Text1.Text = "Enter Name"
Command1.Default = True
End Sub

Private Sub List1_DblClick()
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    List1.RemoveItem (List1.ListIndex)
End If

End Sub


Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHelp_Click()
'Load frmBrowser
Form1.Hide
'frmAbout.Show
frmBrowser.Show
End Sub

Private Sub mnuNewList_Click()
List1.Clear
List2.Clear
Text1.Text = "Enter Name"
Text1_GotFocus
Text1.SetFocus
End Sub

Private Sub mnuOpenList_Click()
On Error GoTo ErrHandler
List1.Clear
List2.Clear
CommonDialog1.Filter = "Christmas Lists (*.chr)|*.chr"
CommonDialog1.Flags = &H1000
CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
        Input #1, chrList
        List1.AddItem chrList
        
'//Set for stealth mode
        Input #1, chrList
        List2.AddItem chrList
'//End stealth mode
    
    Loop
Close #1
ff = CommonDialog1.FileName
Text1.SetFocus
Form1.Caption = "Christmas Gift Giver Picker.  Current file: " & CommonDialog1.FileTitle
Text1.Text = "Enter Name"
Text1_GotFocus
Text1.SetFocus
Exit Sub

ErrHandler:
Exit Sub
End Sub

Private Sub mnuPrint_Click()
    Printer.PrintQuality = -1
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.Print " "
    Printer.Print " "
    Printer.Print " "
    Printer.Print "            **Keith's Christmas List People Picker**"
    Printer.Print " "
    Printer.Print "            Giver                Givee"
    Printer.Print "            ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯  ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯"
        For i = 0 To List1.ListCount
        chrList = "            " & List2.List(i)
            If i >= 20 Then
                Printer.NewPage
            End If
        Printer.Print chrList
        Next i

If vbNo = MsgBox("Print this fine document?", vbYesNo) Then
    Printer.KillDoc
Else
    Printer.EndDoc
End If

End Sub

Private Sub mnuSaveList_Click()
If ff = "" Then
mnuSaveListAs_Click
Else
Open ff For Output As #1
    For x = 0 To List1.ListCount - 1
        Write #1, List1.List(x), List2.List(x)
    Next x
Close #1
End If
Text1_GotFocus
Text1.SetFocus
End Sub

Private Sub mnuSaveListAs_Click()
On Error GoTo ErrHandler
CommonDialog1.Filter = "Christmas Lists (*.chr)|*.chr"
CommonDialog1.Flags = &H2
CommonDialog1.ShowSave
ff = CommonDialog1.FileName
    Open CommonDialog1.FileName For Output As #1
        For x = 0 To List1.ListCount - 1
            Write #1, List1.List(x), List2.List(x)
        Next x
    Close #1
Text1.SetFocus
Exit Sub
ErrHandler:
Exit Sub


End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
