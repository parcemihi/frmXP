VERSION 5.00
Begin VB.Form lib_vhs 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   24
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   23
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   22
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   975
      Left            =   6360
      TabIndex        =   19
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton navigation 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   6720
      TabIndex        =   11
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton navigation 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5760
      TabIndex        =   10
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton navigation 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4800
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Rollback"
      Height          =   975
      Left            =   5160
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "lib_vhs.frx":0000
      Left            =   3960
      List            =   "lib_vhs.frx":0002
      TabIndex        =   7
      Text            =   " - Subject Topic "
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton navigation 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton updateDB 
      Caption         =   "&Update (Commit)"
      Height          =   975
      Left            =   3960
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton deleteDBR 
      Caption         =   "&Delete"
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton addDBR 
      Caption         =   "&Add Record"
      Height          =   975
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox vhs 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   1560
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "Subject"
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
      Index           =   6
      Left            =   3120
      TabIndex        =   20
      Top             =   360
      Width           =   675
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "ID Number"
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
      Index           =   0
      Left            =   480
      TabIndex        =   18
      Top             =   360
      Width           =   960
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "Location"
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
      Index           =   5
      Left            =   675
      TabIndex        =   16
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "Quantity"
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
      Index           =   4
      Left            =   720
      TabIndex        =   15
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "Format / Time"
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
      Index           =   3
      Left            =   4035
      TabIndex        =   14
      Top             =   1680
      Width           =   1245
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "Copyright Date"
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
      Index           =   2
      Left            =   105
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label VHS_Field 
      AutoSize        =   -1  'True
      Caption         =   "Title"
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
      Index           =   1
      Left            =   1050
      TabIndex        =   12
      Top             =   720
      Width           =   390
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Add New Record"
         Index           =   0
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Delete Record"
         Index           =   1
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Update Database"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuNavi 
         Caption         =   "&Navigation"
         Begin VB.Menu mnuFirst 
            Caption         =   "&First       "
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuBack 
            Caption         =   "&Back"
            Shortcut        =   ^J
         End
         Begin VB.Menu mnuNext 
            Caption         =   "&Next"
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuLast 
            Caption         =   "&Last"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort By"
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Title"
            Checked         =   -1  'True
            Index           =   0
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&ID#"
            Index           =   1
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Copyright Date"
            Index           =   2
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Running Time"
            Index           =   3
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Location"
            Index           =   4
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Subject"
            Index           =   5
         End
      End
      Begin VB.Menu Space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubject 
         Caption         =   "Subject &Editor"
      End
   End
End
Attribute VB_Name = "lib_vhs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotalVideos As Integer
Dim RNDBPointer As Integer
Dim VideoDB() As RNVHSDB
Private Type RNVHSDB
    ID As String
    Quantity As String
    Title As String
    Copyright As Integer
    MediaFormat As String
    RunningTime As String
    Location As String
End Type

Private Sub Combo1_LostFocus()
'    Combo1.ListIndex = Combo1.Index
End Sub

Private Sub deleteDBR_Click()
'

















End Sub

Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim Subject As String
ReDim VideoDB(0 To 500)
RNDBPointer = 0
TotalVideos = 1

Open App.Path & "\subject.dat" For Input As #1
    Do Until EOF(1)
        Input #1, Subject
        Combo1.AddItem Subject
    Loop
Close #1

Open App.Path & "\rnvhsdb.dat" For Input As #1
    Do Until EOF(1)
        Input #1, VideoDB(TotalVideos).ID, _
                  VideoDB(TotalVideos).Title, _
                  VideoDB(TotalVideos).Copyright, _
                  VideoDB(TotalVideos).MediaFormat, _
                  VideoDB(TotalVideos).RunningTime, _
                  VideoDB(TotalVideos).Quantity, _
                  VideoDB(TotalVideos).Location
        TotalVideos = TotalVideos + 1
    Loop
Close #1

'TotalVideos = TotalVideos - 1
'ReDim Preserve VideoDB(0 To TotalVideos)
navigation_Click (0)
Exit Sub
ErrorHandler:
Close #1
Exit Sub
End Sub


Private Sub mnuSubject_Click()
lib_vhs.Hide
editor.Show

End Sub

Private Sub navigation_Click(Index As Integer)
Select Case Index
    Case 0
    RNDBPointer = 1
    Case 1
        If RNDBPointer = 1 Then
            RNDBPointer = TotalVideos - 1
        Else
            RNDBPointer = RNDBPointer - 1
        End If
    Case 2
        If RNDBPointer = TotalVideos - 1 Then
            RNDBPointer = 1
        Else
            RNDBPointer = RNDBPointer + 1
        End If
    Case 3
    RNDBPointer = TotalVideos - 1
    Case 4
    
    
End Select


vhs(0).Text = VideoDB(RNDBPointer).ID
vhs(1).Text = VideoDB(RNDBPointer).Title
vhs(2).Text = VideoDB(RNDBPointer).Copyright
vhs(3).Text = VideoDB(RNDBPointer).MediaFormat
vhs(4).Text = VideoDB(RNDBPointer).RunningTime
vhs(5).Text = VideoDB(RNDBPointer).Quantity
vhs(6).Text = VideoDB(RNDBPointer).Location




End Sub



'// Database Record functions

Private Sub addDBR_Click()
RNDBPointer = TotalVideos  '// This is the value to send to sort
TotalVideos = TotalVideos + 1
navigation_Click (4)
Dim i As Integer
For i = 0 To 5
    vhs(i).Text = ""
Next i
End Sub


Private Sub updateDB_Click()
'Dim i As Integer
'Open App.Path & "\rnvhsdb.ful" For Output As #1
'    For i = 0 To TotalVideos
'        Write #1, i, VideoDB(i).Title, VideoDB(i).Copyright, VideoDB(i).MediaFormat, VideoDB(i).RunningTime, VideoDB(i).Quantity, VideoDB(i).Location
'    Next i
'Close #1
End Sub

'// Filter Features

Private Sub mnuSortBy_Click(Index As Integer)
Dim i As Integer
For i = 0 To 5
    If i = Index Then
        mnuSortBy(i).Checked = True
    Else
        mnuSortBy(i).Checked = False
    End If
Next i
End Sub

Private Sub vhs_GotFocus(Index As Integer)
vhs(Index).SelStart = 0
vhs(Index).SelLength = Len(vhs(Index))
End Sub

Private Sub vhs_LostFocus(Index As Integer)
    VideoDB(RNDBPointer).ID = Val(vhs(0).Text)
    VideoDB(RNDBPointer).Title = vhs(1).Text
    VideoDB(RNDBPointer).Copyright = Val(vhs(2).Text)
    VideoDB(RNDBPointer).MediaFormat = vhs(3).Text
    VideoDB(RNDBPointer).RunningTime = vhs(4).Text
    VideoDB(RNDBPointer).Quantity = vhs(5).Text
    VideoDB(RNDBPointer).Location = vhs(6).Text
End Sub
