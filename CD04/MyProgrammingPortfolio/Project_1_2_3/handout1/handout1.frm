VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2760
      Top             =   4560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      Height          =   3375
      ItemData        =   "handout1.frx":0000
      Left            =   240
      List            =   "handout1.frx":0002
      TabIndex        =   10
      Top             =   510
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   8
      Top             =   3555
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   7
      Top             =   3195
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   6
      Top             =   2835
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   2475
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   2115
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   1755
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Modify"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\Project_1_2_3\handout1\st_grade.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score 5"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   3600
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score 4"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   3240
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score 3"
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
      Left            =   2160
      TabIndex        =   14
      Top             =   2880
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score 2"
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
      Left            =   2160
      TabIndex        =   13
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Score 1"
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
      Left            =   2160
      TabIndex        =   12
      Top             =   2160
      Width           =   690
   End
   Begin VB.Label Label1 
      Height          =   1095
      Left            =   2040
      TabIndex        =   11
      Top             =   510
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x As Integer, i As Integer
Dim CapOn As Boolean
Dim LabCap As String
Dim PR As ADODB.Parameter
Dim CM As ADODB.Command

Dim CN As ADODB.Connection
Dim RS As ADODB.Recordset

Dim DB As StudentGrade
Private Type StudentGrade
    SName(0 To 499) As String * 20
    SGrade(0 To 4, 0 To 499)
End Type

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Command1_Click()
Set RS = New ADODB.Connection
RS.Open "provider=Microsoft.Jet.OLEDB.3.51;data source=" & App.Path & "\st_grade.mdb"
'RS.Open "select * from grades order by description"
RS.AddNew
RS!Name = Text1(0).Text
RS!G1 = Text1(0).Text
RS!G2 = Text1(0).Text
RS!G3 = Text1(0).Text
RS!G4 = Text1(0).Text
RS!G5 = Text1(0).Text

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Form_Load()
x = 0


CapOn = False
Data1.DatabaseName = App.Path & "\st_grade.mdb"
Data1.RecordSource = "grades"
Data1.Refresh

   Do Until Data1.Recordset.EOF
      DB.SName(x) = Data1.Recordset!Name
      List1.AddItem DB.SName(x)
      DB.SGrade(0, x) = Data1.Recordset!G1
      DB.SGrade(1, x) = Data1.Recordset!G2
      DB.SGrade(2, x) = Data1.Recordset!G3
      DB.SGrade(3, x) = Data1.Recordset!G4
      DB.SGrade(4, x) = Data1.Recordset!G5
      Data1.Recordset.MoveNext
      x = x + 1
   Loop





End Sub

Private Sub Form_Resize()
CapOn = True
End Sub


Private Sub List1_Click()
If CapOn = True Then
x = 0
For i = 0 To 4
x = x + DB.SGrade(i, List1.ListIndex)
Next i
x = x / 5
Label1.Caption = List1.List(List1.ListIndex) & Chr(10) & "Average score is " & x
End If
End Sub

Private Sub List1_DblClick()
'
End Sub
