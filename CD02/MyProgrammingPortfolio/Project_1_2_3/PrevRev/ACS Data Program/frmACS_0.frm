VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmACS_0 
   Caption         =   "ACS Data Program"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Combo1 source"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Desktop\Customer\Customers.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "State"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   5
      Left            =   1320
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   4
      Left            =   1320
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   2
      Left            =   1320
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   1
      Left            =   1320
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "EmployeeID"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   495
      Index           =   4
      Left            =   4920
      TabIndex        =   0
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "Previous"
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "Update"
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "Delete"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "Add"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   480
      Left            =   240
      Top             =   3840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   847
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   12632256
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=acsMain.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=acsMain.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Customers"
      Caption         =   "Customers"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
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
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmACS_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProgLoad As Boolean
Dim fWidth As Double, fHeight As Double
Dim NaviHght As Integer, NaviWdth As Integer
Dim i As Integer, j As Integer, k As Integer
Dim x As Integer, y As Integer, z As Integer
Dim lblWidth As Integer, lblHeight As Integer

Private Sub Form_Activate()
x = 0
Do Until Data1.Recordset.EOF
    For i = 0 To 5
        Combo1(i).AddItem Data1.Recordset.Fields(i).Value, x
    Next i
        x = x + 1
        Data1.Recordset.MoveNext
Loop
Data1.Recordset.MoveFirst

End Sub


Private Sub Form_Load()

Data1.DatabaseName = App.Path & "\acsMain.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\acsMain.mdb"

Label1(0).Caption = "ID Number"
Label1(1).Caption = "First Name"
Label1(2).Caption = "Last Name"
Label1(3).Caption = "Address"
Label1(4).Caption = "City"
Label1(5).Caption = "State"

For i = 0 To 5
    If Label1(i).Width > j Then
    j = Label1(i).Width
    End If
Next i

End Sub

Private Sub Form_Resize()
'Set Window State
If ProgLoad = False Then
    ProgLoad = True
    frmACS_0.WindowState = 2
End If
'Set variables
    x = frmACS_0.Width
    y = frmACS_0.Height
    NaviWdth = x / 8
    NaviHght = y / 8
    z = (Combo1(0).Height - Label1(0).Height) / 2 'For centering labels
'Arrange Navigational Command Bar
For i = 0 To 4
    cmdNavi(i).Width = NaviWdth
    cmdNavi(i).Height = NaviHght
    cmdNavi(i).Top = 0 + y * 0.01
    cmdNavi(i).Left = i * NaviWdth + x * 0.01
Next i
For i = 0 To 5
    Label1(i).Top = (NaviHght + y * 0.02 + z) + i * Label1(0).Height * 1.5
    Label1(i).Left = x * 0.01
    Combo1(i).Width = 3000
    Combo1(i).Text = ""
    Combo1(i).Top = (NaviHght + y * 0.02) + i * Label1(0).Height * 1.5
    Combo1(i).Left = Label1(0).Width + x * 0.02

Next i
End Sub

