VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditMenu 
   Caption         =   "Form1"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Item"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Menu Item"
      DataSource      =   "Data1"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Price"
      DataSource      =   "Data1"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   3720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmEditMenu.frx":0000
      OLEDBString     =   $"frmEditMenu.frx":00B0
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Menu Items"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\Pasta Tree Price Program\menurest.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Menu Items"
      Top             =   3240
      Width           =   3255
   End
End
Attribute VB_Name = "frmEditMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Data1.Recordset.AddNew

End Sub

Private Sub Command2_Click()
Data1.Recordset.Update
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\menurest.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\menurest.mdb"


End Sub

Private Sub Form_Activate()

Data1.Recordset.MoveFirst
  Do Until Data1.Recordset.EOF
   Combo1.AddItem Data1.Recordset.Fields(0).Value, x
'   Combo2.AddItem Data1.Recordset.Fields(1).Value, x
   Data1.Recordset.MoveNext
   x = x + 1
 Loop
 Data1.Recordset.MoveFirst
End Sub

