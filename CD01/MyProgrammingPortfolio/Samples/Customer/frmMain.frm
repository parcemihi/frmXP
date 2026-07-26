VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customers"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   480
      Left            =   180
      Top             =   4305
      Width           =   4470
      _ExtentX        =   7885
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=Customers.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=Customers.mdb"
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
   Begin VB.Data Data1 
      Caption         =   "Combo1 source"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\ACS\Customer\Customers.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3330
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customers"
      Top             =   615
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "New"
      Height          =   495
      Index           =   4
      Left            =   5715
      Picture         =   "frmMain.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Dodaj novi zapis"
      Top             =   4300
      Width           =   735
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Exit"
      Height          =   495
      Index           =   6
      Left            =   7275
      Picture         =   "frmMain.frx":0086
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4300
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "Delete"
      Height          =   495
      Index           =   5
      Left            =   6555
      Picture         =   "frmMain.frx":011A
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4300
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Fax"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   10
      Left            =   5490
      TabIndex        =   22
      Top             =   3705
      Width           =   2250
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   9
      Left            =   2175
      TabIndex        =   20
      Top             =   3675
      Width           =   2340
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Country"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   8
      Left            =   5490
      TabIndex        =   18
      Top             =   3300
      Width           =   2250
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PostalCode"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   7
      Left            =   2175
      TabIndex        =   16
      Top             =   3255
      Width           =   2340
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Region"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   6
      Left            =   5490
      TabIndex        =   14
      Top             =   2910
      Width           =   2250
   End
   Begin VB.TextBox txtFields 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   5
      Left            =   2175
      TabIndex        =   12
      Top             =   2880
      Width           =   2355
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   4
      Left            =   2175
      TabIndex        =   10
      Top             =   2430
      Width           =   4890
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactTitle"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   3
      Left            =   5055
      TabIndex        =   8
      Top             =   1995
      Width           =   1995
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactName"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   2
      Left            =   2175
      TabIndex        =   6
      Top             =   1980
      Width           =   2160
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CompanyName"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   1
      Left            =   2175
      TabIndex        =   4
      Top             =   1605
      Width           =   3120
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CustomerID"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   2175
      TabIndex        =   2
      Top             =   1215
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   270
      TabIndex        =   0
      Top             =   615
      Width           =   2940
   End
   Begin VB.Label RecNo 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6045
      TabIndex        =   27
      Top             =   15
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Find Customer ID:"
      Height          =   285
      Left            =   270
      TabIndex        =   26
      Top             =   315
      Width           =   1665
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fax:"
      Height          =   255
      Index           =   10
      Left            =   4815
      TabIndex        =   21
      Top             =   3720
      Width           =   435
   End
   Begin VB.Label lblLabels 
      Caption         =   "Phone:"
      Height          =   255
      Index           =   9
      Left            =   285
      TabIndex        =   19
      Top             =   3705
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Country:"
      Height          =   255
      Index           =   8
      Left            =   4740
      TabIndex        =   17
      Top             =   3300
      Width           =   705
   End
   Begin VB.Label lblLabels 
      Caption         =   "Postal code:"
      Height          =   255
      Index           =   7
      Left            =   255
      TabIndex        =   15
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Region:"
      Height          =   255
      Index           =   6
      Left            =   4815
      TabIndex        =   13
      Top             =   2940
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "City:"
      Height          =   255
      Index           =   5
      Left            =   255
      TabIndex        =   11
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Address:"
      Height          =   255
      Index           =   4
      Left            =   255
      TabIndex        =   9
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Title:"
      Height          =   255
      Index           =   3
      Left            =   4575
      TabIndex        =   7
      Top             =   1995
      Width           =   555
   End
   Begin VB.Label lblLabels 
      Caption         =   "ContactName:"
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   5
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CompanyName:"
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   3
      Top             =   1605
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Customer ID:"
      Height          =   255
      Index           =   0
      Left            =   255
      TabIndex        =   1
      Top             =   1215
      Width           =   1815
   End
End
Attribute VB_Name = "frmCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Customers 1.0 Beta
'Code by Domagoj Barisic (domki@usa.net)
'August 21, 2000 (22:18)
'Made in Croatia (Hrvatska)
'-------------------------------------------
Option Explicit
Dim x

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error Resume Next
RecNo.Caption = Adodc1.Recordset.Bookmark & "/" & Adodc1.Recordset.RecordCount
End Sub
Private Sub cmdNav_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
Adodc1.Recordset.MoveFirst
Case 1
Adodc1.Recordset.MovePrevious
Case 2
Adodc1.Recordset.MoveNext
Case 3
Adodc1.Recordset.MoveLast
Case 4
Adodc1.Recordset.AddNew
txtFields(0).SetFocus
Case 5
Adodc1.Recordset.Delete
Case 6
'Exit
Unload Me
End Select
End Sub
Private Sub Combo1_Change()
On Error Resume Next
Adodc1.Recordset.Find "CustomerID  = '" & Combo1.Text & "'", , , 0
End Sub
Private Sub Combo1_Click()
Combo1_Change
End Sub
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1_Change
End Sub
Private Sub Combo1_LostFocus()
Combo1_Change
End Sub
Private Sub Form_Activate()
Data1.Recordset.MoveFirst
  Do Until Data1.Recordset.EOF
   Combo1.AddItem Data1.Recordset.Fields(0).Value, x
   Data1.Recordset.MoveNext
   x = x + 1
 Loop
 Data1.Recordset.MoveFirst
End Sub
Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\Customers.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Customers.mdb"
End Sub
