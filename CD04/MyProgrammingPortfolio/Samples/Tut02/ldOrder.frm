VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   2640
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsProducts As ADODB.Recordset

Private Sub Form_Load()
   Dim strPath As String
   ' Change this path if needed.
   strPath = "skateboard.mdb"
   With Adodc1
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
      "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .CommandType = adCmdText
      .RecordSource = "SELECT * FROM Products"
      .Refresh
   End With
   
   With Text1
      .DataField = "ProductName"
      Set .DataSource = Adodc1
   End With
   With Text2
      .DataField = "QuantityPerUnit"
      Set .DataSource = Adodc1
   End With
   
   ' Set the Recordset of the ADODC to the
   ' Recordset object.
   Set rsProducts = Adodc1.Recordset
      
   ' Change CommandButton captions.
   Command1.Caption = "MoveNext"
   Command2.Caption = "MovePrevious"
End Sub

Private Sub Command1_Click()
   ' Use the rsProducts recordset object,
   ' and move through the recordset.
   If Not rsProducts.EOF Then
      rsProducts.MoveNext
   Else
      rsProducts.MoveFirst
   End If
End Sub

Private Sub Command2_Click()
   If Not rsProducts.BOF Then
      rsProducts.MovePrevious
   Else
      rsProducts.MoveLast
   End If
End Sub

