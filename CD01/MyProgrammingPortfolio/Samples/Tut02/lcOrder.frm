VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skate-Away Sales"
   ClientHeight    =   5955
   ClientLeft      =   1095
   ClientTop       =   1470
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5955
   ScaleWidth      =   7770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFirstName 
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Ne&w"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N&ext"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   23
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pre&vious"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   22
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox txtCity 
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
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate Order"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   5280
      Width           =   1695
   End
   Begin VB.TextBox txtLastName 
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
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtState 
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
      Left            =   1320
      TabIndex        =   9
      Text            =   "IL"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtZip 
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
      Left            =   3000
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtBlue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
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
      Left            =   3000
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox txtYellow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
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
      Left            =   3000
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
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
      Left            =   6120
      TabIndex        =   17
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
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
      Left            =   6120
      TabIndex        =   19
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Order"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   26
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   27
      Top             =   5280
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   4680
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   873
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
   Begin VB.Line Line1 
      X1              =   240
      X2              =   3960
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "First &Name:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1020
   End
   Begin VB.Image imgLogo 
      Height          =   495
      Left            =   240
      Picture         =   "lcOrder.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skate-Away Sales Order Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   960
      TabIndex        =   28
      Top             =   360
      Width           =   3210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Address:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ci&ty:"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "&State:"
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
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   510
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "&Zip:"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Blue skateboards ordered:"
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
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   2400
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "&Yellow skateboards ordered:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   2595
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total price:"
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
      Left            =   4320
      TabIndex        =   31
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Skateboard pr&ice:"
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
      Left            =   4320
      TabIndex        =   16
      Top             =   1080
      Width           =   1605
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales tax &rate:"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   1560
      Width           =   1260
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total skateboards:"
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
      Left            =   4320
      TabIndex        =   29
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblTprice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
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
      Height          =   495
      Left            =   6120
      TabIndex        =   32
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblTboards 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   495
      Left            =   6120
      TabIndex        =   30
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SkateClient As ADODB.Recordset


Private Sub cmdDelete_Click()
    SkateClient.Delete (adAffectCurrent)
    SkateClient.MoveNext
End Sub

Private Sub Form_Load()
   Dim strPath As String
   strPath = "A:\Visual Basic\_Tut_02\skateboard.mdb"
   With Adodc1
      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
      "Persist Security Info=False;Data Source=" & strPath & _
      "; Mode=Read|Write"
      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .CommandType = adCmdText
      .RecordSource = "SELECT * " & _
       "FROM SkateClient ORDER BY LastName, FirstName"
      .Refresh
   
   End With

   With txtLastName
      .DataField = "LastName"
      Set .DataSource = Adodc1
   End With

   With txtFirstName
      .DataField = "FirstName"
      Set .DataSource = Adodc1
   End With

   With txtAddress
      .DataField = "Address"
      Set .DataSource = Adodc1
   End With

   With txtCity
      .DataField = "City"
      Set .DataSource = Adodc1
   End With

   With txtState
      .DataField = "State"
      Set .DataSource = Adodc1
   End With

   With txtZip
      .DataField = "Zip"
      Set .DataSource = Adodc1
   End With

   ' Recordset object.
      Set SkateClient = Adodc1.Recordset
      

   ' CommandButton captions.
   Command1.Caption = "Next"
   Command2.Caption = "Previous"
   cmdNew.Caption = "New"
   cmdDelete.Caption = "Delete"
End Sub

Private Sub Command1_Click()
   ' move through the recordset.
   If Not SkateClient.EOF Then
      SkateClient.MoveNext
   Else
      SkateClient.MoveFirst
   End If
End Sub

Private Sub Command2_Click()
   If Not SkateClient.BOF Then
      SkateClient.MovePrevious
   Else
      SkateClient.MoveLast
   End If
End Sub
Private Sub cmdNew_Click()
    Dim strLastName As String
    Dim strFirstName As String
    Dim strAddress As String
    Dim strCity As String
    Dim strState As String
    Dim strZip As String
    
' Sets new entry (otherwise it will overwrite)
    SkateClient.AddNew

    strLastName = Trim(InputBox( _
        "Enter the last name:"))
    strFirstName = Trim(InputBox( _
        "Enter the first name:"))
    strAddress = Trim(InputBox( _
        "Enter the address:"))
    strCity = Trim(InputBox( _
        "Enter the city:"))
    strState = Trim(InputBox( _
        "Enter the state:"))
    strZip = Trim(InputBox( _
        "Enter the zip:"))


txtLastName = strLastName
txtFirstName = strFirstName
txtAddress = strAddress
txtCity = strCity
txtState = strState
txtZip = strZip


SkateClient.UpdateBatch


'RecordSource = "SELECT * " & _
'       "FROM SkateClient ORDER BY LastName, FirstName"
'Form_Load
End Sub


'-------------------------------------------------------------------

Private Sub cmdCalc_Click()
    'Calculate total number of skateboards
    lblTboards.Caption = Val(txtBlue.Text) + Val(txtYellow.Text)
    'Calculate total price
    lblTprice.Caption = Val(lblTboards.Caption) * Val(txtPrice.Text) * (1 + Val(txtRate.Text))
    txtPrice.Text = Format(txtPrice.Text, "currency")
    lblTprice.Caption = Format(lblTprice.Caption, "currency")
    cmdPrint.SetFocus
    
End Sub

Private Sub cmdClear_Click()
    'Clear the screen for the next order
    
If Not SkateClient.BOF Then
        SkateClient.MoveFirst
    End If
        SkateClient.MovePrevious
    
    
    
'    txtName.Text = ""
'    txtAddress.Text = ""
'    txtCity.Text = ""
'    txtState.Text = ""
'    txtZip.Text = ""
    txtBlue.Text = ""
    txtYellow.Text = ""
    lblTboards.Caption = ""
    lblTprice.Caption = ""
    txtLastName.SetFocus
    
End Sub

Private Sub cmdExit_Click()
    End
    
End Sub

Private Sub cmdPrint_Click()
'Hide the command button before printing the form
    cmdCalc.Visible = False
    cmdPrint.Visible = False
    cmdClear.Visible = False
    cmdExit.Visible = False
    PrintForm
    
'Display the command button after printing the form
    cmdCalc.Visible = True
    cmdPrint.Visible = True
    cmdClear.Visible = True
    cmdExit.Visible = True
    
End Sub
