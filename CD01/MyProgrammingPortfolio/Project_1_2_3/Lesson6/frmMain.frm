VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Ledger"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBalance 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtDebit 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtCredit 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdDeposit 
      Caption         =   "Deposit"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid grdLedger 
      Height          =   3015
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Balance"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Debit"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Credit"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblTotals 
      Caption         =   "Totals"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rsAccess As ADODB.Recordset


Private Sub cmdCheck_Click()
    Load frmCheck
    frmCheck.Show vbModal
    
End Sub

Private Sub cmdDeposit_Click()
    Load frmDeposit
    frmDeposit.Show vbModal
End Sub

Private Sub cmdQuery_Click()
    Dim rsTotal As ADODB.Recordset
    Dim stSQL As String
    
    Screen.MousePointer = vbHourglass
    rsAccess.Requery
    
    '------------------totals---------------------
    'Credit----------------
    stSQL = "select sum(credit) as credit_total from ledger"
    
    Set rsTotal = New ADODB.Recordset
    rsTotal.Open stSQL, objAccessConnection, adOpenKeyset, adLockOptimistic
    
    If Not rsTotal.EOF Then
        txtCredit = rsTotal("credit_total")
    End If
        
    'Debit---------------
    stSQL = "select sum(debit) as debit_total from ledger"
    
    Set rsTotal = New ADODB.Recordset
    rsTotal.Open stSQL, objAccessConnection, adOpenKeyset, adLockOptimistic
    
    If Not rsTotal.EOF Then
        txtDebit = rsTotal("debit_total")
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim sSQL As String
    
    Set objAccessConnection = New ADODB.Connection
    objAccessConnection.CursorLocation = adUseClient
   
    objAccessConnection.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;" & _
    " Data Source=" & AppPath & "quicken.mdb;"
    
    sSQL = "select description,credit,debit,transactiondate,checknumber from ledger"
    
    Set rsAccess = New ADODB.Recordset
    rsAccess.Open sSQL, objAccessConnection, adOpenKeyset, adLockOptimistic

    Set grdLedger.DataSource = rsAccess
    
End Sub

