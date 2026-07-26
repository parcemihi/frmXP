VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Checking 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton datFirst 
      Caption         =   "&First"
      Height          =   495
      Left            =   600
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton datLast 
      Caption         =   "&Last"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton datPrevious 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton datNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton datAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "PayTo"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "Amount"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      DataField       =   "CheckNumber"
      DataSource      =   "Adodc1"
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   960
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   360
      Top             =   6360
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   873
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
      Connect         =   $"Checking.frx":0000
      OLEDBString     =   $"Checking.frx":009E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Withdrawl"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount:"
      Height          =   195
      Index           =   3
      Left            =   3600
      TabIndex        =   7
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date:"
      Height          =   195
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check No:"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Paid to:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1410
      Width           =   540
   End
End
Attribute VB_Name = "Checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CHK As ADODB.Recordset

Private Sub datAdd_Click()
    CHK.AddNew
    
End Sub

Private Sub datFirst_Click()
If Not CHK.BOF Then
    CHK.MoveFirst
End If
End Sub

Private Sub datLast_Click()
If Not CHK.EOF Then
    CHK.MoveLast
End If
End Sub

Private Sub datNext_Click()
If Not CHK.EOF Then
    CHK.MoveNext
End If
End Sub

Private Sub datPrevious_Click()
If Not CHK.BOF Then
    CHK.MovePrevious
End If
End Sub

Private Sub Form_Load()
Set CHK = Adodc1.Recordset

End Sub

Private Sub Text2_Change()
     
End Sub
