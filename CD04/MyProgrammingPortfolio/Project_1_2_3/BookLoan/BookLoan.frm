VERSION 5.00
Begin VB.Form BookLoan 
   Caption         =   "Registered Nursing - Book Loan Database"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5880
      TabIndex        =   32
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "BookTitle"
      DataSource      =   "RNBookLoan"
      Height          =   315
      ItemData        =   "BookLoan.frx":0000
      Left            =   5640
      List            =   "BookLoan.frx":0002
      TabIndex        =   31
      Text            =   "Combo1"
      Top             =   240
      Width           =   3735
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   5760
      TabIndex        =   30
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Data RNBookLoan 
      Caption         =   "Registered Nursing Book Loan Information"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\Project_1_2_3\BookLoan\bookloan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "bookloan"
      Top             =   5280
      Width           =   5220
   End
   Begin VB.CommandButton ending 
      Caption         =   "Move to &Last"
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton begining 
      Caption         =   "Move to F&irst"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "OwnerOfBook"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   2250
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "Returned"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   11
      Left            =   3840
      TabIndex        =   4
      Top             =   1785
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "Received"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   3
      Top             =   1785
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "ISBN"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   9
      Left            =   2040
      TabIndex        =   2
      Top             =   945
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "Author"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   1
      Top             =   570
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "BookTitle"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   0
      Top             =   210
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "Telephone"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   11
      Top             =   3720
      Width           =   3495
   End
   Begin VB.CommandButton ExitRec 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton UpdateRec 
      Caption         =   "&Update"
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton NextRec 
      Caption         =   "&Next"
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton PreviousRec 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   1680
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton DeleteRec 
      Caption         =   "&Delete Record"
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton AddRec 
      Caption         =   "&Add Record"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Address1"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   2610
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "State"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   4
      Left            =   3840
      TabIndex        =   9
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox Text1 
      DataField       =   "Address2"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   7
      Top             =   2985
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      DataField       =   "Zip"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   5
      Left            =   4560
      TabIndex        =   10
      Top             =   3345
      Width           =   975
   End
   Begin VB.TextBox Text1 
      DataField       =   "City"
      DataSource      =   "RNBookLoan"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   8
      Top             =   3345
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Book Author:"
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   29
      Top             =   660
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Telephone:"
      Height          =   195
      Index           =   10
      Left            =   1080
      TabIndex        =   28
      Top             =   3810
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City/State/Zip:"
      Height          =   195
      Index           =   9
      Left            =   840
      TabIndex        =   27
      Top             =   3435
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address 2:"
      Height          =   195
      Index           =   8
      Left            =   1140
      TabIndex        =   26
      Top             =   3075
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address 1:"
      Height          =   195
      Index           =   7
      Left            =   1140
      TabIndex        =   25
      Top             =   2700
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Owner of Book:"
      Height          =   195
      Index           =   5
      Left            =   780
      TabIndex        =   24
      Top             =   2340
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   195
      Index           =   4
      Left            =   1845
      TabIndex        =   23
      Top             =   2250
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Received / Returned:"
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   22
      Top             =   1875
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   195
      Index           =   2
      Left            =   1470
      TabIndex        =   21
      Top             =   1035
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Title of Book:"
      Height          =   195
      Index           =   0
      Left            =   945
      TabIndex        =   20
      Top             =   300
      Width           =   945
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "BookLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RNBDB As Recordset

Private Sub AddRec_Click()
    RNBDB.AddNew
    Text1(0).SetFocus
End Sub
Private Sub begining_Click()
    RNBDB.MoveFirst
End Sub

Private Sub Command1_Click()
'Dim dbsCurrent As Database
Dim QueryDataBase As QueryDef
'Set QueryDataBase = dbsCurrent.CreateQueryDef(" ")
'QueryDataBase.Connect = "ODBC;DATABASE=pubs;UID=sa;PWD=;" & "DSN=bookloan"
'QueryDataBase.SQL = "SELECT title"
QueryDataBase.Execute







End Sub

Private Sub DeleteRec_Click()
If Not RNBDB.EOF Then
    RNBDB.Delete
    RNBDB.MovePrevious
End If
End Sub

Private Sub ending_Click()
    RNBDB.MoveLast
End Sub

Private Sub ExitRec_Click()
    End
End Sub

Private Sub NextRec_Click()
If Not RNBDB.EOF Then
    RNBDB.MoveNext
End If
End Sub

Private Sub PreviousRec_Click()
If Not RNBDB.BOF Then
    RNBDB.MovePrevious
End If
End Sub
Private Sub RNBookLoan_Reposition()
    Static RNDataLoad As Boolean
    If RNDataLoad = False Then
        Set RNBDB = RNBookLoan.Recordset
        RNDataLoad = True
    End If
    
    
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

