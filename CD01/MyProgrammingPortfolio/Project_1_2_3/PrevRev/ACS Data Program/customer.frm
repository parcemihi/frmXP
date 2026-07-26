VERSION 5.00
Begin VB.Form customer 
   AutoRedraw      =   -1  'True
   Caption         =   "Customer data file"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "EMP"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\Nwind.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employees"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "PostalCode"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   4080
      TabIndex        =   21
      Text            =   "Text3"
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      DataField       =   "City"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2160
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "Region"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   3480
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   3120
      Width           =   495
   End
   Begin VB.ComboBox Combo4 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2160
      TabIndex        =   18
      Text            =   "Combo4"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Address"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Text            =   "Combo3"
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2160
      TabIndex        =   16
      Text            =   "Combo2"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "LastName"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2160
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   7080
      TabIndex        =   14
      Top             =   0
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3960
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Customer"
      Default         =   -1  'True
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Employee"
      Height          =   975
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      Height          =   975
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   975
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      Height          =   975
      Left            =   3960
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reports"
      Height          =   975
      Left            =   4920
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Calander"
      Height          =   975
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Notes:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "City / State / Zip"
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
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Street Address"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Street Address"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "First Name"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Last Name"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Menu file 
      Caption         =   "&File"
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
customer.Hide
employee.Show
End Sub

Private Sub Form_Load()
'WindowState = 2


End Sub

