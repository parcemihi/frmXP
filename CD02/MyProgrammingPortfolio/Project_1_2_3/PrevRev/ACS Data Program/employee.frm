VERSION 5.00
Begin VB.Form employee 
   AutoRedraw      =   -1  'True
   Caption         =   "Employee data file"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text12 
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
      Height          =   285
      Left            =   7920
      TabIndex        =   29
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9000
      MaxLength       =   4
      TabIndex        =   28
      Text            =   "0000"
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      Text            =   "Current"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   8520
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "00"
      Top             =   1680
      Width           =   375
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7920
      MaxLength       =   3
      TabIndex        =   8
      Text            =   "000"
      Top             =   1680
      Width           =   495
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
      Height          =   2325
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4080
      Width           =   5775
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4800
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4200
      TabIndex        =   5
      Text            =   "IL"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   3
      Top             =   2640
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Customer"
      Default         =   -1  'True
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New"
      Height          =   975
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Edit"
      Height          =   975
      Left            =   2040
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Update"
      Height          =   975
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Reports"
      Height          =   975
      Left            =   3960
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Calander"
      Height          =   975
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "9999"
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
      Left            =   8760
      TabIndex        =   27
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Employee ID #"
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
      Left            =   7080
      TabIndex        =   26
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Payroll Status"
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
      Left            =   6240
      TabIndex        =   25
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Pay Rate"
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
      Left            =   6240
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Social Security"
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
      Left            =   6240
      TabIndex        =   23
      Top             =   1680
      Width           =   1560
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
      TabIndex        =   22
      Top             =   3720
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
Attribute VB_Name = "employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
customer.Show
employee.Hide
End Sub

Private Sub Form_Load()
'WindowState = 2


End Sub

