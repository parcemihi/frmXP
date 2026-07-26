VERSION 5.00
Begin VB.Form frmData 
   Caption         =   "Database Sample"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtZip 
      BackColor       =   &H80000018&
      DataField       =   "Zip"
      DataSource      =   "dtaBiblio"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   3000
      Width           =   1055
   End
   Begin VB.TextBox txtState 
      DataField       =   "State"
      DataSource      =   "dtaBiblio"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3000
      Width           =   555
   End
   Begin VB.TextBox txtCity 
      BackColor       =   &H80000018&
      DataField       =   "City"
      DataSource      =   "dtaBiblio"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H80000018&
      DataField       =   "Address"
      DataSource      =   "dtaBiblio"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2400
      Width           =   4215
   End
   Begin VB.TextBox txtPublisher 
      BackColor       =   &H80000018&
      DataField       =   "Company Name"
      DataSource      =   "dtaBiblio"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Data dtaBiblio 
      Caption         =   "Click to control data display"
      Connect         =   "Access"
      DatabaseName    =   "K:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Publishers"
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7080
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lblZip 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblState 
      Alignment       =   1  'Right Justify
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblCity 
      Alignment       =   1  'Right Justify
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      Caption         =   "Publisher:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblAnnounce 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer Book Publishers"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
