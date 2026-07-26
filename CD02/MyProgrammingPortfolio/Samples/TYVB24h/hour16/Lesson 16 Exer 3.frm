VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Titles"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Print"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Titles"
      Connect         =   "Access"
      DatabaseName    =   "K:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Titles"
      Top             =   4200
      Width           =   2700
   End
   Begin VB.Label Label2 
      Caption         =   "(Modified) Set your Biblio.mdb's path name in the Data Control"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "ISBN"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Year Published"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Title"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label lblFields 
      Alignment       =   1  'Right Justify
      Caption         =   "ISBN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblFields 
      Alignment       =   1  'Right Justify
      Caption         =   "Year Published:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblFields 
      Alignment       =   1  'Right Justify
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1560
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
Private Sub Command1_Click()
  Unload Me
  End
End Sub

Private Sub Command2_Click()
  ' Warning: No error testing is done here
  ' for an existing online printer!
  ' Also, you may want to go to the
  ' top of the page the first time printed
  ' and about 50 titles later if the user
  ' prints that many.
  Printer.Print "Title: "; Label1(0)
  Printer.Print "Year Published: "; Label1(1)
  Printer.EndDoc
End Sub
