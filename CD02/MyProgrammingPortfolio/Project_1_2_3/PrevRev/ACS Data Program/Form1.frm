VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\My Documents\ACS\acsMain.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employees"
      Top             =   1920
      Width           =   2940
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "SocialSecurityNumber"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Text            =   "Combo3"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "FirstName"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Text            =   "Combo2"
      Top             =   840
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "LastName"
      DataSource      =   "Data1"
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   480
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
