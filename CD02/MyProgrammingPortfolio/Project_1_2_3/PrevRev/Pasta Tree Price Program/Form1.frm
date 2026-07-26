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
   Begin VB.TextBox Text2 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Width           =   4380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim vntTemp As Variant
   Dim vnDummy As Variant

Data1.DatabaseName = App.Path & "\menurest.mdb"
Data1.RecordSource = "SELECT DISTINCT EmployeeID FROM employees"
Data1.Refresh

Do While Not Data1.Recordset.EOF
    vntTemp = Data1.Recordset!EmployeeID
    If IsNull(vntTemp) Then vntTemp = ""
    Combo1.AddItem CStr(vntTemp)
    Data1.Recordset.MoveNext
Loop

Data1.RecordSource = "Employees"
Data1.RecordsetType = vbRSTypeTable

Text1.DataField = "FirstName"
Text2.DataField = "LastName"


End Sub
