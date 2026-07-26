VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form WinePastaTree 
   Caption         =   "The Wine Bar at The Pasta Tree"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "PastaTreeWine.frx":0000
      Height          =   2775
      Left            =   120
      TabIndex        =   12
      Top             =   5280
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4895
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   6840
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Text Files"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Text Files"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Text;"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\Project_1_2_3\Version 3"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "sales"
      Top             =   4680
      Width           =   3660
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PastaTreeWine.frx":0015
      Height          =   2295
      Left            =   5160
      TabIndex        =   11
      Top             =   5880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      RowDividerStyle =   1
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Customer Order"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "Menu Item"
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
         Caption         =   "Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         BeginProperty Column00 
            DividerStyle    =   0
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            DividerStyle    =   0
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   3630
      ItemData        =   "PastaTreeWine.frx":002A
      Left            =   9120
      List            =   "PastaTreeWine.frx":002C
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   8295
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "Wine Family"
      Top             =   600
      Width           =   4695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "PastaTreeWine.frx":002E
      Left            =   120
      List            =   "PastaTreeWine.frx":0030
      TabIndex        =   7
      Text            =   "Customer Order Number"
      Top             =   240
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Total"
      Height          =   495
      Index           =   4
      Left            =   6085
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   3
      Left            =   4565
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Index           =   2
      Left            =   3045
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Order"
      Height          =   495
      Index           =   1
      Left            =   1525
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Customer"
      Height          =   495
      Index           =   0
      Left            =   5
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      ItemData        =   "PastaTreeWine.frx":0032
      Left            =   4920
      List            =   "PastaTreeWine.frx":0034
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      ItemData        =   "PastaTreeWine.frx":0036
      Left            =   120
      List            =   "PastaTreeWine.frx":0038
      TabIndex        =   0
      Top             =   1020
      Width           =   4695
   End
End
Attribute VB_Name = "WinePastaTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents rsNames As Adodc.Recordset
Attribute rsNames.VB_VarHelpID = -1
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim x As Variant, y As Variant, z As Variant

Dim WIndex As WineIndex
Private Type WineIndex
    IndexCat(100, 100) As Integer
    WineName As String
End Type


Dim WList As WineList
Private Type WineList
    Index As Integer
    Cat As Integer
    Name As String
    Price As Integer
    
End Type

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub Form_Load()
j = 0
Open App.Path & "\WineIndex.dat" For Input As #1
Do Until EOF(1)
    Input #1, x, y, WIndex.WineName
    Combo2.AddItem WIndex.WineName
    For i = 0 To y - 1
        WIndex.IndexCat(i, x) = j
        j = j + 1
    Next i
Loop
Close #1
Combo2.Text = "Wine Family"


End Sub

Private Sub Combo2_Click()
i = 0
List1.Clear
Open App.Path & "\WineList.dat" For Input As #1
Do Until EOF(1)
    Input #1, WList.Index, WList.Cat, WList.Name, WList.Price
    
    
    If Combo2.ListIndex = WList.Cat Then
'        If List1.Index = 0 Then
        
        
        
        
'        End If
    List1.AddItem WList.Name
    End If
Loop
Close #1















End Sub


Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
x = Combo1.ListCount + 1
Combo1.AddItem x
    Open App.Path & "\" & x & ".dat" For Output As #1
    Write #1, "Order#", x
    Close #1
Case 1
Open App.Path & "\" & Combo1.ListIndex & ".dat" For Append As #1
Close #1
Case 2
DataGrid1.AddNewMode













End Select
End Sub


Private Sub List1_DblClick()
If Combo1.Text = "Customer Order Number" Then
x = MsgBox("You must select a customer order number", vbCritical, "Invalid Order Number")
Else




j = 0
Open App.Path & "\WineList.dat" For Input As #1
y = 0
x = ""
Do Until EOF(1)
For i = 0 To Combo2.ListIndex
    
    Input #1, WList.Index, WList.Cat, WList.Name, WList.Price
    If WList.Index = WIndex.IndexCat(List1.ListIndex, Combo2.ListIndex) Then
        
    List2.AddItem WList.Name
        For j = 0 To 5 - Len(Format(WList.Price, "#,##0.00"))
            x = x & " "
            'x = Len(Format(WList.Price, "#,##0.00"))
        Next j
        x = x & Format(WList.Price, "#,##0.00")
    List3.AddItem x
'    j = WIndex.IndexCat(List1.ListIndex, Combo2.ListIndex)
    Open App.Path & "\sales.dat" For Append As #2
    Write #2, WList.Index, WList.Cat, WList.Name, WList.Price
    Close #2
    Exit Do
    End If
'    y = y + x
    
Next i
    
Loop
Close #1
    
'    If Combo2.ListIndex = WList.Cat Then
'    List1.AddItem WList.Name
'    End If


'List2.AddItem List1.List(List1.ListIndex)







End If
End Sub
