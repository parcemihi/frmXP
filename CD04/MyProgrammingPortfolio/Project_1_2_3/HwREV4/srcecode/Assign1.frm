VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Assignment 1 and 2"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Delete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   1770
      TabIndex        =   12
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Update 
      Caption         =   "&Update"
      Height          =   495
      Left            =   4350
      TabIndex        =   14
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton NewOrder 
      Caption         =   "&New Order"
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Purchase 
      Caption         =   "&Purchase"
      Height          =   495
      Left            =   3060
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   2580
      TabIndex        =   0
      Text            =   "Customer Order No."
      Top             =   563
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   2580
      TabIndex        =   10
      Top             =   3083
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   5
      Top             =   2003
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   2580
      TabIndex        =   4
      Top             =   2003
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   2580
      TabIndex        =   7
      Top             =   2483
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   9
      Top             =   2483
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   8
      Top             =   2483
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   3
      Top             =   1523
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   1523
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2580
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5460
      TabIndex        =   56
      Top             =   4455
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   4140
      TabIndex        =   55
      Top             =   5655
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Expert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4680
      TabIndex        =   54
      Top             =   4080
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Intermediate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3360
      TabIndex        =   53
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Beginner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   52
      Top             =   4043
      Width           =   810
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   5460
      TabIndex        =   51
      Top             =   5655
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   2820
      TabIndex        =   50
      Top             =   5655
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5460
      TabIndex        =   49
      Top             =   5055
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   4140
      TabIndex        =   48
      Top             =   5055
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2820
      TabIndex        =   47
      Top             =   5055
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4140
      TabIndex        =   46
      Top             =   4455
      Width           =   90
   End
   Begin VB.Label price 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2820
      TabIndex        =   45
      Top             =   4455
      Width           =   90
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   6
      Left            =   2025
      TabIndex        =   44
      Top             =   5655
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   3
      Left            =   2025
      TabIndex        =   43
      Top             =   5055
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   0
      Left            =   2025
      TabIndex        =   42
      Top             =   4455
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   8
      Left            =   4665
      TabIndex        =   41
      Top             =   5655
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   7
      Left            =   3345
      TabIndex        =   40
      Top             =   5655
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   5
      Left            =   4665
      TabIndex        =   39
      Top             =   5055
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   4
      Left            =   3345
      TabIndex        =   38
      Top             =   5055
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   2
      Left            =   4665
      TabIndex        =   37
      Top             =   4455
      Width           =   195
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   " At"
      Height          =   195
      Index           =   1
      Left            =   3345
      TabIndex        =   36
      Top             =   4455
      Width           =   195
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   1
      Left            =   6060
      TabIndex        =   35
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Packages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   34
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   5700
      TabIndex        =   33
      Top             =   3690
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Shipping/Handling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   32
      Top             =   3083
      Width           =   1665
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Order Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   31
      Top             =   6810
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Delivery Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   30
      Top             =   6210
      Width           =   1260
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   4
      Left            =   4980
      TabIndex        =   29
      Top             =   6855
      Width           =   1800
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   3
      Left            =   6060
      TabIndex        =   28
      Top             =   6240
      Width           =   720
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   2
      Left            =   6060
      TabIndex        =   27
      Top             =   5655
      Width           =   720
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   0
      Left            =   6060
      TabIndex        =   26
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   540
      TabIndex        =   25
      Top             =   3683
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Package C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   24
      Top             =   5610
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Package B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   23
      Top             =   5010
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Package A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   22
      Top             =   4410
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Software Package C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   21
      Top             =   2483
      Width           =   1830
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Expert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5460
      TabIndex        =   20
      Top             =   1163
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Intermediate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4020
      TabIndex        =   19
      Top             =   1163
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Beginner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2580
      TabIndex        =   18
      Top             =   1163
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Software Package B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   17
      Top             =   2040
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Software Package A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   16
      Top             =   1523
      Width           =   1830
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Database"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Data"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim luppe As Integer
Dim nachricht As Integer
Dim pricelist(0 To 8) As Currency

Dim Customer(100) As CDF
Private Type CDF
    ID As String
    purchases(0 To 8) As Integer
    shipcost As Single
End Type

Private Sub Form_Load()
Dim indicies As Integer
indicies = 0
Open App.Path & "\Custom1.dat" For Input As #1
    Do Until EOF(1)
        Input #1, Customer(indicies).ID, Customer(indicies).purchases(0), Customer(indicies).purchases(1), Customer(indicies).purchases(2), Customer(indicies).purchases(3), Customer(indicies).purchases(4), Customer(indicies).purchases(5), Customer(indicies).purchases(6), Customer(indicies).purchases(7), Customer(indicies).purchases(8), Customer(indicies).shipcost
        Combo1.AddItem Customer(indicies).ID
        indicies = indicies + 1
    Loop
Close #1
Dim n As Integer
Open App.Path & "\pricelist.dat" For Input As #1
    For n = 0 To 8
        Input #1, pricelist(n)
        price(n).Caption = Format(pricelist(n), "currency")
    Next n
Close #1
End Sub

Private Sub Combo1_Click()
Dim x As Variant
For luppe = 0 To 8
    Text1(luppe).Text = Customer(Combo1.ListIndex).purchases(luppe)
Next luppe
Text1(9).Text = Format(Customer(Combo1.ListIndex).shipcost, "Currency")
SubTotal
End Sub

Public Function SubTotal()
Dim x As Variant
Dim n As Integer
For n = 0 To 8
    QtySft(n).Caption = " At"
Next n
For n = 0 To 8
    QtySft(n).Caption = Customer(Combo1.ListIndex).purchases(n) & QtySft(n).Caption
Next n






End Function

Private Sub Exit_Click()
    End
'if
'    nachricht = MsgBox("Would you like to update your datafile before exiting?", vbYesNoCancel)
'end if
'Select Case nachricht
'    Case 6 'yes
'    Case 7 'no
'    End
'    Case 2 'cancel
'End Select
End Sub

'Option Explicit
'Dim i As Integer, j As Integer, k As Integer, l As Integer
'Dim P(0 To 2) As Single
'Dim x As Variant, y As Variant, z As Variant
'Dim onLoad As Integer
'Dim CustData As String
'Dim PriceLst As String
'Private Sub Form_Load()
'    onLoad = 0
'    LoadCustomer_datafiles
'End Sub
'Public Sub LoadCustomer_datafiles()
'    onLoad = 0
'Preface:
'On Error GoTo OnloadError
'Select Case onLoad
'Case 0 'Load file names. On error create file
'    FileName = "\appth.dat"
'    Open App.Path & FileName For Input As #1
'    Input #1, CustData
'    FileName = "\" & CustData
'    Close #1
'    onLoad = 1
'    GoTo Preface
'Case 1 'Load customer order information
'    Fins (App.Path & FileName)
'    onLoad = 2
'    GoTo Preface
'Case 2 'Populate Combo1 with customer order numbers
'    For i = 0 To BN
'    Combo1.AddItem kdc.Index_(i)
'    Next i
'    onLoad = 3
'    GoTo Preface
'Case 3 'Load Pricelist
'    i = 0
'    FileName = "\pricelist.dat"
'    Open App.Path & FileName For Input As #1
'    Do Until EOF(1)
'    Input #1, kdc.price(i)
'    i = i + 1
'    Loop
'    Close #1
'    For i = 0 To 8
'    price(i).Caption = kdc.price(i)
'    Next i
'    Exit Sub
'End Select
'OnloadError:
'errNumber.ErrorHandler (Err.Number)
'Resume
'End Sub
'Private Sub SubTotal()
'    P(0) = 0
'    P(1) = 0
'    P(2) = 0
'    For i = 0 To 8
'        QtySft(i).Caption = kdc.bank(i, Combo1.ListIndex)
'        P(Int(i / 3)) = P(Int(i / 3)) + kdc.price(i) * kdc.bank(i, Combo1.ListIndex)
'        tot(Int(i / 3)).Caption = Format(P(Int(i / 3)), "Currency")
'    Next i
'    i = 0
'    tot(4).Caption = Format(P(0) + P(1) + P(2), "Currency")
'End Sub
'
''// COMMAND INTERFACE
'Private Function CheckChar()
'preface3:
'x = UCase(InputBox("Enter customer ID", "Add new customer"))
'For i = 1 To Len(x)
'    For j = 48 To 57
'        If Mid(x, i, 1) = Chr(j) Then
'        Exit For
'        End If
'    Next j
'    For k = 65 To 90
'        If Mid(x, i, 1) = Chr(k) Then
'        Exit For
'        End If
'    Next k
'    If j = 58 And k = 91 Then
'        GoTo Error
'    End If
'Next i
'Exit Function
'Error:
'y = MsgBox("Invalid ID." & Chr(13) & "Numbers and letters only.", vbCritical, "Invalid ID.")
'GoTo preface3
'End Function
'Private Sub NewOrder_Click()
'Preface2:
'CheckChar
'If Not x = "" Then
'    For i = 0 To 2500
'        If kdc.Index_(i) = x Then
'        x = MsgBox("The customer ID you have chosen is already taken." & Chr(13) & "Please enter a different customer ID.", vbOKOnly, "Customer ID")
'        GoTo Preface2
'        ElseIf kdc.Index_(i) = "" Then
'        Exit For
'        End If
'    Next i
'BN = BN + 1
'    x = x
'    If i = 0 Then
'    kdc.Index_(i) = x
'    kdc.Fields(i) = 10
'        For j = 0 To 9
'        kdc.bank(j, i) = 0
'        Next j
'    Combo1.AddItem kdc.Index_(i), i
'    Combo1.ListIndex = i
'    End If
'    Do Until i = 0
'        If x < kdc.Index_(i - 1) Then
'        kdc.Index_(i) = kdc.Index_(i - 1)
'        kdc.Fields(i) = kdc.Fields(i - 1)
'            For j = 0 To 9
'            kdc.bank(j, i) = kdc.bank(j, i - 1)
'            Next j
'        i = i - 1
'        Else
'        kdc.Index_(i) = x
'        kdc.Fields(i) = 10
'            For j = 0 To 9
'            kdc.bank(j, i) = 0
'            Next j
'        Combo1.AddItem kdc.Index_(i), i
'        Combo1.ListIndex = i
'        Exit Do
'        End If
'    If i = 0 Then
'    kdc.Index_(i) = x
'    kdc.Fields(i) = 10
'        For j = 0 To 9
'        kdc.bank(j, i) = 0
'        Next j
'    Combo1.AddItem kdc.Index_(i), i
'    Combo1.ListIndex = i
'    End If
'    Loop
'End If
'End Sub
'Private Sub Delete_Click()
'If Not Combo1.ListIndex = -1 Then
'For j = 0 To 2500
'    kdc.Fields(j + Combo1.ListIndex) = kdc.Fields(j + Combo1.ListIndex + 1)
'    kdc.Index_(j + Combo1.ListIndex) = kdc.Index_(j + Combo1.ListIndex + 1)
'    For i = 0 To 9
'    kdc.bank(i, j + Combo1.ListIndex) = kdc.bank(i, j + Combo1.ListIndex + 1)
'    Next i
'    If kdc.Index_(j) = "" Then
'    Exit For
'    End If
'Next j
'x = Combo1.ListIndex
'Combo1.RemoveItem (Combo1.ListIndex)
'If x < Combo1.ListCount Then
'Combo1.ListIndex = x
'ElseIf Combo1.ListCount = 0 Then
'Else
'Combo1.ListIndex = Combo1.ListCount - 1
'End If
'End If
'Combo1.SetFocus
'End Sub
'Private Sub Purchase_Click()
'Fpurch (Combo1.Text)
'End Sub
'Private Sub Update_Click()
'    Fouts (CustData)
'End Sub
'Private Sub Exit_Click()
'    End
'End Sub
'
''// PROGRAM BEHAVIOR
'Private Sub Combo1_Click()
'For i = 0 To 9
'    Text1(i).Text = kdc.bank(i, Combo1.ListIndex)
'Next i
'i = 0
'SubTotal
'End Sub
'Private Sub Text1_GotFocus(Index As Integer)
'    Text1(Index).SelStart = 0
'    Text1(Index).SelLength = Len(Text1(Index).Text)
'End Sub
'Private Sub Text1_LostFocus(Index As Integer)
'    On Error GoTo ErrorHandler
'    kdc.bank(Index, Combo1.ListIndex) = Val(Text1(Index).Text)
'    SubTotal
'    Exit Sub
'ErrorHandler:
'    errNumber.ErrorHandler (Err.Number)
'End Sub
