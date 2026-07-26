VERSION 5.00
Begin VB.Form frmMingo 
   Caption         =   "Mingo Sales"
   ClientHeight    =   3645
   ClientLeft      =   1665
   ClientTop       =   1545
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   5025
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear Screen"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Commission"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtFlorida 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00;(""$""#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtMaine 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00;(""$""#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtNewYork 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00;(""$""#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblComm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00;(""$""#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblTsales 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00;(""$""#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission:"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Width           =   870
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total sales:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Florida sales:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Maine sales:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&New York sales:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1155
   End
End
Attribute VB_Name = "frmMingo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
    Dim curNewyork As Integer, curMaine As Integer, curFlorida As Integer
    Dim curTsales As Integer, curComm As Currency
    
    curComm = InputBox("Please enter the commision rate.", "Commision rate")
    curFlorida = txtFlorida.Text
    curMaine = txtMaine.Text
    curNewyork = txtNewYork.Text
    
    curTsales = curFlorida + curMaine + curNewyork
'    curComm = curTsales * 0.05
    
    lblTsales = Format(curTsales, "currency")
    lblComm = Format(curTsales * curComm, "currency")
    
    cmdPrint.SetFocus

End Sub

Private Sub cmdClear_Click()
    txtNewYork.Text = "$0.00"
    txtMaine.Text = "$0.00"
    txtFlorida.Text = "$0.00"
    lblComm.Caption = "$0.00"
    lblTsales.Caption = "$0.00"
    txtNewYork.SetFocus
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdPrint_Click()
    cmdCalc.Visible = False
    cmdClear.Visible = False
    cmdExit.Visible = False
    cmdPrint.Visible = False
    PrintForm
    cmdClear.SetFocus
End Sub

Private Sub Form_Load()
Dim curFlorida As Currency, curMaine As Currency, curNewyork As Currency
    txtFlorida.Text = "$0.00"
    txtMaine.Text = "$0.00"
    txtNewYork.Text = "$0.00"
    txtMaine.SelLength = Len(txtMaine.Text)
    txtNewYork.SelLength = Len(txtNewYork.Text)
    txtFlorida.SelLength = Len(txtFlorida.Text)
End Sub

Private Sub txtFlorida_LostFocus()
    txtFlorida.Text = Format(txtFlorida.Text, "currency")

End Sub

Private Sub txtMaine_LostFocus()
    txtMaine.Text = Format(txtMaine.Text, "currency")

End Sub

Private Sub txtNewYork_LostFocus()
    txtNewYork.Text = Format(txtNewYork.Text, "currency")

End Sub
