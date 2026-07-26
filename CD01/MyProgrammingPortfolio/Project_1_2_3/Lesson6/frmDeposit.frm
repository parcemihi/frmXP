VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeposit 
   Caption         =   "Deposit Check"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTransactionDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton cmdToday 
      Caption         =   "Get Today"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtCredit 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   1560
      TabIndex        =   7
      Top             =   2040
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24510465
      CurrentDate     =   36621
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim objTransaction As New clsTransaction
    
    objTransaction.AddDeposit
    
End Sub

Private Sub cmdToday_Click()
    txtTransactionDate = Format(Now(), "mm/dd/yyyy")
End Sub

Private Sub Form_Load()
    txtCredit = "0"
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    txtTransactionDate = Format(DateClicked, "mm/dd/yyyy")
End Sub


Private Sub txtCredit_GotFocus()
    
    With txtCredit
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub

