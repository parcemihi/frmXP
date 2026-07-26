VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheck 
   Caption         =   "Check Data Entry"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdToday 
      Caption         =   "Get Today"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
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
   Begin VB.TextBox txtTransactionDate 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.HScrollBar HScrollAmount 
      Height          =   255
      Left            =   4320
      Max             =   2000
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtDescription 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtCheckNumber 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Check#"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmCheck"
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
    
    objTransaction.AddCheck
    
End Sub

Private Sub cmdToday_Click()
    txtTransactionDate = Format(Now(), "mm/dd/yyyy")
End Sub

Private Sub Form_Load()
    txtAmount = "0"
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    txtTransactionDate = Format(DateClicked, "mm/dd/yyyy")
End Sub

Private Sub txtAmount_Change()
    If Val(txtAmount) < HScrollAmount.Max Then
    HScrollAmount.Value = Val(txtAmount)
    Else
    HScrollAmount.Value = HScrollAmount.Max
    End If
End Sub

Private Sub txtAmount_GotFocus()
    
    With txtAmount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub
