VERSION 5.00
Begin VB.Form frmInterest 
   Caption         =   "Interest Calculation"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute Interest"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Click to compute final investment"
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox txtEnding 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Compounded Investment"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox txtInvest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      ToolTipText     =   "Money you invested"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtTerm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      ToolTipText     =   "Number of periods"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtRate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "Annual rate investment grows"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblEnding 
      Alignment       =   1  'Right Justify
      Caption         =   "Ending Investment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Label lblInvest 
      Alignment       =   1  'Right Justify
      Caption         =   "I&nvestment Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label lblTerm 
      Alignment       =   1  'Right Justify
      Caption         =   "&Term (annual periods):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      Caption         =   "Interest Rate (8 for 8%):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Interest Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2090
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmInterest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
' Use a For loop to calculate a final total
' investment using compound interest.
'
' intNum is a loop control variable
' sngIRate is the annual interest rate
' intTerm is the Number of years in the investment
' curInitInv is the investor's initial investment
' sngInterest is the total interest paid
   Dim sngIRate As Single, sngInterest As Single
   Dim intTerm As Integer, intNum As Integer
   Dim curInitInv As Currency
   
   ' Error-checking
   If ErrorCheck() = 1 Then
     Exit Sub
   End If

   sngIRate = txtRate.Text / 100#
   intTerm = txtTerm.Text
   
   curInitInv = txtInvest.Text
   sngInterest = 1#   ' Begin at one for first compound

   ' Use loop to calculate total compound amount
   For intNum = 1 To intTerm
      sngInterest = sngInterest * (1 + sngIRate)
   Next intNum

   ' Now we have total interest,
   ' calculate the total investment
   ' at the end of N years
   txtEnding.Text = Format(curInitInv * sngInterest, "$###,##0.00")

End Sub
Private Sub cmdExit_Click()
   ' Unload the form and terminate application
   Unload frmInterest
   End
End Sub
