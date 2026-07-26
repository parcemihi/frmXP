VERSION 5.00
Begin VB.Form frmIMY 
   Caption         =   "IMY Industries"
   ClientHeight    =   4485
   ClientLeft      =   1110
   ClientTop       =   1620
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   6165
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&New Hourly Pay"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtCurrent3 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtCurrent2 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtCurrent1 
      Alignment       =   1  'Right Justify
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
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   14
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Label lblNew3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2880
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblNew2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2880
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblNew1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2880
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Job code &3:"
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Job code &2:"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Job code &1:"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Hourly Pay"
      Height          =   195
      Left            =   2880
      TabIndex        =   10
      Top             =   360
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Hourly Pay"
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   360
      Width           =   1320
   End
End
Attribute VB_Name = "frmIMY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
Dim curJC1 As Currency, curJC2 As Currency, curJC3 As Currency
Dim raise1 As Currency, raise2 As Currency, raise3 As Currency
Dim RaisePercent As String, msgRaise As String
    msgRaise = "Raise percentage: "
    RaisePercent = "Raise percentage: "
    
    curJC1 = txtCurrent1.Text
    curJC2 = txtCurrent2.Text
    curJC3 = txtCurrent3.Text

    raise1 = Val(InputBox("Raise percent for Job Code 1.", RaisePercent))
    raise2 = Val(InputBox("Raise percent for Job Code 2.", RaisePercent))
    raise3 = Val(InputBox("Raise percent for Job Code 3.", RaisePercent))
    
    lblNew1.Caption = format(curJC1 * (raise1 + 1), "currency")
    lblNew2.Caption = format(curJC2 * (raise2 + 1), "currency")
    lblNew3.Caption = format(curJC3 * (raise3 + 1), "currency")


lblMsg.Caption = _
    msgRaise & format(raise1, "percent") & vbNewLine & _
    msgRaise & format(raise2, "percent") & vbNewLine & _
    msgRaise & format(raise3, "percent")







End Sub

Private Sub cmdExit_Click()
    End
End Sub
 
Private Sub cmdPrint_Click()
    PrintForm
    txtCurrent1.SetFocus
End Sub

Private Sub Form_Load()
    txtCurrent1 = 0
    txtCurrent2 = 0
    txtCurrent3 = 0
    txtCurrent1.Text = format(txtCurrent1.Text, "currency")
    txtCurrent2.Text = format(txtCurrent2.Text, "currency")
    txtCurrent3.Text = format(txtCurrent3.Text, "currency")
End Sub

Private Sub txtCurrent1_GotFocus()
    txtCurrent1.SelLength = Len(txtCurrent1)
End Sub

Private Sub txtCurrent1_LostFocus()
    txtCurrent1.Text = format(txtCurrent1.Text, "currency")
End Sub

Private Sub txtCurrent2_GotFocus()
    txtCurrent2.SelLength = Len(txtCurrent2)
End Sub

Private Sub txtCurrent2_LostFocus()
    txtCurrent2.Text = format(txtCurrent2.Text, "currency")
End Sub

Private Sub txtCurrent3_GotFocus()
    txtCurrent3.SelLength = Len(txtCurrent3)
End Sub

Private Sub txtCurrent3_LostFocus()
    txtCurrent3.Text = format(txtCurrent3.Text, "currency")
End Sub
