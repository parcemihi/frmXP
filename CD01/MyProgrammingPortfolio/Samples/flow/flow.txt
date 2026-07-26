VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton discount 
      Caption         =   "Price - 40 %"
      Height          =   495
      Index           =   3
      Left            =   3000
      TabIndex        =   4
      Top             =   6975
      Width           =   1215
   End
   Begin VB.CommandButton discount 
      Caption         =   "Price - 30 %"
      Height          =   495
      Index           =   2
      Left            =   3000
      TabIndex        =   3
      Top             =   5415
      Width           =   1215
   End
   Begin VB.CommandButton discount 
      Caption         =   "Price - 20 %"
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   2
      Top             =   3855
      Width           =   1215
   End
   Begin VB.CommandButton discount 
      Caption         =   "Price - 10 %"
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   2295
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   1095
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   195
      Index           =   16
      Left            =   1560
      TabIndex        =   25
      Top             =   6375
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   195
      Index           =   15
      Left            =   1560
      TabIndex        =   24
      Top             =   4815
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "NO"
      Height          =   195
      Index           =   14
      Left            =   1560
      TabIndex        =   23
      Top             =   3255
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "YES"
      Height          =   195
      Index           =   13
      Left            =   2520
      TabIndex        =   22
      Top             =   5415
      Width           =   345
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "YES"
      Height          =   195
      Index           =   12
      Left            =   2520
      TabIndex        =   21
      Top             =   3855
      Width           =   345
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "YES"
      Height          =   195
      Index           =   11
      Left            =   2520
      TabIndex        =   20
      Top             =   2295
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   4920
      TabIndex        =   9
      Top             =   7095
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   4920
      TabIndex        =   8
      Top             =   5535
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   4920
      TabIndex        =   7
      Top             =   3975
      Width           =   825
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Start"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   255
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   44
      X1              =   1920
      X2              =   1920
      Y1              =   975
      Y2              =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Left            =   4920
      TabIndex        =   5
      Top             =   2415
      Width           =   825
   End
   Begin VB.Line Line1 
      Index           =   43
      X1              =   4800
      X2              =   6000
      Y1              =   6975
      Y2              =   6975
   End
   Begin VB.Line Line1 
      Index           =   42
      X1              =   4560
      X2              =   4800
      Y1              =   7455
      Y2              =   6975
   End
   Begin VB.Line Line1 
      Index           =   41
      X1              =   5760
      X2              =   6000
      Y1              =   7455
      Y2              =   6975
   End
   Begin VB.Line Line1 
      Index           =   40
      X1              =   4560
      X2              =   5760
      Y1              =   7455
      Y2              =   7455
   End
   Begin VB.Line Line1 
      Index           =   39
      X1              =   4800
      X2              =   6000
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Line Line1 
      Index           =   38
      X1              =   4560
      X2              =   4800
      Y1              =   5895
      Y2              =   5415
   End
   Begin VB.Line Line1 
      Index           =   37
      X1              =   5760
      X2              =   6000
      Y1              =   5895
      Y2              =   5415
   End
   Begin VB.Line Line1 
      Index           =   36
      X1              =   4560
      X2              =   5760
      Y1              =   5895
      Y2              =   5895
   End
   Begin VB.Line Line1 
      Index           =   35
      X1              =   4800
      X2              =   6000
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line1 
      Index           =   34
      X1              =   4560
      X2              =   4800
      Y1              =   4335
      Y2              =   3855
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   5760
      X2              =   6000
      Y1              =   4335
      Y2              =   3855
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   4560
      X2              =   5760
      Y1              =   4335
      Y2              =   4335
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   4800
      X2              =   6000
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   4560
      X2              =   4800
      Y1              =   2775
      Y2              =   2295
   End
   Begin VB.Line Line1 
      Index           =   29
      X1              =   5760
      X2              =   6000
      Y1              =   2775
      Y2              =   2295
   End
   Begin VB.Line Line1 
      Index           =   28
      X1              =   4560
      X2              =   5760
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   4200
      X2              =   4680
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   4680
      X2              =   4200
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   4200
      X2              =   4680
      Y1              =   5655
      Y2              =   5655
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   4200
      X2              =   4680
      Y1              =   7215
      Y2              =   7215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   135
      Width           =   975
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   1920
      X2              =   3000
      Y1              =   7215
      Y2              =   7215
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   1920
      X2              =   1920
      Y1              =   3495
      Y2              =   3135
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   1200
      X2              =   2400
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   1920
      X2              =   1920
      Y1              =   4695
      Y2              =   5055
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   1920
      X2              =   1920
      Y1              =   1935
      Y2              =   1455
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   2400
      X2              =   2640
      Y1              =   1455
      Y2              =   975
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   1200
      X2              =   1440
      Y1              =   1455
      Y2              =   975
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   1440
      X2              =   2640
      Y1              =   975
      Y2              =   975
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   2520
      X2              =   3000
      Y1              =   5655
      Y2              =   5655
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   3000
      X2              =   2520
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   2520
      X2              =   3000
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   1920
      X2              =   1920
      Y1              =   6255
      Y2              =   7215
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   1920
      X2              =   2520
      Y1              =   6255
      Y2              =   5655
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   1920
      X2              =   1320
      Y1              =   6255
      Y2              =   5655
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1320
      X2              =   1920
      Y1              =   5655
      Y2              =   5055
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   1920
      X2              =   2520
      Y1              =   5055
      Y2              =   5655
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   1920
      X2              =   2520
      Y1              =   4695
      Y2              =   4095
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1920
      X2              =   1320
      Y1              =   4695
      Y2              =   4095
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1320
      X2              =   1920
      Y1              =   4095
      Y2              =   3495
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   1920
      X2              =   2520
      Y1              =   3495
      Y2              =   4095
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   1920
      X2              =   2520
      Y1              =   3135
      Y2              =   2535
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1920
      X2              =   1320
      Y1              =   3135
      Y2              =   2535
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1320
      X2              =   1920
      Y1              =   2535
      Y2              =   1935
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1920
      X2              =   2520
      Y1              =   1935
      Y2              =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "10%"
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   13
      Top             =   2655
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "20%"
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   16
      Top             =   4215
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "30%"
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   19
      Top             =   5775
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Discount"
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   18
      Top             =   5535
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Is"
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   17
      Top             =   5295
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Discount"
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   15
      Top             =   3975
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Is"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   14
      Top             =   3735
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Discount"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   12
      Top             =   2415
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Is"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Top             =   2175
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Input Price"
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   10
      Top             =   1095
      Width           =   795
   End
   Begin VB.Menu Exit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu ViewCode 
      Caption         =   "&View Code"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const HelpCNT = &HB

Private Sub discount_Click(Index As Integer)
    Dim price As String
    price = Val(Text1.Text) * (1 - ((Index + 1) / 10))
    Label1(Index).Caption = Format(price, "$##,##0.00")
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub Form_Load()
    App.HelpFile = App.Path & "\flow.hlp"
End Sub

Private Sub ViewCode_Click()
   With CommonDialog1
      .HelpFile = App.Path & "\flow.hlp"
      .HelpCommand = HelpCNT Or cdlHelpSetContents
      .ShowHelp
   End With
End Sub
