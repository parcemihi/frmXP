VERSION 5.00
Begin VB.Form frmSales 
   Caption         =   "Examples of local and form-level variables"
   ClientHeight    =   3615
   ClientLeft      =   1335
   ClientTop       =   1515
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3615
   ScaleWidth      =   5595
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display Sales"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmd2Comm 
      Caption         =   "&2% Commission"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmd5Comm 
      Caption         =   "&5% Commission"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Commission:"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   570
   End
   Begin VB.Label lblSales 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblCommission 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Commission Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curSales As Currency

Private Sub cmd2Comm_Click()
    lblCommission.Caption = curSales * 0.02
    
End Sub

Private Sub cmd5Comm_Click()
    Dim curSales As Currency
    lblCommission.Caption = curSales * 0.05
    
End Sub

Private Sub cmdDisplay_Click()
    curSales = 500
    lblSales.Caption = curSales
    
End Sub

Private Sub cmdExit_Click()
    End
    
End Sub


