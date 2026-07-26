VERSION 5.00
Begin VB.Form frmWelcome 
   Caption         =   "Welcome Screen"
   ClientHeight    =   4860
   ClientLeft      =   1320
   ClientTop       =   1530
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4860
   ScaleWidth      =   5595
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Image imgLogo 
      Height          =   1575
      Left            =   1320
      Picture         =   "laWelcom.frx":0000
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdExit_Click()
    End
    
End Sub

Private Sub Form_Load()
    Const conMsg1 As String = "Hello "
    Const conMsg2 As String = "This is an example of a global variable!"
    lblWelcome.Caption = conMsg1 & strName & "," & vbNewLine & vbNewLine & conMsg2
    
End Sub






