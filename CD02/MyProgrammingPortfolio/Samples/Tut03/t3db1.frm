VERSION 5.00
Begin VB.Form frmAverage 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "Average Calculator"
   ClientHeight    =   1905
   ClientLeft      =   1470
   ClientTop       =   2025
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1905
   ScaleWidth      =   3885
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "E&xit"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdAvg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Calculate Average"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Average:"
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
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   825
   End
   Begin VB.Label lblAnswer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmAverage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAvg_Click()
    Dim intNum1 As Integer
    Dim intNum2 As Integer
    Dim sngAvg As Integer
    intNum1 = Val(InputBox("Enter the first number:", "First Number"))
    intNum2 = Val(InputBox("Enter the second number:", "Second Number"))
    sngAvg = (intNum1 + intNum2) / 2
    lblAnswer.Caption = sngAvg

End Sub

Private Sub CmdExit_Click()
    End
End Sub

Private Sub Form_Load()
    frmAverage.Top = (Screen.Height - frmAverage.Height) / 2
    frmAverage.Left = (Screen.Width - frmAverage.Width) / 2
End Sub


