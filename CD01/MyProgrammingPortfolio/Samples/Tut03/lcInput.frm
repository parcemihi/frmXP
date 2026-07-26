VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Screen"
   ClientHeight    =   1215
   ClientLeft      =   1350
   ClientTop       =   2295
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1215
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   1200
   End
   Begin VB.TextBox txtSperson 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Salesperson's &name:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1905
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    strSperson = txtSperson.Text
    frmOrder.Show
    Unload frmInput
    
End Sub

Private Sub Form_Load()
    'center the form
    frmInput.Top = (Screen.Height - frmInput.Height) / 2
    frmInput.Left = (Screen.Width - frmInput.Width) / 2
    
End Sub
