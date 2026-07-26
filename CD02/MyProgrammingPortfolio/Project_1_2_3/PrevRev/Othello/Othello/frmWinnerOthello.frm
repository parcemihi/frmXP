VERSION 5.00
Begin VB.Form frmWinnerOthello 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Here are the results."
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdWinner 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   907
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1402
      TabIndex        =   5
      Top             =   765
      Width           =   210
   End
   Begin VB.Label lblWinnerIs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2107
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "And the winner is:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   187
      TabIndex        =   3
      Top             =   240
      Width           =   1875
   End
   Begin VB.Label lblBlack 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1912
      TabIndex        =   2
      Top             =   765
      Width           =   270
   End
   Begin VB.Label lblWhite 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "99"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   832
      TabIndex        =   1
      Top             =   765
      Width           =   270
   End
End
Attribute VB_Name = "frmWinnerOthello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdWinner_Click()
    frmOthello.Show
    frmWinnerOthello.Hide
End Sub

Private Sub Form_Load()
    If frmOthello.Wstat > frmOthello.Bstat Then
    lblWinnerIs = " White"
    ElseIf frmOthello.Bstat > frmOthello.Wstat Then
    lblWinnerIs = " Black"
    ElseIf frmOthello.Wstat = frmOthello.Bstat Then
    lblWinnerIs = " Tie!"
    End If
    
    lblBlack.Caption = frmOthello.Bstat
    lblWhite.Caption = frmOthello.Wstat
    
End Sub
