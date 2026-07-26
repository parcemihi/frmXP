VERSION 5.00
Begin VB.Form frmDebug3 
   Caption         =   "Debugging Exercise 3"
   ClientHeight    =   2325
   ClientLeft      =   2355
   ClientTop       =   1800
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Display"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDebug3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDisplay_Click()
    lblMessage.Caption = "OK"
End Sub


Private Sub cmdExit_Click()
    End
End Sub


