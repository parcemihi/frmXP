VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Text            =   "None"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2880
      TabIndex        =   8
      Text            =   "None"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Text            =   "None"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Text            =   "None"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Door 2"
      Height          =   1575
      Left            =   4320
      TabIndex        =   4
      Top             =   375
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Door 1"
      Height          =   1575
      Left            =   3360
      TabIndex        =   3
      Top             =   375
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Gatekeeper 2"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   735
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Gatekeeper 1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   735
      Width           =   1335
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Click here for the Gatekeeper's answer."
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   1455
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Get Out Alive!"
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   495
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()

End Sub
