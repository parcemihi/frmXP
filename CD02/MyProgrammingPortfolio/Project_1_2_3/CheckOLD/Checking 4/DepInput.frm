VERSION 5.00
Begin VB.Form DepInput 
   Caption         =   "Form2"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Done"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   1590
      Width           =   1215
   End
   Begin VB.TextBox DepInput 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   1110
      Width           =   2535
   End
   Begin VB.TextBox DepInput 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   750
      Width           =   2535
   End
   Begin VB.TextBox DepInput 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   390
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   2
      Left            =   615
      TabIndex        =   6
      Top             =   1155
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   795
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   4
      Top             =   435
      Width           =   345
   End
End
Attribute VB_Name = "DepInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
DepInput(0).Enabled = False



End Sub
