VERSION 5.00
Begin VB.Form frmWithdrawls 
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   645
      TabIndex        =   9
      Top             =   1530
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   2
      Left            =   390
      TabIndex        =   8
      Top             =   1170
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   7
      Top             =   810
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   450
      Width           =   1065
   End
End
Attribute VB_Name = "frmWithdrawls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Dim i As Integer
For i = 0 To totChks



Checks(CheckIndx).c1Numb = 1









Next i


End Sub
