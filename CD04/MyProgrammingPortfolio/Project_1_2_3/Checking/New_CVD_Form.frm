VERSION 5.00
Begin VB.Form New_CVD_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton New_CVD 
      Caption         =   "E&xit"
      Height          =   495
      Index           =   3
      Left            =   2880
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton New_CVD 
      Caption         =   "&Deposit"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton New_CVD 
      Caption         =   "&VISA"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton New_CVD 
      Caption         =   "&Check"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "New_CVD_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub New_CVD_Click(Index As Integer)
Dim i As Integer
If Index < 3 Then
    TotalEntries = TotalEntries + 1
    EntryPointer = TotalEntries
    For i = 0 To 6
        account.txtCheckEdit(i).Text = ""
    Next i
End If
Select Case Index
    Case 0
    checking.account(checking.EntryPointer).ENT0_CVD_TYPE = 0
    account.ChkDep_OPbttn(0).Value = 1
    Case 1
    checking.account(checking.EntryPointer).ENT0_CVD_TYPE = 1
    account.ChkDep_OPbttn(1).Value = 1
    Case 2
    checking.account(checking.EntryPointer).ENT0_CVD_TYPE = 2
    account.ChkDep_OPbttn(2).Value = 1
End Select
    account.checkboxCleared.Value = 0
    account.txtCheckEdit(4).Text = 2002
    account.Enabled = True
    account.txtCheckEdit(6).SetFocus
    account.UpdateData_Click
    Unload New_CVD_Form
End Sub
