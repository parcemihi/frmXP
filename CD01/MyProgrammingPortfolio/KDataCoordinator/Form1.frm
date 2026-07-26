VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5640
      TabIndex        =   21
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1560
      TabIndex        =   20
      Top             =   5880
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1560
      TabIndex        =   19
      Top             =   5400
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   18
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   17
      Top             =   4440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   16
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   15
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   14
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   13
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cInsert 
      Caption         =   "Insert"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cAppend 
      Caption         =   "Append"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cInput 
      Caption         =   "Input"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cOutput 
      Caption         =   "Output"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim T1 As Integer, T2 As Integer, T3 As Integer, T4 As Integer, T5 As Integer, T6 As Integer, T7 As Integer, T8 As Integer, T9 As Integer, T10 As Integer

Private Sub cInsert_Click()

For i = 0 To BN
    If kdc.Index_(i) > Val(Text2.Text) Then
    BN = BN + 1
    j = BN
    Do Until j = i
        kdc.Index_(j) = kdc.Index_(j - 1)
        kdc.Fields(j) = kdc.Fields(j - 1)
        For l = 0 To 9
            kdc.bank(l, j) = kdc.bank(l, j - 1)
        Next l
        j = j - 1
    Loop
    
    kdc.Index_(j) = Val(Text2.Text)
    For l = 0 To 9
        kdc.bank(l, j) = Val(Text1(l).Text)
    Next l
    kdc.Fields(j) = l
    
    
    
    Exit For
    End If
Next i

Combo1.AddItem kdc.Index_(i), i

End Sub

Private Sub cOutput_Click()
Print ""

End Sub

Private Sub cAppend_Click()
Print ""

i = Combo1.ListIndex

For i = 0 To 9
kdc.bank(i, Combo1.ListIndex) = Val(Text1(i).Text)
Next i






End Sub

Private Sub Combo1_Click()
Print ""

For i = 0 To 9
Text1(i).Text = kdc.bank(i, Combo1.ListIndex)
Next i
End Sub


Private Sub Form_Load()
Fins (App.Path & "\custom1.dat")

For i = 0 To BN
Combo1.AddItem kdc.Index_(i)
Next i

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
