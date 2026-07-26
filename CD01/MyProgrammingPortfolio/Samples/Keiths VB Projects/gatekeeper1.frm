VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Answer"
      Height          =   1095
      Left            =   3120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "YOU"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ARE"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CommandButton2 
      Caption         =   "Backspace"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton CommandButton1 
      Caption         =   "HOW"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox ListBox1 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    EntryCount = EntryCount '+ 1
    ListBox1.AddItem (EntryCount & "HOW")
End Sub
Private Sub Command1_Click()
    EntryCount = EntryCount '+ 1
    ListBox1.AddItem (EntryCount & "ARE")
End Sub

Private Sub Command2_Click()
    EntryCount = EntryCount '+ 1
    ListBox1.AddItem (EntryCount & "YOU")
End Sub

Private Sub Command3_Click()
If EntryCount = True Then

Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "online help?"
Style = vbYesNo + vbApplicationModal + vbDefaultButton2
Title = "MsgBox Demonstration"
Help = "DEMO.HLP"
Ctxt = 1000


Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then GoTo Line1 Else GoTo Line2

Line1:
MyAbout = "Yes"

Line2:
MyAbout = "No"

End If
End Sub


Private Sub CommandButton2_Click()
    ListBox1.SetFocus

    'Ensure ListBox contains list items
    If ListBox1.ListCount >= 1 Then
        'If no selection, choose last list item.
        If ListBox1.ListIndex = -1 Then
            ListBox1.ListIndex = ListBox1.ListCount - 1
        End If
        ListBox1.RemoveItem (ListBox1.ListIndex)
    End If
End Sub

Private Sub UserForm_Initialize()
    EntryCount = 0
    CommandButton1.Caption = "Add Item"
    CommandButton2.Caption = "Remove Item"
    CommandButton3.Caption = "1 Add Item"
    CommandButton4.Caption = "1 Remove Item"
End Sub
