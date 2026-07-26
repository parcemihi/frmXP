VERSION 5.00
Begin VB.Form editor 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ExitEdit 
      Caption         =   "E&xit"
      Height          =   405
      Left            =   3000
      TabIndex        =   3
      Top             =   1995
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton DeleteSubject 
      Caption         =   "&Delete Subject"
      Height          =   405
      Left            =   3000
      TabIndex        =   2
      Top             =   1490
      Width           =   2535
   End
   Begin VB.CommandButton AddSubject 
      Caption         =   "&Add Subject"
      Height          =   405
      Left            =   3000
      TabIndex        =   1
      Top             =   985
      Width           =   2535
   End
   Begin VB.ListBox SubjectList 
      Height          =   1815
      Left            =   660
      TabIndex        =   4
      Top             =   585
      Width           =   2175
   End
End
Attribute VB_Name = "editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer
For i = 0 To lib_vhs.Combo1.ListCount
    If Not lib_vhs.Combo1.List(i) = "" Then
        SubjectList.AddItem lib_vhs.Combo1.List(i)
    End If
Next i
End Sub

Private Sub AddSubject_Click()
Dim i As Integer
Dim match As Boolean
match = False
i = -1
If Not Text1.Text = "" Then
    Do Until i = SubjectList.ListCount
        i = i + 1
        If Text1.Text = SubjectList.List(i) Then
            match = True
            i = SubjectList.ListCount
        End If
    Loop
    If match = False Then
        SubjectList.AddItem Text1.Text
    End If
End If
End Sub

Private Sub DeleteSubject_Click()
If Not SubjectList.ListIndex = -1 Then
    SubjectList.RemoveItem SubjectList.ListIndex
End If
End Sub

Private Sub ExitEdit_Click()
Form_Unload 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
lib_vhs.Combo1.Clear

Open App.Path & "\subject.dat" For Output As #1
    For i = 0 To SubjectList.ListCount - 1
        lib_vhs.Combo1.AddItem SubjectList.List(i)
        Write #1, SubjectList.List(i)
    Next i
Close #1
lib_vhs.Show
Unload editor

End Sub

Private Sub SubjectList_Click()

Text1.Text = SubjectList.List(SubjectList.ListIndex)
    
End Sub

