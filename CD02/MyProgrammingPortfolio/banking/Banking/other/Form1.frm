VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Part(0 To 9) As String
Dim Part1 As Date
Dim Part2 As Integer
Dim Part3 As String
Dim Part4 As Double
Dim Part5 As Double
Dim Part6 As Double
Dim Part7 As Double


Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Open App.Path & "\account.dat" For Input As #1
Do Until EOF(1)
    Input #1, Part(0), Part(1), Part(2), Part(3), Part(4), Part(5), Part(6)
If Part(3) = "00/00/00" Then
Part(3) = "01/01/50"
End If
Part1 = Part(3)
Part2 = Part(4)
Part3 = Part(5)
Part4 = -Part(6)
Part5 = Part(0)
Part6 = Part(1)
Part7 = Part(2)
    Open App.Path & "\format.dat" For Append As #2
        Write #2, Part1, Part2, Part3, Part4, Part5, Part6
    Close #2
Loop
Close #1
End Sub

