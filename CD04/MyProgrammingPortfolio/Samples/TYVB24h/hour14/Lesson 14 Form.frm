VERSION 5.00
Begin VB.Form frmMath 
   Caption         =   "Math Practice"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Test Your Math Skills"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Unload Me
  End
End Sub

Private Sub Command2_Click()
' Math skills test
   Dim lngBefore As Long
   Dim lngAfter As Long
   Dim lngTimeDiff As Long
   Dim strAns As String

   lngBefore = Timer   ' Save seconds since midnight
   Do
      strAns = InputBox("What is 150 + 235?", "Hurry")
   Loop Until Val(strAns) = 385

   lngAfter = Timer    ' Save seconds since midnight now

   ' The difference between the stored time values
   ' is how many seconds the user took to answer
   lngTimeDiff = lngAfter - lngBefore
   MsgBox ("That took you only" & Str(lngTimeDiff) & _
    " seconds!")

End Sub

Private Sub Form_Load()

End Sub
