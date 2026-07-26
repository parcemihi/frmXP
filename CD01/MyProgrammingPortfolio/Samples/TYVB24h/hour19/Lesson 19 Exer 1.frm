VERSION 5.00
Begin VB.Form frmShapes 
   Caption         =   "Shapes"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Change Shape"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape shpShape 
      Height          =   975
      Left            =   2520
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Static intShape As Integer
  If (intShape = 5) Then
    intShape = 0
  Else
    intShape = intShape + 1
  End If
  shpShape.Shape = intShape
End Sub

Private Sub Command2_Click()
  Unload Me
  End
End Sub
