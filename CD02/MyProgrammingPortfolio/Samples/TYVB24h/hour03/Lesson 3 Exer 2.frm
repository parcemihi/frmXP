VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Exit program"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Message"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      ToolTipText     =   "Click to see message"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblClick 
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "The message appears here!"
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  ' Make sure the label's name is lblClick
  lblClick.Caption = "Clicked!"
End Sub

Private Sub Command2_Click()
  ' Ends the running application
  End
End Sub
