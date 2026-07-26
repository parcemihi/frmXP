VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   3360
      Top             =   1800
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   100
      Top             =   100
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim down As Integer




Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 200
End Sub








Private Sub Timer1_Timer()


    If Shape1.Top < 1000 Then
    
    down = down + 100
    Shape1.Top = down

    Else: Shape1.Top = 0
    End If


End Sub
