VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   2880
      Top             =   1800
   End
   Begin VB.Shape Shape2 
      Height          =   360
      Left            =   3600
      Top             =   5400
      Width           =   360
   End
   Begin VB.Shape Shape1 
      Height          =   5400
      Left            =   2160
      Top             =   360
      Width           =   3600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim center As Integer, down As Integer, msg As String, time As Integer
Private Sub Form_KeyDown(KeyCode As Integer, _
        Shift As Integer)
    If KeyCode = vbKeyLeft Then
        If center = 2160 Then
            center = 2520
        Else
        center = center - 360
        Shape2.Left = center
        End If
    End If
    If KeyCode = vbKeyRight Then
        If center = 5400 Then
            center = 5040
        End If
        center = center + 360
        Shape2.Left = center
    End If
    If KeyCode = vbKeyF2 Then
        down = 0
        Timer1.Enabled = True
        Shape2.Visible = True




'        Timer1.Enabled = False
'        Else: Timer1.Enabled = True
'        End If
    End If
    If KeyCode = vbKeyF3 Then
        msg = MsgBox("Pause", vbOKOnly, "Game paused")
        If vbOKOnly = 0 Then
        Timer1.Enabled = True
        End If
    End If

End Sub

Private Sub Form_Load()
    time = 950
    center = 3600
    Timer1.Enabled = False
    Timer1.Interval = time
    Shape2.Left = center
    Shape2.Top = down + 360
    Shape2.Visible = False
End Sub


Private Sub Timer1_Timer()
    down = down + 360
    Shape2.Top = down
    
    If Shape2.Top = 5400 Then
        down = 0
        If time = 50 Then
        time = 50
        Else
        time = time - 150
        End If
        Timer1.Interval = time
    End If
End Sub

