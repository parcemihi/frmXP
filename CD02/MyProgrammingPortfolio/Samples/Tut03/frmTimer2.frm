VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Funky stopwatch. (watch the time fly like a drunk duck)"
   ClientHeight    =   3870
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7110
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer6 
      Left            =   840
      Top             =   1560
   End
   Begin VB.Timer Timer5 
      Left            =   1800
      Top             =   1560
   End
   Begin VB.Timer Timer4 
      Left            =   2880
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Left            =   3840
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Left            =   4920
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   1560
   End
   Begin VB.Image cmdStart 
      Height          =   750
      Left            =   4800
      Picture         =   "frmTimer2.frx":0000
      Top             =   2640
      Width           =   1800
   End
   Begin VB.Image cmdReset 
      Height          =   750
      Left            =   2640
      Picture         =   "frmTimer2.frx":5A45
      Top             =   2640
      Width           =   1800
   End
   Begin VB.Image cmdStop 
      Height          =   750
      Left            =   480
      Picture         =   "frmTimer2.frx":B682
      Top             =   2640
      Width           =   1800
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00400000&
      BorderWidth     =   10
      Height          =   3615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   6855
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   1680
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   840
      Width           =   135
   End
   Begin VB.Image img6 
      Height          =   1605
      Left            =   480
      Top             =   480
      Width           =   840
   End
   Begin VB.Image img5 
      Height          =   1605
      Left            =   1440
      Top             =   480
      Width           =   840
   End
   Begin VB.Image img4 
      Height          =   1605
      Left            =   2640
      Top             =   480
      Width           =   840
   End
   Begin VB.Image img3 
      Height          =   1605
      Left            =   3600
      Top             =   480
      Width           =   840
   End
   Begin VB.Image img2 
      Height          =   1605
      Left            =   4800
      Top             =   480
      Width           =   840
   End
   Begin VB.Image img1 
      Height          =   1605
      Left            =   5760
      Top             =   480
      Width           =   840
   End
   Begin VB.Image Image0 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":11124
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":1178F
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":11BCE
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image3 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":121E5
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image4 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":127F1
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image5 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":12D49
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image6 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":13362
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image7 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":139DA
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image8 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":13E97
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image9 
      Height          =   1605
      Left            =   3120
      Picture         =   "frmTimer2.frx":14576
      Top             =   1440
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00400000&
      BorderWidth     =   30
      Height          =   4095
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   -120
      Width           =   7095
   End
   Begin VB.Menu exit 
      Caption         =   "E&xit"
      NegotiatePosition=   3  'Right
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public start As Integer
Private Sub cmdStop_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False

End Sub


Private Sub exit_Click()
End

End Sub

Private Sub Form_Load()
img1.Picture = Image0
img2.Picture = Image0
img3.Picture = Image0
img4.Picture = Image0
img5.Picture = Image0
img6.Picture = Image0
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    
    Timer1.Interval = 10
    Timer2.Interval = 100
    Timer3.Interval = 1000
End Sub

Private Sub cmdStart_Click()
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True

End Sub

Private Sub cmdReset_Click()
img1.Picture = Image0
img2.Picture = Image0
img3.Picture = Image0
img4.Picture = Image0
img5.Picture = Image0
img6.Picture = Image0

End Sub

Private Sub Timer1_Timer()
    
    If img1.Picture = Image0 Then
    img1.Picture = Image1
    ElseIf img1.Picture = Image1 Then
    img1.Picture = Image2
    ElseIf img1.Picture = Image2 Then
    img1.Picture = Image3
    ElseIf img1.Picture = Image3 Then
    img1.Picture = Image4
    ElseIf img1.Picture = Image4 Then
    img1.Picture = Image5
    ElseIf img1.Picture = Image5 Then
    img1.Picture = Image6
    ElseIf img1.Picture = Image6 Then
    img1.Picture = Image7
    ElseIf img1.Picture = Image7 Then
    img1.Picture = Image8
    ElseIf img1.Picture = Image8 Then
    img1.Picture = Image9
    ElseIf img1.Picture = Image9 Then
    img1.Picture = Image0
    End If
End Sub

Private Sub Timer2_Timer()
    If img2.Picture = Image0 Then
    img2.Picture = Image1
    ElseIf img2.Picture = Image1 Then
    img2.Picture = Image2
    ElseIf img2.Picture = Image2 Then
    img2.Picture = Image3
    ElseIf img2.Picture = Image3 Then
    img2.Picture = Image4
    ElseIf img2.Picture = Image4 Then
    img2.Picture = Image5
    ElseIf img2.Picture = Image5 Then
    img2.Picture = Image6
    ElseIf img2.Picture = Image6 Then
    img2.Picture = Image7
    ElseIf img2.Picture = Image7 Then
    img2.Picture = Image8
    ElseIf img2.Picture = Image8 Then
    img2.Picture = Image9
    ElseIf img2.Picture = Image9 Then
    img2.Picture = Image0
    End If

End Sub

Private Sub Timer3_Timer()
    If img3.Picture = Image0 Then
    img3.Picture = Image1
    ElseIf img3.Picture = Image1 Then
    img3.Picture = Image2
    ElseIf img3.Picture = Image2 Then
    img3.Picture = Image3
    ElseIf img3.Picture = Image3 Then
    img3.Picture = Image4
    ElseIf img3.Picture = Image4 Then
    img3.Picture = Image5
    ElseIf img3.Picture = Image5 Then
    img3.Picture = Image6
    ElseIf img3.Picture = Image6 Then
    img3.Picture = Image7
    ElseIf img3.Picture = Image7 Then
    img3.Picture = Image8
    ElseIf img3.Picture = Image8 Then
    img3.Picture = Image9
    ElseIf img3.Picture = Image9 Then
    img3.Picture = Image0

    If img4.Picture = Image0 Then
    img4.Picture = Image1
    ElseIf img4.Picture = Image1 Then
    img4.Picture = Image2
    ElseIf img4.Picture = Image2 Then
    img4.Picture = Image3
    ElseIf img4.Picture = Image3 Then
    img4.Picture = Image4
    ElseIf img4.Picture = Image4 Then
    img4.Picture = Image5
    ElseIf img4.Picture = Image5 Then
    img4.Picture = Image0

    If img5.Picture = Image0 Then
    img5.Picture = Image1
    ElseIf img5.Picture = Image1 Then
    img5.Picture = Image2
    ElseIf img5.Picture = Image2 Then
    img5.Picture = Image3
    ElseIf img5.Picture = Image3 Then
    img5.Picture = Image4
    ElseIf img5.Picture = Image4 Then
    img5.Picture = Image5
    ElseIf img5.Picture = Image5 Then
    img5.Picture = Image6
    ElseIf img5.Picture = Image6 Then
    img5.Picture = Image7
    ElseIf img5.Picture = Image7 Then
    img5.Picture = Image8
    ElseIf img5.Picture = Image8 Then
    img5.Picture = Image9
    ElseIf img5.Picture = Image9 Then
    img5.Picture = Image0

    If img6.Picture = Image0 Then
    img6.Picture = Image1
    ElseIf img6.Picture = Image1 Then
    img6.Picture = Image2
    ElseIf img6.Picture = Image2 Then
    img6.Picture = Image3
    ElseIf img6.Picture = Image3 Then
    img6.Picture = Image4
    ElseIf img6.Picture = Image4 Then
    img6.Picture = Image5
    ElseIf img6.Picture = Image5 Then
    img6.Picture = Image0
    End If
    End If
    End If
    End If

End Sub
