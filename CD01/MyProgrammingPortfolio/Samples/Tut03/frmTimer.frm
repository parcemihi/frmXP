VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Timer Timer6 
      Left            =   480
      Top             =   1440
   End
   Begin VB.Timer Timer5 
      Left            =   1440
      Top             =   1440
   End
   Begin VB.Timer Timer4 
      Left            =   2520
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Left            =   3480
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Left            =   4560
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Left            =   5520
      Top             =   1440
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   720
      Width           =   135
   End
   Begin VB.Image img6 
      Height          =   1605
      Left            =   120
      Top             =   360
      Width           =   840
   End
   Begin VB.Image img5 
      Height          =   1605
      Left            =   1080
      Top             =   360
      Width           =   840
   End
   Begin VB.Image img4 
      Height          =   1605
      Left            =   2280
      Top             =   360
      Width           =   840
   End
   Begin VB.Image img3 
      Height          =   1605
      Left            =   3240
      Top             =   360
      Width           =   840
   End
   Begin VB.Image img2 
      Height          =   1605
      Left            =   4440
      Top             =   360
      Width           =   840
   End
   Begin VB.Image img1 
      Height          =   1605
      Left            =   5400
      Top             =   360
      Width           =   840
   End
   Begin VB.Image Image0 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image1 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":467A
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image2 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":8CF4
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image3 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":D36E
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image4 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":119E8
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image5 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":16062
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image6 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":1A6DC
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image7 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":1ED56
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image8 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":233D0
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image Image9 
      Height          =   1605
      Left            =   2760
      Picture         =   "frmTimer.frx":27A4A
      Top             =   360
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Menu exit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public start As Integer
Private Sub cmdStart_Click()
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True

End Sub
Private Sub cmdStop_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False

End Sub

Private Sub Command1_Click()
img1.Picture = Image0
img2.Picture = Image0
img3.Picture = Image0
img4.Picture = Image0
img5.Picture = Image0
img6.Picture = Image0
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
