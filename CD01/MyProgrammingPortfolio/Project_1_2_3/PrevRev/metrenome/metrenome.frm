VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KEG Metrenome"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Sensitivity"
      Height          =   855
      Left            =   233
      TabIndex        =   6
      Top             =   2280
      Width           =   3255
      Begin MSComctlLib.Slider Slider2 
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         _Version        =   393216
         SmallChange     =   100
         Min             =   57000
         Max             =   63000
         SelectRange     =   -1  'True
         SelStart        =   60000
         TickFrequency   =   300
         Value           =   58500
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BPM"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
      Begin MSComctlLib.Slider Slider1 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   873
         _Version        =   393216
         Min             =   40
         Max             =   240
         SelectRange     =   -1  'True
         SelStart        =   40
         TickFrequency   =   10
         Value           =   40
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1913
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2993
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   353
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   1673
      Top             =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2393
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   233
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   1313
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   2273
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   3233
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   353
      Shape           =   3  'Circle
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, k As Integer, j As Integer, bpm As Integer
Dim sensitivity As Long

Private Sub Command2_Click()
If Not Label1.Caption = "240" Then
bpm = bpm + 1
j = (sensitivity / k) / bpm
Slider1.Value = bpm
Label1.Caption = bpm
End If
End Sub

Private Sub Command3_Click()
If Not Label1.Caption = "40" Then
bpm = bpm - 1
j = (sensitivity / k) / bpm
Slider1.Value = bpm
Label1.Caption = bpm
End If
End Sub

Private Sub Form_Load()
k = 1
sensitivity = 58500
Slider2.Value = 58500
bpm = 120
j = (sensitivity / k) / bpm
Slider1.Value = bpm
i = 0
Label1.Caption = bpm
Command1.Caption = "Start"
Timer1.Enabled = False
End Sub

Private Sub Slider1_Click()
j = (sensitivity / k) / Slider1.Value
Label1.Caption = Slider1.Value
bpm = Slider1.Value
End Sub

Private Sub Slider2_Click()
sensitivity = Slider2.Value + 5
j = (sensitivity / k) / Slider1.Value
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = j
i = i + 1
Select Case i
    Case 1
        Shape1(0).Visible = True
        Shape1(3).Visible = False
        Beep
    Case 2
        Shape1(1).Visible = True
        Shape1(0).Visible = False
        Beep
    Case 3
        Shape1(2).Visible = True
        Shape1(1).Visible = False
        Beep
    Case 4
        Shape1(3).Visible = True
        Shape1(2).Visible = False
        Beep
        i = 0
End Select
End Sub


Private Sub Command1_Click()
    If Timer1 = True Then
    Timer1.Enabled = False
    Command1.Caption = "Start"
    i = 0
    For i = 0 To 3
    
    Next i
    Else: Timer1.Enabled = True
    Command1.Caption = "Stop"
    Timer1_Timer
    
    End If
End Sub


