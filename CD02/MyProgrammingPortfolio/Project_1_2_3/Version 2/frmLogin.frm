VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3120
   ClientLeft      =   2835
   ClientTop       =   3765
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   248
      TabIndex        =   9
      Top             =   2160
      Width           =   5415
      Begin VB.OptionButton Option1 
         Caption         =   "No Access"
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Full Access"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "User Access"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   248
      TabIndex        =   0
      Top             =   600
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   975
      Left            =   3008
      TabIndex        =   2
      Top             =   1080
      Width           =   2625
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   975
      Left            =   248
      TabIndex        =   3
      Top             =   1080
      Width           =   2625
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3008
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2625
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   248
      TabIndex        =   7
      Top             =   240
      Width           =   2625
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   3008
      TabIndex        =   8
      Top             =   240
      Width           =   2625
   End
   Begin VB.Menu mnuCreateUser 
      Caption         =   "CreateUser"
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LoginOk As Boolean

Private Sub cmdOK_Click()
Dim login As String
login = userpass.Login_UserPass(Text1.Text, Text2.Text)

























End Sub

Private Sub mnuCreateUser_Click()
Dim CreateUser As String
Dim Access As Integer
Dim i As Integer
Login_Error
If LoginOk = True Then
    For i = 1 To 3
    If Option1(i).Value = True Then
    Access = i
    Exit For
    End If
    Next i
    CreateUser = userpass.CreateUser(Text1.Text, Text2.Text, Access)
End If
End Sub

Public Function Login_Error()
Dim Msg As String
    If Text1.Text = "" Then
    Msg = MsgBox("You must supply a user name.", vbExclamation, "Login Error")
    LoginOk = False
ElseIf Text2.Text = "" Then
    Msg = MsgBox("The password you entered is invalid.", vbExclamation, "Login Error")
    LoginOk = False
    Else
    LoginOk = True
    End If
End Function
Private Sub Form_Load()
Frame1.Visible = False
mnuCreateUser.Visible = False
End Sub
Private Sub cmdCancel_Click()
End
End Sub
