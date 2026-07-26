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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Answer to the question."
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   735
         Left            =   3120
         TabIndex        =   2
         Top             =   1800
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Click the pick."
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2775
         Begin VB.Image imgDoor 
            Height          =   495
            Left            =   960
            Picture         =   "lc2Done.frx":0000
            Top             =   720
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Image imgHand 
            Height          =   900
            Left            =   720
            Picture         =   "lc2Done.frx":0E2B
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmdNo 
         Caption         =   "&No"
         Height          =   735
         Left            =   3120
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
         Height          =   735
         Left            =   3120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Are we having fun yet?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2400
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub
Private Sub cmdNo_Click()
imgHand.Visible = True
imgDoor.Visible = False
End Sub
Private Sub cmdYes_Click()
imgDoor.Visible = True
imgHand.Visible = False
End Sub
Private Sub imgDoor_Click()
Dim Msg, Style, Title, Response, MyString
Msg = "                   You are ?"
Style = vbYesNo
Title = "MsgBox Demonstration"
Response = MsgBox(Msg, Style, Title)

If Response = vbYes Then
imgDoor.Visible = False
imgHand.Visible = False
Else
imgDoor.Visible = False
imgHand.Visible = False
End If

End Sub

Private Sub imgHand_Click()
Dim Msg, Style, Title, Response, MyString
Msg = "               No we are not."
Style = vbOK
Title = "MsgBox Demonstration"
Response = MsgBox(Msg, Style, Title)

If Response = vbYes Then
imgDoor.Visible = False
imgHand.Visible = False
Else
imgDoor.Visible = False
imgHand.Visible = False
End If

End Sub
