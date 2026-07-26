VERSION 5.00
Begin VB.Form loadform 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "frmcapt"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "introload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Proceed to Game"
      Height          =   495
      Left            =   1500
      TabIndex        =   8
      Top             =   -15
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   135
      TabIndex        =   7
      Top             =   -15
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrlblmove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer tmrmove 
      Interval        =   50
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer tmrcounter 
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer tmrpset 
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to continue?"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblintro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "written by Marty Forde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   9615
   End
   Begin VB.Label lblintro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"introload.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6735
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   8280
      Width           =   9615
   End
   Begin VB.Label lblintro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "bEgInNeRs pOnG v.1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "introload.frx":04A9
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "loadform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'written by Marty Forde

Private Sub Command1_Click()
    
    'yesno msgbox dialog
    
    Dim password As String
    
    password = MsgBox("Level Objectives:" & vbNewLine & "1. Stay alive for 200 seconds" & vbNewLine & "2.Get 1000 points" & vbNewLine & "Controls:" & vbNewLine & "1.Use the up and down on the keypad to move the paddle" & vbNewLine & "Note: Ball will move faster with every hit and at certain points the background will change without notice" & vbNewLine & "Are u sure u want to continue?", vbYesNo, "bEgInNeRs pOnG v.1.0")
    
    If password = 6 Then
        playform.Show
        Unload loadform
    ElseIf password = 7 Then
        MsgBox "Good-Bye and Good Choice", , "bEgInNeRs pOnG v.1.0"
        End
    End If
    
End Sub

Private Sub Command2_Click()
     
    'ends game
     
    End
    MsgBox "Good-Bye and Good Choice", , "bEgInNeRs pOnG v.1.0"
    
End Sub

Private Sub Command3_Click()
    
    End
    
End Sub

Private Sub Command4_Click()
 'yesno msgbox dialog
    
    Dim password As String
    
    password = MsgBox("Level Objectives:" & vbNewLine & "1. Stay alive for 200 seconds" & vbNewLine & "2.Get 1000 points" & vbNewLine & "Controls:" & vbNewLine & "1.Use the up and down on the keypad to move the paddle" & vbNewLine & "Note: Ball will move faster with every hit and at certain points the background will change without notice" & vbNewLine & "Are u sure u want to continue?", vbYesNo, "bEgInNeRs pOnG v.1.0")
    
    If password = 6 Then
        playform.Show
        Unload loadform
    ElseIf password = 7 Then
        MsgBox "Good-Bye and Good Choice", , "bEgInNeRs pOnG v.1.0"
        End
    End If
End Sub

Private Sub Form_Load()
    
    'sets num variable
    
    num = 0

End Sub

Private Sub Timer1_Timer()
    
End Sub

Private Sub tmrcounter_Timer()
    
    'counts num added up
    
    Label1.Caption = Val(Label1.Caption + 1)

End Sub

Private Sub tmrlblmove_Timer()
    
    'moves the label
    
    lblintro(0).Move lblintro(0).Left + 0, lblintro(0).Top - 25
    lblintro(1).Move lblintro(1).Left + 0, lblintro(1).Top - 25
    lblintro(2).Move lblintro(2).Left + 0, lblintro(2).Top - 25
    
    num = num + 1
    
    'if the labels are at a certain height then make the command
    'buttons visible
    
    If num = 580 Then
        Label2.Visible = True
        Command1.Visible = True
        Command2.Visible = True
    End If
    
End Sub

Private Sub tmrmove_Timer()
    
    'moves the happy face
    
    If Label1.Caption < "3" Then
        Image1.Move Image1.Left + 60, Image1.Top + 250
    ElseIf Label1.Caption = "3" Or Label1.Caption = "4" Then
        Image1.Move Image1.Left + 60, Image1.Top - 250
    ElseIf Label1.Caption = "5" Or Label1.Caption = "6" Then
        Image1.Move Image1.Left + 60, Image1.Top + 250
    ElseIf Label1.Caption = "7" Then
        Image1.Move Image1.Left + 60, Image1.Top - 250
    ElseIf Label1.Caption = "8" Then
        Image1.Move Image1.Left + 60, Image1.Top + 250
        tmrlblmove.Enabled = True
    End If
    
End Sub

Private Sub tmrpset_Timer()
    
    'dot movement
    
    Call P_Set

End Sub

Public Sub P_Set()
    
    'pset dot movement
    
    R = 255 * Rnd
    G = 255 * Rnd
    B = 255 * Rnd
    xpos = Rnd * ScaleWidth
    ypos = Rnd * ScaleHeight
    PSet (xpos, ypos), RGB(R, G, B)

End Sub
