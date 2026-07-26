VERSION 5.00
Begin VB.Form playform 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   Icon            =   "playform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   1560
   End
   Begin VB.Timer tmrbw 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer Tmrcntdwn 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Time 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7080
      Top             =   6720
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To quit press the mouse buttons or press enter. Gracias."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "by cOrKsCrEw"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "bEgInNeRs pOnG v.1"
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   6600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label lbllives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lives:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "200"
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
      Height          =   495
      Left            =   8160
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   960
      Top             =   3960
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   375
      Left            =   6000
      Shape           =   2  'Oval
      Top             =   3720
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   9480
      Y1              =   6600
      Y2              =   6600
   End
End
Attribute VB_Name = "playform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''
'written by Marty Forde              '
''''''''''''''''''''''''''''''''''''''
'



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    lblcode = KeyCode
    
    'if user presses enter key end the program
    
    If KeyCode = 13 Then
        End
    End If
    
    'if the keyboad is pressed a certain thing will happen
    
    If KeyCode = 38 Then
        If Shape2.Top <= 0 Then
            Exit Sub
        End If
        Shape2.Move Shape2.Left - 0, Shape2.Top - pm
    ElseIf KeyCode = 40 Then
        If Shape2.Top >= (playform.Height - Shape2.Height) - 680 Then
            Exit Sub
        End If
        Shape2.Move Shape2.Left - 0, Shape2.Top + pm
    End If

End Sub

Private Sub Form_Load()
    
    'sets balls pace and speed
    
    bm = 250
    pm = 250
    Motion = 1

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'if the user presses a mouse key then end program
    
    If Button = 1 Then
        End
    ElseIf Button = 2 Then
        End
    ElseIf Button = 3 Then
        End
    End If
    
End Sub

Private Sub lblcode_Click()

End Sub

Private Sub Time_Timer()
    
    'counts the time left to play
    
    lbltime.Caption = Val(lbltime.Caption - 1)

End Sub

Private Sub Timer1_Timer()
    
    Let player1angle = (Shape1.Top - Shape2.Top) / 5
    
    'states if ball hits paddle randomly
    'choose a motion that bounces the ball
    'off the paddle
    
    If Shape1.Left <= Shape2.Left + Shape2.Width And Shape1.Left >= Shape2.Left - Shape2.Width Then
        If Shape1.Top + Shape1.Height >= Shape2.Top And Shape1.Top <= Shape2.Top + Shape2.Height Then
            Label1.Caption = Val(Label1.Caption + 50)
            rebound = Int((2 * Rnd) + 1)
            If rebound = 1 Then
                Motion = 2
            ElseIf rebound = 2 Then
                Motion = 3
            End If
        End If

    End If
        
    'motions' that move the ball
        
    Select Case Motion
        Case 1
            Shape1.Move Shape1.Left - bm, Shape1.Top - bm
            If Shape1.Left <= 0 Then
                
                'if the ball hits the left side of the screen
                'set the speed and countdown process
                
                Shape1.Move 6000, 3720
                bm = 250
                pm = 250
                deathplay = Int((4 * Rnd) + 1)
                
                'according to rnd num move the ball
                
                If deathplay = 1 Then
                    Motion = 1
                ElseIf deathplay = 2 Then
                    Motion = 2
                ElseIf deathplay = 3 Then
                    Motion = 3
                ElseIf deathplay = 4 Then
                    Motion = 4
                End If
                
                'starts to countdown and takes a life off
                
                Label6.Visible = True
                Label6.Caption = Val(5)
                Tmrcntdwn.Enabled = True
                Timer1.Enabled = False
                lbllives.Caption = Val(lbllives.Caption - 1)
            ElseIf Shape1.Top <= 0 Then
                
                'if ball bounces increase speed
                
                bm = bm + 5
                pm = pm + 10
                Motion = 4
            End If
        Case 2
            Shape1.Move Shape1.Left + bm, Shape1.Top - bm
            If Shape1.Left >= (playform.Width - Shape1.Width) Then
                
                'if ball bounces increase speed
                
                bm = bm + 5
                pm = pm + 10
                Motion = 1
            ElseIf Shape1.Top <= 0 Then
                
                'if ball bounces increase speed
                
                bm = bm + 5
                pm = pm + 10
                Motion = 3
            End If
        Case 3
            Shape1.Move Shape1.Left + bm, Shape1.Top + bm
            If Shape1.Left >= (playform.Width - Shape1.Width) Then
                
                'if ball bounces increase speed
                
                bm = bm + 5
                pm = pm + 10
                Motion = 4
            ElseIf Shape1.Top >= (playform.Height - Shape1.Height) - 680 Then
                
                'if ball bounces increase speed
                
                bm = bm + 5
                pm = pm + 10
                Motion = 2
            End If
        Case 4
            Shape1.Move Shape1.Left - bm, Shape1.Top + bm
            If Shape1.Left <= 0 Then
                
                'if the ball hits the left side of the screen
                'set the speed and countdown process
                
                Shape1.Move 6000, 3720
                bm = 250
                pm = 250
                deathplay = Int((4 * Rnd) + 1)
                If deathplay = 1 Then
                    Motion = 1
                ElseIf deathplay = 2 Then
                    Motion = 2
                ElseIf deathplay = 3 Then
                    Motion = 3
                ElseIf deathplay = 4 Then
                    Motion = 4
                End If
                
                'according to rnd num move the ball
                
                Label6.Visible = True
                Label6.Caption = Val(5)
                Tmrcntdwn.Enabled = True
                Timer1.Enabled = False
                lbllives.Caption = Val(lbllives.Caption - 1)
                Motion = 3
            ElseIf Shape1.Top >= (playform.Height - Shape1.Height) - 680 Then
                
                'if ball bounces increase speed
                
                bm = bm + 5
                pm = pm + 10
                Motion = 1
            End If
    End Select
    
End Sub


Private Sub Timer2_Timer()
    
    If Val(Label1.Caption) >= 1000 And Val(lbllives.Caption) > 0 And Val(lbltime.Caption) > 0 Then
        
        'if you win the game then stop ball movement
        'and end game
        
        Timer1.Enabled = False
        Time.Enabled = False
        MsgBox "you beat the level good for you", , "bEgInNeRs pOnG v.1.0"
        Timer2.Enabled = False
        End
    ElseIf lbllives.Caption = "0" Then
        
        'if you lose end game
        
        MsgBox "im sorry you lost", , "bEgInNeRs pOnG v.1.0"
        'This shuts down the system!  That's probaby why the author's
        'name for the game is EvilPong.  Removed by FreeVBCode.com
        
        'ShutdownSystem
        
        End
    ElseIf lbltime.Caption = "0" And Val(Label1.Caption) <> 1000 Then
        
        'if you lose end game and restart
        
        MsgBox "im sorry you lost", , "bEgInNeRs pOnG v.1.0"
        'See above
        'ShutdownSystem
        
        End
    End If
    
    If Label1.Caption = "200" Then
        
        'if you have 200 points then change background color
        
        playform.BackColor = vbRed
    ElseIf Label1.Caption = "400" Then
        
        'if points = 400 then go to black\white sequence
        
        tmrbw.Enabled = True
    ElseIf Label1.Caption = "600" Then
        
        'if you have 600 points then change background color
        
        tmrbw.Enabled = False
        Shape2.BackColor = &HFFFF00
        Shape1.BackColor = &HFFFF00
        playform.BackColor = &H808080
    ElseIf Label1.Caption = "800" Then
        
        'if you have 800 points then change background color
        'and set the paddle speed
        
        playform.BackColor = vbBlack
        Shape1.BackColor = vbWhite
        Shape2.BackColor = vbWhite
        pm = 1000
    Else
        Exit Sub
    End If
    
End Sub

Private Sub tmrbw_Timer()
    
    'black\white sequence
    
    If playform.BackColor = vbBlack Then
        playform.BackColor = vbWhite
    Else
        playform.BackColor = vbBlack
    End If
    
End Sub

Private Sub Tmrcntdwn_Timer()
    
    'if ball hits left then start new countdown
    
    Label6.Caption = Val(Label6.Caption - 1)
    
    If Label6.Caption = 0 Then
        Label6.Visible = False
        Time.Enabled = True
        Timer1.Enabled = True
        Tmrcntdwn.Enabled = False
    End If
    
End Sub

