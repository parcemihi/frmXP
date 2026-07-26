VERSION 5.00
Begin VB.Form frmConversion 
   Caption         =   "Temperature Conversions"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Calculate new tempurature"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "conversion.frx":0000
      Left            =   2280
      List            =   "conversion.frx":000D
      TabIndex        =   2
      Text            =   "Fahrenheit"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "conversion.frx":002E
      Left            =   240
      List            =   "conversion.frx":003B
      TabIndex        =   1
      Text            =   "Celsius"
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   240
      OLEDragMode     =   1  'Automatic
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "to"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim K As Integer, C As Integer, F As Integer, err
K = Str(Val(Text1.Text))
C = Str(Val(Text1.Text))
F = Str(Val(Text1.Text))


If Combo1.Text = "Celsius" Then
    If Combo2.Text = "Kelvin" Then
    Text1.Text = FormatNumber(C + 273, 1)
    ElseIf Combo2.Text = "Fahrenheit" Then
    Text1.Text = FormatNumber((9 * C) / 5 + 32, 1)
    End If

    If Combo2.Text = "Celsius" Then
    err = MsgBox("Please select a conversion method", vbOKOnly, "No conversion")
    Combo1.SetFocus
    End If

End If

If Combo1.Text = "Kelvin" Then
    If Combo2.Text = "Celsius" Then
    Text1.Text = FormatNumber(K - 273, 1)
    ElseIf Combo2.Text = "Fahrenheit" Then
    Text1.Text = FormatNumber((9 * (K - 273)) / 5 + 32, 1)
    End If
    
    If Combo2.Text = "Kelvin" Then
    err = MsgBox("Please select a conversion method", vbOKOnly, "No conversion")
    Combo1.SetFocus
    End If

End If

If Combo1.Text = "Fahrenheit" Then
    If Combo2.Text = "Kelvin" Then
    Text1.Text = FormatNumber((5 / 9 * (F - 32)) + 273, 1)
    ElseIf Combo2.Text = "Celsius" Then
    Text1.Text = FormatNumber(5 / 9 * (F - 32), 1)
    End If
    
    If Combo2.Text = "Fahrenheit" Then
    err = MsgBox("Please select a conversion method", vbOKOnly, "No conversion")
    Combo1.SetFocus
    End If


End If

Text1.SetFocus
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)





End Sub

Private Sub exit_Click()
End
End Sub

'Private Sub Text1_GotFocus()
'Text1.Text = ""

'End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Command1_Click
Text1.SetFocus

End If
End Sub
