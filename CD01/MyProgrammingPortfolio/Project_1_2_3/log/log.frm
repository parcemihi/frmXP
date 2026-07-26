VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "b"
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "LOG    X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3735
      TabIndex        =   4
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim base As Integer

Private Sub Command1_Click()
Dim x As Integer
Dim Loga As String * 10
On Error GoTo ErrorHandler

List1.Clear
base = Val(Text1.Text)
For x = 1 To 10000
    Select Case base
        Case 2
            Loga = Log2(x)
        Case 3
            Loga = Log3(x)
        Case 4
            Loga = Log4(x)
        Case 5
            Loga = Log5(x)
        Case 6
            Loga = Log6(x)
        Case 7
            Loga = Log7(x)
        Case 8
            Loga = Log8(x)
        Case 9
            Loga = Log9(x)
        Case 10
            Loga = Log10(x)
        Case 11
            Loga = Log11(x)
        Case 12
            Loga = Log12(x)
        Case 13
            Loga = Log13(x)
        Case 14
            Loga = Log14(x)
        Case 15
            Loga = Log15(x)
        Case 16
            Loga = Log16(x)
        Case 17
            Loga = Log17(x)
        Case 18
            Loga = Log18(x)
        Case 19
            Loga = Log19(x)
        Case 20
            Loga = Log20(x)
        Case Else
        GoTo ErrorHandler
    End Select

List1.AddItem "  " & Loga

Next x
Exit Sub
ErrorHandler:
    Loga = MsgBox("That Logarithm is either not valid, or has not yet been implimented", vbOKOnly, "LOG Error")

End Sub

Static Function Log2(x)
   Log2 = Log(x) / Log(2#)
End Function
Static Function Log3(x)
   Log3 = Log(x) / Log(3#)
End Function
Static Function Log4(x)
   Log4 = Log(x) / Log(4#)
End Function
Static Function Log5(x)
   Log5 = Log(x) / Log(5#)
End Function
Static Function Log6(x)
   Log6 = Log(x) / Log(6#)
End Function
Static Function Log7(x)
   Log7 = Log(x) / Log(7#)
End Function
Static Function Log8(x)
   Log8 = Log(x) / Log(8#)
End Function
Static Function Log9(x)
   Log9 = Log(x) / Log(9#)
End Function
Static Function Log10(x)
   Log10 = Log(x) / Log(10#)
End Function
Static Function Log11(x)
   Log11 = Log(x) / Log(11#)
End Function
Static Function Log12(x)
   Log12 = Log(x) / Log(12#)
End Function
Static Function Log13(x)
   Log13 = Log(x) / Log(13#)
End Function
Static Function Log14(x)
   Log14 = Log(x) / Log(14#)
End Function
Static Function Log15(x)
   Log15 = Log(x) / Log(15#)
End Function
Static Function Log16(x)
   Log16 = Log(x) / Log(16#)
End Function
Static Function Log17(x)
   Log17 = Log(x) / Log(17#)
End Function
Static Function Log18(x)
   Log18 = Log(x) / Log(18#)
End Function
Static Function Log19(x)
   Log19 = Log(x) / Log(19#)
End Function
Static Function Log20(x)
   Log20 = Log(x) / Log(20#)
End Function

Private Sub Form_Resize()
List1.Height = Form1.Height - List1.Top - 100
End Sub
