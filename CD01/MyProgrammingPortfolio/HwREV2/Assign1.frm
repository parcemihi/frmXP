VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   7500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Update"
      Height          =   495
      Left            =   4293
      TabIndex        =   56
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton NewOrder 
      Caption         =   "&New Order"
      Height          =   495
      Left            =   480
      TabIndex        =   55
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   5565
      TabIndex        =   54
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Purchase 
      Caption         =   "&Purchase"
      Height          =   495
      Left            =   1751
      TabIndex        =   53
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Appenda 
      Caption         =   "&Append"
      Height          =   495
      Left            =   3022
      TabIndex        =   52
      Top             =   8040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   2580
      TabIndex        =   0
      Text            =   "Customer Order No."
      Top             =   563
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   9
      Left            =   2580
      TabIndex        =   11
      Top             =   3083
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   5640
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   8
      Top             =   2003
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   2580
      TabIndex        =   2
      Top             =   2003
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   2580
      TabIndex        =   3
      Top             =   2483
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   5640
      TabIndex        =   6
      Top             =   2483
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   9
      Top             =   2483
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   4
      Top             =   1523
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   1523
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   2580
      TabIndex        =   1
      Top             =   1523
      Width           =   1215
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   6
      Left            =   2040
      TabIndex        =   51
      Top             =   5655
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   50
      Top             =   5055
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   49
      Top             =   4440
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   8
      Left            =   4680
      TabIndex        =   48
      Top             =   5655
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   7
      Left            =   3360
      TabIndex        =   47
      Top             =   5655
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   5
      Left            =   4680
      TabIndex        =   46
      Top             =   5055
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   4
      Left            =   3360
      TabIndex        =   45
      Top             =   5055
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   44
      Top             =   4455
      Width           =   180
   End
   Begin VB.Label QtySft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   1
      Left            =   3360
      TabIndex        =   43
      Top             =   4455
      Width           =   180
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 900"
      Height          =   195
      Index           =   8
      Left            =   5040
      TabIndex        =   42
      Top             =   5655
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 675"
      Height          =   195
      Index           =   7
      Left            =   5040
      TabIndex        =   41
      Top             =   5055
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 450"
      Height          =   195
      Index           =   6
      Left            =   5040
      TabIndex        =   40
      Top             =   4455
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 300"
      Height          =   195
      Index           =   5
      Left            =   3720
      TabIndex        =   39
      Top             =   5655
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 225"
      Height          =   195
      Index           =   4
      Left            =   3720
      TabIndex        =   38
      Top             =   5055
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 150"
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   37
      Top             =   4455
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 100"
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   36
      Top             =   5655
      Width           =   600
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 75"
      Height          =   195
      Index           =   1
      Left            =   2400
      TabIndex        =   35
      Top             =   5055
      Width           =   510
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "At $ 50"
      Height          =   195
      Index           =   0
      Left            =   2400
      TabIndex        =   34
      Top             =   4455
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Intermediate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3360
      TabIndex        =   33
      Top             =   4080
      Width           =   1110
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   1
      Left            =   6060
      TabIndex        =   32
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Beginner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2040
      TabIndex        =   31
      Top             =   4043
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Packages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   30
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Expert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4680
      TabIndex        =   29
      Top             =   4080
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total Cost"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   5700
      TabIndex        =   28
      Top             =   3690
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Shipping/Handling"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   27
      Top             =   3083
      Width           =   1665
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Order Total"
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
      Left            =   540
      TabIndex        =   26
      Top             =   6810
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Delivery Total"
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
      Left            =   540
      TabIndex        =   25
      Top             =   6210
      Width           =   1260
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   4
      Left            =   6060
      TabIndex        =   24
      Top             =   6855
      Width           =   720
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   3
      Left            =   6060
      TabIndex        =   23
      Top             =   6240
      Width           =   720
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   2
      Left            =   6060
      TabIndex        =   22
      Top             =   5655
      Width           =   720
   End
   Begin VB.Label tot 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   0
      Left            =   6060
      TabIndex        =   21
      Top             =   4455
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   540
      TabIndex        =   20
      Top             =   3683
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Package C"
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
      Index           =   1
      Left            =   540
      TabIndex        =   19
      Top             =   5610
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Package B"
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
      Index           =   1
      Left            =   540
      TabIndex        =   18
      Top             =   5010
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Package A"
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
      Index           =   1
      Left            =   540
      TabIndex        =   17
      Top             =   4410
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Software Package C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   16
      Top             =   2483
      Width           =   1830
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Expert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5460
      TabIndex        =   15
      Top             =   1163
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Intermediate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4020
      TabIndex        =   14
      Top             =   1163
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Beginner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2580
      TabIndex        =   13
      Top             =   1163
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Software Package B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   12
      Top             =   2040
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Software Package A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   540
      TabIndex        =   10
      Top             =   1523
      Width           =   1830
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim onLoad As Integer

Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim x As Variant, y As Variant

Dim CustData As String
Dim xn(19) As Integer
Dim NPurch As Boolean
Dim NOrder As Boolean


Dim Sftwr As SftPkg
Private Type SftPkg
    PKA(0 To 2) As Integer
    PKB(0 To 2) As Integer
    PKC(0 To 2) As Integer
End Type

Private Sub Appenda_Click()
Open App.Path & "\custom1.dat" For Input As #1
    Do Until EOF(1)
    Input #1, xn(0), xn(1), xn(2), xn(3), xn(4), xn(5), xn(6), xn(7), xn(8), xn(9), xn(10), xn(11), xn(12), xn(13), xn(14), xn(15), xn(16), xn(17), xn(18), xn(19)
        If Not xn(0) = Combo1.List(Combo1.ListIndex) Then
        Open App.Path & "\custom2.dat" For Append As #2
            Write #2, xn(0), xn(1), xn(2), xn(3), xn(4), xn(5), xn(6), xn(7), xn(8), xn(9), xn(10), xn(11), xn(12), xn(13), xn(14), xn(15), xn(16), xn(17), xn(18), xn(19)
        Close #2
        Else
        Open App.Path & "\custom2.dat" For Append As #3
            Write #3, Combo1.List(Combo1.ListIndex), Val(QtySft(0).Caption), Val(QtySft(1).Caption), Val(QtySft(2).Caption), Val(QtySft(3).Caption), Val(QtySft(4).Caption), Val(QtySft(5).Caption), Val(QtySft(6).Caption), Val(QtySft(7).Caption), Val(QtySft(8).Caption), Sftwr.PKA(0), Sftwr.PKA(1), Sftwr.PKA(2), Sftwr.PKB(0), Sftwr.PKB(1), Sftwr.PKB(2), Sftwr.PKC(0), Sftwr.PKC(1), Sftwr.PKC(2), tot(3).Caption
        Close #3
        End If
'        CustData = "\custom2.dat"
    Loop

Close #1
End Sub

Private Sub Form_Load()
'IOA.Fins (App.Path & "\appth.dat")

LoadCustomer_datafiles

End Sub

Private Sub Combo1_Click()
For i = 0 To 8
Text1(i).Text = "0"
QtySft(i).Caption = "0"
Next i
Text1(9).Text = "0"

If NOrder = False Then
Open App.Path & CustData For Input As #1
    Do Until EOF(1)
        For i = 0 To 19
            Input #1, xn(i)
        Next i
        If Combo1.ListIndex + 100 = xn(0) Then
            For j = 1 To 9
                Text1(j - 1).Text = xn(j)
                QtySft(j - 1).Caption = xn(j)
            Next j
            Exit Do
        End If
    Loop
'If Not Combo1.List(Combo1.ListIndex) = xn(0) Then
'    For j = 0 To 8
'        Text1(j).Text = "0"
'        QtySft(j).Caption = 0
'    Next j
'    Text1(9).Text = "0"
'End If
Close #1

End If
NOrder = False
Text1(9).Text = Format("0", "Currency")
SubTotal
End Sub

Private Sub NewOrder_Click()
    For i = 0 To 9
    Text1(i).Text = "0"
    Next i
    For i = 0 To 8
    QtySft(i).Caption = "0"
    Next i
Purchase.Default = True
Combo1.AddItem 100 + Combo1.ListCount
Combo1.ListIndex = Combo1.ListCount - 1
NPurch = False
NOrder = True
'Else
'    For i = 0 To 9
'    Text1(i).Text = ""
'    Next i
'    For i = 0 To 8
'    QtySft(i).Caption = "0"
'    Next i
'    Purchase.Default = True
'End If
End Sub
Private Sub Exit_Click()
    End
End Sub

Private Sub Purchase_Click()
x = 0

Open App.Path & CustData For Input As #1
    Do Until EOF(1)
            Input #1, xn(0), xn(1), xn(2), xn(3), xn(4), xn(5), xn(6), xn(7), xn(8), xn(9), xn(10), xn(11), xn(12), xn(13), xn(14), xn(15), xn(16), xn(17), xn(18), xn(19)
        If Combo1.ListIndex + 100 = xn(0) Then
        Close #1
        Exit Sub
        End If
        Loop
        
        x = Val(Combo1.List(Combo1.ListIndex))
        Open App.Path & CustData For Append As #1
        Write #1, x, Val(QtySft(0).Caption), Val(QtySft(1).Caption), Val(QtySft(2).Caption), _
                 Val(QtySft(3).Caption), Val(QtySft(4).Caption), Val(QtySft(5).Caption), _
                 Val(QtySft(6).Caption), Val(QtySft(7).Caption), Val(QtySft(8).Caption), _
                 Sftwr.PKA(0), Sftwr.PKA(1), Sftwr.PKA(2), _
                 Sftwr.PKB(0), Sftwr.PKB(1), Sftwr.PKB(2), _
                 Sftwr.PKC(0), Sftwr.PKC(1), Sftwr.PKC(2), _
                 tot(3).Caption
        Close #1
            

End Sub

Private Sub Text1_GotFocus(Index As Integer)
If Combo1.ListCount = 0 Then
Combo1.AddItem 100 + Combo1.ListCount
Combo1.ListIndex = 1
End If

    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub
Private Sub Text1_LostFocus(Index As Integer)
For i = 0 To 8
    QtySft(i).Caption = Text1(i).Text
    If Text1(i).Text = "" Then
    QtySft(i).Caption = 0
    End If
    Next i
    
    If Index = 9 Then
        Text1(9).Text = Format(Text1(9).Text, "Currency")
    x = Text1(9).Text
        tot(3).Caption = Format(Text1(9).Text, "Currency")
    x = tot(3).Caption
    End If
    
SubTotal
End Sub
Private Sub SubTotal()
    For i = 0 To 2
        tot(i).Caption = "0"
    Next i
    For i = 0 To 2
        tot(0).Caption = Format(tot(0).Caption + Sftwr.PKA(i) * QtySft(i + 0).Caption, "Currency")
        tot(1).Caption = Format(tot(1).Caption + Sftwr.PKB(i) * QtySft(i + 3).Caption, "Currency")
        tot(2).Caption = Format(tot(2).Caption + Sftwr.PKC(i) * QtySft(i + 6).Caption, "Currency")
    Next i
        
        x = tot(3).Caption
        x = Text1(9).Text
        tot(3).Caption = Text1(9).Text



'    tot(0).Caption = Format(tot(0).Caption, "Currency")
'    x = tot(0).Caption
'    tot(1).Caption = Format(tot(1).Caption, "Currency")
'    tot(2).Caption = Format(tot(2).Caption, "Currency")
'    tot(3).Caption = Format(tot(3).Caption, "Currency")
'    x = tot(0).Caption
    x = 0
    For i = 0 To 3
        x = x + tot(i).Caption * 1
    
    Next i
    tot(4).Caption = Format(x, "Currency")




End Sub

Public Sub LoadCustomer_datafiles()
onLoad = 0
Preface:
Select Case onLoad
Case 0
    On Error GoTo OnloadError
    filename = "\appth.dat"
    Open App.Path & filename For Input As #1
    Input #1, CustData
    Close #1
    onLoad = 1
    GoTo Preface
Case 1
    For i = 0 To 8
    Text1(i).Text = "0"
    QtySft(i).Caption = 0
    tot(Int(i * 0.49)).Caption = 0
    Next i
    onLoad = 2
    GoTo Preface
Case 2
    j = 0
    x = 7
        Combo1.AddItem "Customer Order No"
        Fins (App.Path & CustData)
Case 3
Case 4
Case 5
Case 777
    Sftwr.PKA(0) = 50: Sftwr.PKA(1) = 150: Sftwr.PKA(2) = 450
    Sftwr.PKB(0) = 75: Sftwr.PKB(1) = 225: Sftwr.PKB(2) = 675
    Sftwr.PKC(0) = 100: Sftwr.PKC(1) = 300: Sftwr.PKC(2) = 900
End Select
'///////////////////////////////////////////////////////////////////////////
OnloadError:
Select Case Err.Number
Case 52
Case 53
    Select Case onLoad
    Case 0
    Close #1
    filename = "\appth.dat"
    dod.Fouts(0) = 0
    ifNeedFile (filename)
    onLoad = 0
    GoTo Preface
    End Select

Case 62
    Select Case onLoad
    Case 0
    Close #1
    msg62 = InputBox("Please enter the name of the purchase order database. Max 7 letters", "Create new database")
    msg62 = "\" & msg62
        Open App.Path & filename For Output As #5
        Write #5, msg62
        Close #5
'    errNumber.ErrorHandler
    GoTo Preface
    End Select
End Select
End Sub

Public Function ifNeedFile(filename As String)
    errNumber.ErrorHandler
Err.Clear
LoadCustomer_datafiles
'Resume
End Function
