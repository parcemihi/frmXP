VERSION 5.00
Begin VB.Form frmConvert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Numeric Format Convertor"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   5
         Format          =   ""
         HaveTrueFalseNull=   1
         TrueValue       =   "True"
         FalseValue      =   "False"
         NullValue       =   ""
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   7
      EndProperty
      Height          =   375
      Left            =   444
      TabIndex        =   4
      Top             =   1406
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3118
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1406
      Width           =   2055
   End
   Begin VB.CommandButton CmdBin 
      Caption         =   "&Bianary"
      Height          =   495
      Left            =   1562
      TabIndex        =   6
      Top             =   2149
      Width           =   1215
   End
   Begin VB.CommandButton CmdDec 
      Caption         =   "&Decimal"
      Height          =   495
      Left            =   4202
      TabIndex        =   8
      Top             =   2149
      Width           =   1215
   End
   Begin VB.CommandButton CmdHex 
      Caption         =   "&Hexidecimal"
      Default         =   -1  'True
      Height          =   495
      Left            =   242
      TabIndex        =   5
      Top             =   2149
      Width           =   1215
   End
   Begin VB.CommandButton CmdOct 
      Caption         =   "&Octal"
      Height          =   495
      Left            =   2882
      TabIndex        =   7
      Top             =   2149
      Width           =   1215
   End
   Begin VB.OptionButton Opt3Bin 
      Caption         =   "BI&N"
      Height          =   495
      Left            =   3118
      TabIndex        =   2
      Top             =   431
      Width           =   735
   End
   Begin VB.OptionButton Opt1Dec 
      Caption         =   "DE&C"
      Height          =   495
      Left            =   478
      TabIndex        =   0
      Top             =   431
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton Opt2Hex 
      Caption         =   "HE&X"
      Height          =   495
      Left            =   1798
      TabIndex        =   1
      Top             =   431
      Width           =   735
   End
   Begin VB.OptionButton Opt4Oct 
      Caption         =   "OC&T"
      Height          =   495
      Left            =   4438
      TabIndex        =   3
      Top             =   431
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Value Type"
      Height          =   855
      Left            =   238
      TabIndex        =   12
      Top             =   191
      Width           =   5175
   End
   Begin VB.Frame FrameAnswer 
      Caption         =   "Answer"
      Height          =   975
      Left            =   2878
      TabIndex        =   10
      Top             =   1046
      Width           =   2535
   End
   Begin VB.Frame FrameValue 
      Caption         =   "Enter Value"
      Height          =   975
      Left            =   238
      TabIndex        =   11
      Top             =   1046
      Width           =   2535
   End
   Begin VB.Menu MenuFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
      Begin VB.Menu MenuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "&Options"
      Begin VB.Menu menuDefault 
         Caption         =   "Set &Default"
         Begin VB.Menu menuHexidecimal 
            Caption         =   "&Hexidecimal"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuBianary 
            Caption         =   "&Bianary"
         End
         Begin VB.Menu menuOctal 
            Caption         =   "&Octal"
         End
         Begin VB.Menu menuDecimal 
            Caption         =   "&Decimal"
         End
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdDec_Click()
    If Opt1Dec = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        DEC2DEC (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt2Hex = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        HEX2DEC (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If

    If Opt3Bin = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        BIN2DEC (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt4Oct = True Then
        pass = Text1.Text
        If errorCheck = 1 Then
        clr
        OCT2DEC (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
End Sub
Private Sub CmdHex_Click()
    If Opt1Dec = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        DEC2HEX (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt2Hex = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        HEX2HEX (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If

    If Opt3Bin = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        BIN2HEX (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt4Oct = True Then
        pass = Text1.Text
        If errorCheck = 1 Then
        clr
        OCT2HEX (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
End Sub
Private Sub CmdBin_Click()
    If Opt1Dec = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        DEC2BIN (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt2Hex = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        HEX2BIN (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If

    If Opt3Bin = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        BIN2BIN (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt4Oct = True Then
        pass = Text1.Text
        If errorCheck = 1 Then
        clr
        OCT2BIN (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
End Sub
Private Sub CmdOct_Click()
    If Opt1Dec = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        DEC2OCT (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt2Hex = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        HEX2OCT (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If

    If Opt3Bin = True Then
    pass = Text1.Text
        If errorCheck = 1 Then
        clr
        BIN2OCT (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
    
    If Opt4Oct = True Then
        pass = Text1.Text
        If errorCheck = 1 Then
        clr
        OCT2OCT (Text1.Text)
        Text2.Text = pass
        selen
        Else
        selen
        End If
    End If
End Sub

Private Sub menuAbout_Click()
frmAbout.Show
End Sub

'Set Default for conversion to
Private Sub menuHexidecimal_Click()
    CmdHex.Default = True
    
    menuHexidecimal.Checked = True
    menuBianary.Checked = False
    menuOctal.Checked = False
    menuDecimal.Checked = False

End Sub

Private Sub menuBianary_Click()
    CmdBin.Default = True
    
    menuHexidecimal.Checked = False
    menuBianary.Checked = True
    menuOctal.Checked = False
    menuDecimal.Checked = False

End Sub

Private Sub menuOctal_Click()
    CmdOct.Default = True
    
    menuHexidecimal.Checked = False
    menuBianary.Checked = False
    menuOctal.Checked = True
    menuDecimal.Checked = False

End Sub
Private Sub menuDecimal_Click()
    
    CmdDec.Default = True
    
    menuHexidecimal.Checked = False
    menuBianary.Checked = False
    menuOctal.Checked = False
    menuDecimal.Checked = True

End Sub

'Miscellaneous Subs and Functions
Private Sub About_Click()
    frmAbout.Show
End Sub
Private Sub MenuExit_Click()
    End
End Sub
Public Sub clr()
    Text2.Text = ""
End Sub
Public Sub selen()
    Text1.SetFocus
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

