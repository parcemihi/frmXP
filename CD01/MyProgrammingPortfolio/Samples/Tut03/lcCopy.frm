VERSION 5.00
Begin VB.Form frmCopyright 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Interlocking Software Company"
   ClientHeight    =   5115
   ClientLeft      =   1320
   ClientTop       =   1260
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrCopy 
      Interval        =   4000
      Left            =   360
      Top             =   600
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click our logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      MousePointer    =   13  'Arrow and Hourglass
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         Caption         =   "Written By Keith"
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
         Left            =   4200
         TabIndex        =   2
         Top             =   1125
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         Caption         =   "Copyright Year 2000"
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
         Left            =   4200
         TabIndex        =   1
         Top             =   1560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image imgLogo 
         Height          =   3570
         Left            =   240
         Picture         =   "lcCopy.frx":0000
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmCopyright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmCopyright.Top = (Screen.Height - frmCopyright.Height) / 2
    frmCopyright.Left = (Screen.Width - frmCopyright.Width) / 2
    Load frmInput
    Load frmOrder
    
End Sub


Private Sub imgLogo_Click()
    lblAuthor.Visible = True
    lblYear.Visible = True
    
End Sub


Private Sub tmrCopy_Timer()
    frmInput.Show
    Unload frmCopyright

End Sub
