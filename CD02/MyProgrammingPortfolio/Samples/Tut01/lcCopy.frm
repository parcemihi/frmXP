VERSION 5.00
Begin VB.Form frmCopyright 
   BorderStyle     =   0  'None
   Caption         =   "Interlocking Software"
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Copyright.Scroller Scroller1 
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   1085
      BackColor       =   255
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Interlocking Software"
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
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   495
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         Caption         =   "Written By Keith E. Geiser"
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
         TabIndex        =   3
         Top             =   1005
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         Caption         =   "Copyright Year 2001"
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
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Image imgLogo 
         Height          =   3570
         Left            =   240
         Picture         =   "lcCopy.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3675
      End
   End
End
Attribute VB_Name = "frmCopyright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload frmCopyright
    End
    
End Sub

Private Sub imgLogo_Click()
lblAuthor.Visible = True
lblYear.Visible = True

End Sub

