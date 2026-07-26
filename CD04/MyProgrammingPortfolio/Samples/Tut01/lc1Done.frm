VERSION 5.00
Begin VB.Form frmAnimal 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Peoria Animal Shelter"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Click the logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   112
      TabIndex        =   0
      Top             =   120
      Width           =   5760
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Image imgLogo 
         Height          =   1890
         Left            =   720
         Picture         =   "lc1Done.frx":0000
         Stretch         =   -1  'True
         Top             =   1125
         Width           =   4245
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Welcome to the Peoria Animal Shelter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   4560
      End
   End
End
Attribute VB_Name = "frmAnimal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
    
End Sub

Private Sub imgLogo_Click()
lblTitle.Visible = True

End Sub
