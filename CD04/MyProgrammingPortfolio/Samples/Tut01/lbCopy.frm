VERSION 5.00
Begin VB.Form frmCopyright 
   Caption         =   "Interlocking Software"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
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
      TabIndex        =   0
      Top             =   1005
      Width           =   2745
   End
End
Attribute VB_Name = "frmCopyright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
