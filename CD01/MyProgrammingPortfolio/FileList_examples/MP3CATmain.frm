VERSION 5.00
Begin VB.Form MP3CATmain 
   Caption         =   "Form1"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   5280
      TabIndex        =   4
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   5280
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Size"
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
      Index           =   3
      Left            =   4350
      TabIndex        =   9
      Top             =   2565
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Song Title"
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
      Index           =   2
      Left            =   4200
      TabIndex        =   8
      Top             =   2205
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Album"
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
      Left            =   4545
      TabIndex        =   7
      Top             =   1845
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Artist"
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
      Index           =   0
      Left            =   4680
      TabIndex        =   6
      Top             =   1485
      Width           =   435
   End
End
Attribute VB_Name = "MP3CATmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()



End Sub
