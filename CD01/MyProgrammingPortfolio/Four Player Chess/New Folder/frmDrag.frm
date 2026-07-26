VERSION 5.00
Begin VB.Form frmDrag 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Automatic Drag & Drop"
   ClientHeight    =   6360
   ClientLeft      =   1785
   ClientTop       =   1650
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7335
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   615
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H0000FFFF&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4440
      Width           =   615
   End
End
Attribute VB_Name = "frmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    Dim a As Integer
    
    'set all drag mode properties to Automatic
    For a = cmdButton.LBound To cmdButton.UBound
        cmdButton(a).DragMode = 1       'Automatic
    Next a
    
End Sub


Private Sub Form_DragDrop(Source As Control, x As Single, _
                            y As Single)
                            
    Dim w As Integer, h As Integer
    
    'Center control on mouse pointer
    w = x - Source.Width / 2
    h = y - Source.Height / 2
    
    'Move button to location where drop occurs
    Call Source.Move(w, h)
     
 End Sub

