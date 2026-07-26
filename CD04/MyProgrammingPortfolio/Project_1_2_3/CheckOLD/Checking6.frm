VERSION 5.00
Begin VB.Form checking 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton newChkDep 
      Caption         =   "New Credit"
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   26
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton newChkDep 
      Caption         =   "New Debit"
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   25
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "Last"
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   24
      Top             =   2430
      Width           =   615
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "Next"
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   23
      Top             =   2430
      Width           =   615
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "Back"
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   22
      Top             =   2430
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "&Deposit"
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "&Check"
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   5535
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cleared"
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtCheckEdit 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton NaviCheck 
      Caption         =   "First"
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   2430
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   21
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "/"
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
      Left            =   2010
      TabIndex        =   20
      Top             =   1935
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "/"
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
      Left            =   1410
      TabIndex        =   19
      Top             =   1935
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   17
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Credits"
      Height          =   195
      Index           =   2
      Left            =   3945
      TabIndex        =   16
      Top             =   1140
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Debits"
      Height          =   195
      Index           =   1
      Left            =   3975
      TabIndex        =   15
      Top             =   780
      Width           =   450
   End
   Begin VB.Label totals 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label totals 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   13
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Balance"
      Height          =   195
      Index           =   0
      Left            =   3840
      TabIndex        =   12
      Top             =   420
      Width           =   585
   End
   Begin VB.Label totals 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "checking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

