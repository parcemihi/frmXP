VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3330
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2298.425
   ScaleMode       =   0  'User
   ScaleWidth      =   3648.216
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picapicon 
      Height          =   795
      Left            =   2760
      Picture         =   "frmAbout.frx":0ECA
      ScaleHeight     =   735
      ScaleWidth      =   750
      TabIndex        =   6
      Top             =   240
      Width           =   810
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   2880
      Width           =   1260
   End
   Begin VB.Label lblmadeby 
      Caption         =   "Produced by: Mr T"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblnotice 
      Caption         =   "Caution:     At the moment this can't cope with values of over 255. This and binary calculations will be coming up at some point."
      Height          =   615
      Left            =   180
      TabIndex        =   4
      Top             =   2040
      Width           =   3435
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description:     This is an app for converting denary, binary, hex and ascii into one another."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   210
      TabIndex        =   1
      Top             =   1365
      Width           =   3405
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1965
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    
' when clicked the main form is enabled and setfocus to
' then this form is unloaded
    
    frmdentobin.Enabled = True
    frmdentobin.SetFocus
    Unload Me
End Sub
Private Sub Form_Load()

' on the load of this form the caption is set, the version is calculated and the title is generated

    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub
