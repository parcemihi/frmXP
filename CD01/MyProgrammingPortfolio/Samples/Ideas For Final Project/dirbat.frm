VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   3480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1296
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4605
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Name && Location"
      TabPicture(0)   =   "dirbat.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Date"
      TabPicture(1)   =   "dirbat.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Advanced"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu SendTo 
         Caption         =   "Send &To"
         WindowList      =   -1  'True
         Begin VB.Menu Floppy 
            Caption         =   "3½ Floppy (A)"
         End
         Begin VB.Menu Desktop 
            Caption         =   "Desktop (Create Shortcut)"
         End
      End
      Begin VB.Menu CreateShortcut 
         Caption         =   "Create &Shortcut"
      End
      Begin VB.Menu Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu Rename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu Properties 
         Caption         =   "P&roperties"
      End
      Begin VB.Menu OpenFolder 
         Caption         =   "Open Conta&ining Folder"
      End
      Begin VB.Menu SaveSearch 
         Caption         =   "S&ave Search"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu View 
      Caption         =   "&View"
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
