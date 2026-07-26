VERSION 5.00
Begin VB.Form vhs800 
   ClientHeight    =   9060
   ClientLeft      =   165
   ClientTop       =   330
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame libCommand 
      Caption         =   "Frame1"
      Height          =   8535
      Left            =   248
      TabIndex        =   0
      Top             =   240
      Width           =   12975
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3360
         TabIndex        =   52
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3360
         TabIndex        =   51
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   3360
         TabIndex        =   50
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   3360
         TabIndex        =   49
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   3360
         TabIndex        =   48
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   3360
         TabIndex        =   47
         Top             =   3240
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   3360
         TabIndex        =   46
         Top             =   3720
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   0
         Left            =   3000
         Picture         =   "SCR800vhs.frx":0000
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   45
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   1
         Left            =   3000
         Picture         =   "SCR800vhs.frx":068E
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   44
         Top             =   1320
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   2
         Left            =   3000
         Picture         =   "SCR800vhs.frx":0D1C
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   43
         Top             =   1800
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   3
         Left            =   3000
         Picture         =   "SCR800vhs.frx":13AA
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   42
         Top             =   2280
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   4
         Left            =   3000
         Picture         =   "SCR800vhs.frx":1A38
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   41
         Top             =   2760
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   5
         Left            =   3000
         Picture         =   "SCR800vhs.frx":20C6
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   40
         Top             =   3240
         Width           =   255
      End
      Begin VB.PictureBox LEDcounter 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   6
         Left            =   3000
         Picture         =   "SCR800vhs.frx":2754
         ScaleHeight     =   465
         ScaleWidth      =   255
         TabIndex        =   39
         Top             =   3720
         Width           =   255
      End
      Begin VB.TextBox vhs 
         Height          =   1110
         Index           =   1
         Left            =   6840
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   795
         Width           =   5895
      End
      Begin VB.TextBox vhs 
         Height          =   345
         Index           =   2
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   36
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox vhs 
         Height          =   360
         Index           =   3
         Left            =   10725
         TabIndex        =   35
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox vhs 
         Height          =   360
         Index           =   4
         Left            =   11880
         TabIndex        =   34
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox vhs 
         Height          =   360
         Index           =   5
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton addDBR 
         Caption         =   "&Add Record"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6840
         TabIndex        =   32
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton deleteDBR 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8040
         TabIndex        =   31
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton updateDB 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9240
         TabIndex        =   30
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton navigation 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   9240
         TabIndex        =   29
         Top             =   2520
         Width           =   735
      End
      Begin VB.ComboBox SearchFilters 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         ItemData        =   "SCR800vhs.frx":2DE2
         Left            =   240
         List            =   "SCR800vhs.frx":2DE4
         TabIndex        =   28
         Text            =   " - All Courses"
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton navigation 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   10080
         TabIndex        =   27
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton navigation 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   11040
         TabIndex        =   26
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton navigation 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   12000
         TabIndex        =   25
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox vhs 
         Height          =   360
         Index           =   0
         Left            =   11280
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   315
         Width           =   1455
      End
      Begin VB.CommandButton ExitProgram 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   11640
         TabIndex        =   23
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Search"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   22
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox vhs 
         Height          =   360
         Index           =   6
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   21
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   360
         Left            =   960
         TabIndex        =   20
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox SearchFilters 
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         ItemData        =   "SCR800vhs.frx":2DE6
         Left            =   240
         List            =   "SCR800vhs.frx":2DE8
         TabIndex        =   19
         Text            =   " - All Subjects"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox SearchFilters 
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         ItemData        =   "SCR800vhs.frx":2DEA
         Left            =   240
         List            =   "SCR800vhs.frx":2DEC
         TabIndex        =   18
         Text            =   " - All Topic "
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox SortOrder 
         Enabled         =   0   'False
         Height          =   360
         Left            =   240
         TabIndex        =   17
         Text            =   " - Sort Orders "
         Top             =   2280
         Width           =   2535
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   4920
         TabIndex        =   16
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   4920
         TabIndex        =   15
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   4920
         TabIndex        =   14
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   4920
         TabIndex        =   13
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   4920
         TabIndex        =   12
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   4920
         TabIndex        =   11
         Top             =   3240
         Width           =   255
      End
      Begin VB.CommandButton LEDPlusMinus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   13
         Left            =   4920
         TabIndex        =   10
         Top             =   3720
         Width           =   255
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "Location"
         Height          =   495
         Index           =   6
         Left            =   3600
         TabIndex        =   9
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "Quantity"
         Height          =   495
         Index           =   5
         Left            =   3600
         TabIndex        =   8
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "Time"
         Height          =   495
         Index           =   4
         Left            =   3600
         TabIndex        =   7
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "Format"
         Height          =   495
         Index           =   3
         Left            =   3600
         TabIndex        =   6
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "Copyright"
         Height          =   495
         Index           =   2
         Left            =   3600
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "Title"
         Height          =   495
         Index           =   1
         Left            =   3600
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton NaviSortBy 
         Caption         =   "ID Number"
         Height          =   495
         Index           =   0
         Left            =   3600
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.ListBox CurrentSort 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   240
         TabIndex        =   2
         Top             =   4560
         Width           =   12495
      End
      Begin VB.CommandButton interActiveListMode 
         Caption         =   "Inter Active  List Mode  &OFF"
         Height          =   975
         Left            =   5280
         TabIndex        =   1
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CommandButton editModeButton 
         Caption         =   "Edit &Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5040
         TabIndex        =   38
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Image EditMode 
         Height          =   975
         Index           =   1
         Left            =   10440
         Picture         =   "SCR800vhs.frx":2DEE
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "Title"
         Height          =   240
         Index           =   1
         Left            =   6375
         TabIndex        =   60
         Top             =   795
         Width           =   390
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "Copyright Date"
         Height          =   240
         Index           =   2
         Left            =   5430
         TabIndex        =   59
         Top             =   2115
         Width           =   1335
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "Format / Time"
         Height          =   240
         Index           =   3
         Left            =   9360
         TabIndex        =   58
         Top             =   2160
         Width           =   1245
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   240
         Index           =   4
         Left            =   6045
         TabIndex        =   57
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   240
         Index           =   5
         Left            =   6000
         TabIndex        =   56
         Top             =   2835
         Width           =   765
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "ID Number"
         Height          =   240
         Index           =   0
         Left            =   10245
         TabIndex        =   55
         Top             =   315
         Width           =   960
      End
      Begin VB.Label VHS_Field 
         AutoSize        =   -1  'True
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   240
         Index           =   9
         Left            =   240
         TabIndex        =   54
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search Records:"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   53
         Top             =   480
         Width           =   2025
      End
      Begin VB.Image EditMode 
         Height          =   975
         Index           =   0
         Left            =   10440
         Picture         =   "SCR800vhs.frx":660C
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   0
         Left            =   2520
         Picture         =   "SCR800vhs.frx":9E2A
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   1
         Left            =   2760
         Picture         =   "SCR800vhs.frx":A4B8
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   2
         Left            =   3000
         Picture         =   "SCR800vhs.frx":AB46
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   3
         Left            =   3240
         Picture         =   "SCR800vhs.frx":B1D4
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   4
         Left            =   3480
         Picture         =   "SCR800vhs.frx":B862
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   5
         Left            =   3720
         Picture         =   "SCR800vhs.frx":BEF0
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   6
         Left            =   3960
         Picture         =   "SCR800vhs.frx":C57E
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   7
         Left            =   4200
         Picture         =   "SCR800vhs.frx":CC0C
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   8
         Left            =   4440
         Picture         =   "SCR800vhs.frx":D29A
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   9
         Left            =   4680
         Picture         =   "SCR800vhs.frx":D928
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image LEDAplhaNum 
         Height          =   465
         Index           =   10
         Left            =   4920
         Picture         =   "SCR800vhs.frx":DFB6
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Add New Record"
         Index           =   0
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Delete Record"
         Index           =   1
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditRecord 
         Caption         =   "&Update Database"
         Index           =   3
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "&Edit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuNavi 
         Caption         =   "&Navigation"
         Begin VB.Menu mnuNaviCon 
            Caption         =   "&First       "
            Index           =   0
            Shortcut        =   ^H
         End
         Begin VB.Menu mnuNaviCon 
            Caption         =   "&Back"
            Index           =   1
            Shortcut        =   ^J
         End
         Begin VB.Menu mnuNaviCon 
            Caption         =   "&Next"
            Index           =   2
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuNaviCon 
            Caption         =   "&Last"
            Index           =   3
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort By"
         Begin VB.Menu mnuSortBy 
            Caption         =   "&ID#"
            Index           =   0
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Title"
            Checked         =   -1  'True
            Index           =   1
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Copyright Date"
            Index           =   2
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Format"
            Index           =   3
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Running Time"
            Index           =   4
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Location"
            Index           =   5
         End
         Begin VB.Menu mnuSortBy 
            Caption         =   "&Subject"
            Index           =   6
         End
      End
   End
End
Attribute VB_Name = "vhs800"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SortON_OFF As Boolean
Dim InterActiveMode As Boolean
Dim LEDPriorities(0 To 6) As Integer
Dim EditDone As Boolean
Dim EditRecord As Boolean
Dim TotalVideos As Integer
Dim RNDBPointer As Integer
Dim VideoDB() As RNVHSDB
Private Type RNVHSDB
    ID As String
    Quantity As String
    Title As String
    Copyright As String
    MediaFormat As String
    RunningTime As String
    Location As String
End Type
Dim CategoryDB(9) As CATEGORY_SUB
Private Type CATEGORY_SUB
    CategoryID As String
    ID As String
End Type
Dim TopicDB(9) As TOPIC_SUB
Private Type TOPIC_SUB
    TopicID As String
    ID As String
End Type
Dim SubjectDB(9) As SUBJECT_DEC
Private Type SUBJECT_DEC
    SubjectID As String
    ID As String
End Type

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Dim Subject As String
    ReDim VideoDB(0 To 1000)
    RNDBPointer = 0
    TotalVideos = 0
    SortON_OFF = False
    
    SearchFilters(0).AddItem " - All Courses"
    SearchFilters(1).AddItem " - All Subjects"
    SearchFilters(2).AddItem " - All Topics"
    SortOrder.AddItem " - Sort Order"
    
    Open App.Path & "\catalog.dat" For Input As #1
        Do Until EOF(1)
            Input #1, Subject
            'Combo1.AddItem Subject
        Loop
    Close #1
TotalVideos = -1
    Open App.Path & "\rnvhsdb.dat" For Input As #1
        Do Until EOF(1)
            TotalVideos = TotalVideos + 1
            Input #1, VideoDB(TotalVideos).ID, _
                      VideoDB(TotalVideos).Title, _
                      VideoDB(TotalVideos).Copyright, _
                      VideoDB(TotalVideos).MediaFormat, _
                      VideoDB(TotalVideos).RunningTime, _
                      VideoDB(TotalVideos).Quantity, _
                      VideoDB(TotalVideos).Location
            CurrentSort_UpdateList (1)
        Loop
    Close #1
    TotalVideos = TotalVideos - 1
    ENTER_VISUALS (0)
    navigation_Click (0)
    Exit Sub
ErrorHandler:
    Close #1
    Exit Sub
End Sub

Private Function ENTER_VISUALS(Personal_Settings As Integer)
Dim i As Integer
For i = 0 To 6
    LEDcounter(i).Picture = LEDAplhaNum(10)
    LEDPriorities(i) = 1 '1
    LEDPlusMinus_Click (i)
Next i
SortON_OFF = True
End Function

Private Sub Form_Resize()
If Screen.Width = 11520 Then
End If
If Screen.Height = 9000 Then
End If

If Screen.Width = 12000 Then
End If
If Screen.Height = 9000 Then
End If

End Sub

Private Sub interActiveListMode_Click()
'CurrentSort_Inter (True)

If InterActiveMode = True Then
    InterActiveMode = False
    CurrentSort_Inter (InterActiveMode)
    interActiveListMode.Caption = "Inter Active List Mode" & Chr(10) & "&OFF"

Else
    interActiveListMode.Caption = "Inter Active List Mode" & Chr(10) & "&ON "
    InterActiveMode = True
    CurrentSort_Inter (InterActiveMode)

End If




End Sub

Private Sub LEDPlusMinus_Click(Index As Integer)
Dim i As Integer
'//ON Statement

Select Case Index
    Case 0 To 6
        If Not LEDPriorities(Index) = 0 Then
            LEDPriorities(Index) = LEDPriorities(Index) - 1
            LEDcounter(Index).Picture = LEDAplhaNum(LEDPriorities(Index))
            If LEDPriorities(Index) = 0 Then
                LEDPlusMinus(Index).Enabled = False
                LEDcounter(Index).Picture = LEDAplhaNum(10)
                NaviSortBy(Index).Font.Bold = True
            End If
            LEDSort SortON_OFF, Index
        End If
    Case 7 To 13
        If Not LEDPriorities(Index - 7) = 7 Then
            LEDSort SortON_OFF, Index
            
            
            LEDPlusMinus(Index - 7).Enabled = True
'            LEDPriorities(Index - 7) = LEDPriorities(Index - 7) + 1
'            LEDcounter(Index - 7).Picture = LEDAplhaNum(LEDPriorities(Index - 7))
            NaviSortBy(Index - 7).Font.Bold = False
'            LEDSort SortON_OFF, Index
        End If
End Select


End Sub
Private Function LEDSort(SortON_OFF As Boolean, Index_SortCode As Integer)
Dim LEDPrioritySorted As Boolean
Dim x As Integer
Dim i As Integer

If SortON_OFF = True Then
    Select Case Index_SortCode
        Case 0 To 6
            If NaviSortBy(Index_SortCode).Font.Bold = True Then
            Else
            End If
        
        Case 7 To 13
'            For i = 0 To 6
'                LEDPriorities(Index_SortCode - 7) = LEDPriorities(Index_SortCode - 7) + 1
'            Next i
            For i = 0 To 6
'                If LEDPriorities(i) <> 0 And LEDPriorities(i) < HighestNumber Then
'                   If i <> Index_SortCode - 7 Then
                    If i = Index_SortCode - 7 Then
                        LEDcounter(i).Picture = LEDAplhaNum(LEDPriorities(i))
                    Else
                        
                        '// SET Pivot for priority
                        If LEDPriorities(i) <= LEDPriorities(Index_SortCode - 7) Then
                            LEDPriorities(i) = LEDPriorities(i) + 1
                            LEDcounter(i).Picture = LEDAplhaNum(LEDPriorities(i))
                            
                            
                        End If
                        
                    End If
'                End If
            Next i
            
            NaviSortBy(Index_SortCode - 7).Font.Bold = False
    End Select
'// Enter the code for finding the next available priority slot
Else
    x = x
End If

End Function


Private Sub editModeButton_Click()
Dim passthis
If EditMode(0).Visible = False Then
    passthis = 1
Else
    passthis = 0
End If
    EditMode_Click (passthis)
End Sub

Private Sub EditMode_Click(Index As Integer)
Select Case Index
    Case 0
    EditMode(0).Visible = False
    EditMode(1).Visible = True
    EditRecord = True
    
    
    
    Case 1
    EditMode(0).Visible = True
    EditMode(1).Visible = False
    EditRecord = False
    RefreshList
    mnuSortBy_Click (0)
    Case Else
    
    
    
End Select

End Sub

Private Sub RefreshList()
Dim i As Integer
CurrentSort.Clear
For i = 0 To TotalVideos
    
    
    
Next i
End Sub




Private Sub LEDAplhaNum_Click(Index As Integer)
0 = 0
' 1 = 1
' 2 = 2
' 3 = 3
' 4 = 4
' 5 = 5
' 6 = 6
' 7 = 7
' 8 = 8
' 9 = 9
'10 = blank
'11 =
'12 =
'13 =
'14 =
'15 =
'16 =
'17 =
'18 =
'19 =
'20 =
'21 =
'22 =
'23 =
'24 =
'25 =
'26 =
'27 =
End Sub

        
'        i = LEDPriorities(Index)
'        LEDPriorities(Index) = LEDPriorities(Index + 1)



'//End Statement
'For i = 0 To 6
'    LEDcounter(i).Picture = LEDAplhaNum(LEDPriorities(i + 1))
'Next i

'// Filter Features
Private Sub mnuSortBy_Click(Index As Integer)
Dim i As Integer
For i = 0 To 5
    If i = Index Then
        mnuSortBy(i).Checked = True
    Else
        mnuSortBy(i).Checked = False
    End If
Next i








End Sub
                                                    '// Navigation
Private Sub mnuNaviCon_Click(Index As Integer)
    navigation_Click (Index)
End Sub
Private Sub navigation_Click(Index As Integer)
EditRecord = False
Select Case Index
    Case 0
        RNDBPointer = 0
    Case 1
        If Not RNDBPointer = 0 Then
            RNDBPointer = RNDBPointer - 1
        Else
            RNDBPointer = TotalVideos
        End If
    Case 2
        If Not RNDBPointer = TotalVideos Then
            RNDBPointer = RNDBPointer + 1
        Else
            RNDBPointer = 0
        End If
    Case 3
        RNDBPointer = TotalVideos
End Select

'// Insert Sort procedure

Refresh_vhs

EditDone = True
'// Change Light on Edit Mode Button
    EditMode(0).Visible = True
    EditMode(1).Visible = False
'// Resort
End Sub
Private Sub CurrentSort_Inter(ActiveMode_IO As Boolean)
If ActiveMode_IO = True Then
    If CurrentSort.ListIndex > -1 Then
        RNDBPointer = CurrentSort.ListIndex
        vhs(0).Text = VideoDB(RNDBPointer).ID
        vhs(1).Text = VideoDB(RNDBPointer).Title
        vhs(2).Text = VideoDB(RNDBPointer).Copyright
        vhs(3).Text = VideoDB(RNDBPointer).MediaFormat
        vhs(4).Text = VideoDB(RNDBPointer).RunningTime
        vhs(5).Text = VideoDB(RNDBPointer).Quantity
        vhs(6).Text = VideoDB(RNDBPointer).Location
    End If
End If
End Sub
Private Sub CurrentSort_Click()

If InterActiveMode = True Then
    RNDBPointer = CurrentSort.ListIndex
    vhs(0).Text = VideoDB(RNDBPointer).ID
    vhs(1).Text = VideoDB(RNDBPointer).Title
    vhs(2).Text = VideoDB(RNDBPointer).Copyright
    vhs(3).Text = VideoDB(RNDBPointer).MediaFormat
    vhs(4).Text = VideoDB(RNDBPointer).RunningTime
    vhs(5).Text = VideoDB(RNDBPointer).Quantity
    vhs(6).Text = VideoDB(RNDBPointer).Location
End If
End Sub


'// Database Record functions
Private Sub addDBR_Click()
    Dim i As Integer
    TotalVideos = TotalVideos + 1
    RNDBPointer = TotalVideos
    For i = 0 To 6
        vhs(i).Text = ""
    Next i
    EditDone = False
    EditMode_Click (0)
'    If EditDone = True Then
'        interActiveListMode_Click
        
        
        
'    End If
    
    
    
End Sub

Private Sub updateFields_Click()
If EditRecord = True Then
    VideoDB(RNDBPointer).ID = vhs(0).Text
    VideoDB(RNDBPointer).Title = vhs(1).Text
    VideoDB(RNDBPointer).Copyright = vhs(2).Text
    VideoDB(RNDBPointer).MediaFormat = vhs(3).Text
    VideoDB(RNDBPointer).RunningTime = vhs(4).Text
    VideoDB(RNDBPointer).Quantity = vhs(5).Text
    VideoDB(RNDBPointer).Location = vhs(6).Text
End If
End Sub

Private Sub deleteDBR_Click()
    Dim msg As Integer
    Dim i As Integer
    msg = MsgBox("               Delete Record?", vbOKCancel, "Confirm Record Delete.")
    If msg = 1 Then ' OK
        For i = RNDBPointer To TotalVideos
            VideoDB(i).ID = VideoDB(i + 1).ID
            VideoDB(i).Title = VideoDB(i + 1).Title
            VideoDB(i).Copyright = VideoDB(i + 1).Copyright
            VideoDB(i).MediaFormat = VideoDB(i + 1).MediaFormat
            VideoDB(i).RunningTime = VideoDB(i + 1).RunningTime
            VideoDB(i).Quantity = VideoDB(i + 1).Quantity
            VideoDB(i).Location = VideoDB(i + 1).Location
        Next i
        TotalVideos = TotalVideos - 1
        Refresh_vhs
    End If
End Sub
                                                    '// Text Box Functions
Private Sub Refresh_vhs()
    vhs(0).Text = VideoDB(RNDBPointer).ID: vhs(1).Text = VideoDB(RNDBPointer).Title: vhs(2).Text = VideoDB(RNDBPointer).Copyright: vhs(3).Text = VideoDB(RNDBPointer).MediaFormat: vhs(4).Text = VideoDB(RNDBPointer).RunningTime: vhs(5).Text = VideoDB(RNDBPointer).Quantity: vhs(6).Text = VideoDB(RNDBPointer).Location
End Sub

Private Sub NaviSortBy_Click(Index As Integer)

Dim LED As Integer

If NaviSortBy(Index).Font.Bold = True Then
    NaviSortBy(Index).Font.Bold = False
    LEDPlusMinus(Index).Enabled = True
'//    UPDATE Priority
'    LEDSort (Index)
    LEDPlusMinus_Click (Index)
'    LEDSort (Index)
Else
    NaviSortBy(Index).Font.Bold = True
    LEDPlusMinus(Index).Enabled = False
    LEDcounter(Index).Picture = LEDAplhaNum(10)
End If

End Sub



Private Sub vhs_GotFocus(Index As Integer)
    vhs(Index).SelStart = 0
    vhs(Index).SelLength = Len(vhs(Index))
End Sub
Private Sub vhs_LostFocus(Index As Integer)
    updateFields_Click
End Sub

Private Sub ExitProgram_Click()
    End
End Sub
Private Sub mnuExit_Click()
    End
End Sub
Private Function CurrentSort_UpdateList(Include_Item As Integer)

Dim ID As String * 4, Title As String * 50, Copyright As String * 4, MediaFormat As String * 20, RunningTime As String * 4, Quantity As String * 3, Location As String * 6

ID = VideoDB(TotalVideos).ID
Title = VideoDB(TotalVideos).Title
Copyright = VideoDB(TotalVideos).Copyright
MediaFormat = VideoDB(TotalVideos).MediaFormat
RunningTime = VideoDB(TotalVideos).RunningTime
Quantity = VideoDB(TotalVideos).Quantity
Location = VideoDB(TotalVideos).Location

Dim LenOk As Boolean


If Len(VideoDB(TotalVideos).Title) > 46 Then
    CurrentSort.AddItem ID & Mid(Title, 1, 46) & "... " & Copyright & MediaFormat & RunningTime & Quantity & Location
Else
    CurrentSort.AddItem ID & Title & Copyright & MediaFormat & RunningTime & Quantity & Location
End If


End Function



Private Sub updateDB_Click()
    'Dim i As Integer
    'Open App.Path & "\rnvhsdb.ful" For Output As #1
    '    For i = 0 To TotalVideos
    '        Write #1, i, VideoDB(i).Title, VideoDB(i).Copyright, VideoDB(i).MediaFormat, VideoDB(i).RunningTime, VideoDB(i).Quantity, VideoDB(i).Location
    '    Next i
    'Close #1
End Sub


'If EditRecord = True Then
'    Select Case Index
'        Case 0
'        VideoDB(RNDBPointer).ID = vhs(Index).Text
'        Case 1
'        VideoDB(RNDBPointer).Title = vhs(Index).Text
'        Case 2
'        VideoDB(RNDBPointer).Copyright = vhs(Index).Text
'        Case 3
'        VideoDB(RNDBPointer).MediaFormat = vhs(Index).Text
'        Case 4
'        VideoDB(RNDBPointer).RunningTime = vhs(Index).Text
'        Case 5
'        VideoDB(RNDBPointer).Quantity = vhs(Index).Text
'        Case 6
'        VideoDB(RNDBPointer).Location = vhs(Index).Text
'    End Select
'End If

