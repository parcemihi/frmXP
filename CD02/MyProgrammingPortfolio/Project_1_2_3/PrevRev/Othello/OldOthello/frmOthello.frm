VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Keith's Othello"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command100 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command99 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command98 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command97 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command96 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command95 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command94 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command93 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command92 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command91 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton Command90 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command89 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command88 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command87 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command86 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command85 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command84 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command83 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command82 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command81 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton Command80 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command79 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command78 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command77 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   76
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command76 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command75 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command74 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command73 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command72 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command71 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command70 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command69 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command68 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command67 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command66 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command65 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command64 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command63 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command62 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command61 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton Command60 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command59 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command58 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command57 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command56 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command55 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command54 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command53 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command52 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command51 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command50 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command49 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command48 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command47 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command46 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command45 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command44 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command43 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command42 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command41 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton Command40 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command39 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command38 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command37 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command36 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command35 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command34 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command33 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command32 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command31 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command30 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command29 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command28 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command26 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command24 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command23 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command22 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command21 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command20 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command19 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command18 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command17 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgBlank 
      Height          =   735
      Left            =   3120
      Picture         =   "frmOthello.frx":0000
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image imgWhite 
      Height          =   735
      Left            =   3120
      Picture         =   "frmOthello.frx":037F
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image imgBlack 
      Height          =   735
      Left            =   3120
      Picture         =   "frmOthello.frx":075D
      Top             =   3480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Command1.Picture = imgBlank Then
    Command1.Picture = imgWhite
    ElseIf Command1.Picture = imgWhite Then
    Command1.Picture = imgBlack
    Else
    Command1.Picture = imgBlank
    End If
End Sub

Private Sub Command2_Click()
    If Command2.Picture = imgBlank Then
    Command2.Picture = imgWhite
    ElseIf Command2.Picture = imgWhite Then
    Command2.Picture = imgBlack
    Else
    Command2.Picture = imgBlank
    End If

End Sub

Private Sub Command3_Click()
    If Command3.Picture = imgBlank Then
    Command3.Picture = imgWhite
    ElseIf Command3.Picture = imgWhite Then
    Command3.Picture = imgBlack
    Else
    Command3.Picture = imgBlank
    End If

End Sub

Private Sub Command4_Click()
    If Command4.Picture = imgBlank Then
    Command4.Picture = imgWhite
    ElseIf Command4.Picture = imgWhite Then
    Command4.Picture = imgBlack
    Else
    Command4.Picture = imgBlank
    End If

End Sub

Private Sub Command5_Click()
    If Command5.Picture = imgBlank Then
    Command5.Picture = imgWhite
    ElseIf Command5.Picture = imgWhite Then
    Command5.Picture = imgBlack
    Else
    Command5.Picture = imgBlank
    End If

End Sub

Private Sub Command6_Click()
    If Command6.Picture = imgBlank Then
    Command6.Picture = imgWhite
    ElseIf Command6.Picture = imgWhite Then
    Command6.Picture = imgBlack
    Else
    Command6.Picture = imgBlank
    End If

End Sub

Private Sub Command7_Click()
    If Command7.Picture = imgBlank Then
    Command7.Picture = imgWhite
    ElseIf Command7.Picture = imgWhite Then
    Command7.Picture = imgBlack
    Else
    Command7.Picture = imgBlank
    End If

End Sub

Private Sub Command8_Click()
    If Command8.Picture = imgBlank Then
    Command8.Picture = imgWhite
    ElseIf Command8.Picture = imgWhite Then
    Command8.Picture = imgBlack
    Else
    Command8.Picture = imgBlank
    End If

End Sub

Private Sub Command9_Click()
    If Command9.Picture = imgBlank Then
    Command9.Picture = imgWhite
    ElseIf Command9.Picture = imgWhite Then
    Command9.Picture = imgBlack
    Else
    Command9.Picture = imgBlank
    End If

End Sub

Private Sub Command10_Click()
    If Command10.Picture = imgBlank Then
    Command10.Picture = imgWhite
    ElseIf Command10.Picture = imgWhite Then
    Command10.Picture = imgBlack
    Else
    Command10.Picture = imgBlank
    End If

End Sub
Private Sub Command11_Click()
    If Command11.Picture = imgBlank Then
    Command11.Picture = imgWhite
    ElseIf Command11.Picture = imgWhite Then
    Command11.Picture = imgBlack
    Else
    Command11.Picture = imgBlank
    End If
End Sub

Private Sub Command12_Click()
    If Command12.Picture = imgBlank Then
    Command12.Picture = imgWhite
    ElseIf Command12.Picture = imgWhite Then
    Command12.Picture = imgBlack
    Else
    Command12.Picture = imgBlank
    End If

End Sub

Private Sub Command13_Click()
    If Command13.Picture = imgBlank Then
    Command13.Picture = imgWhite
    ElseIf Command13.Picture = imgWhite Then
    Command13.Picture = imgBlack
    Else
    Command13.Picture = imgBlank
    End If

End Sub

Private Sub Command14_Click()
    If Command14.Picture = imgBlank Then
    Command14.Picture = imgWhite
    ElseIf Command14.Picture = imgWhite Then
    Command14.Picture = imgBlack
    Else
    Command14.Picture = imgBlank
    End If

End Sub

Private Sub Command15_Click()
    If Command15.Picture = imgBlank Then
    Command15.Picture = imgWhite
    ElseIf Command15.Picture = imgWhite Then
    Command15.Picture = imgBlack
    Else
    Command15.Picture = imgBlank
    End If

End Sub

Private Sub Command16_Click()
    If Command16.Picture = imgBlank Then
    Command16.Picture = imgWhite
    ElseIf Command16.Picture = imgWhite Then
    Command16.Picture = imgBlack
    Else
    Command16.Picture = imgBlank
    End If

End Sub

Private Sub Command17_Click()
    If Command17.Picture = imgBlank Then
    Command17.Picture = imgWhite
    ElseIf Command17.Picture = imgWhite Then
    Command17.Picture = imgBlack
    Else
    Command17.Picture = imgBlank
    End If

End Sub

Private Sub Command18_Click()
    If Command18.Picture = imgBlank Then
    Command18.Picture = imgWhite
    ElseIf Command18.Picture = imgWhite Then
    Command18.Picture = imgBlack
    Else
    Command18.Picture = imgBlank
    End If

End Sub

Private Sub Command19_Click()
    If Command19.Picture = imgBlank Then
    Command19.Picture = imgWhite
    ElseIf Command19.Picture = imgWhite Then
    Command19.Picture = imgBlack
    Else
    Command19.Picture = imgBlank
    End If

End Sub

'Private Sub Command10_Click()
'    If Command10.Picture = imgBlank Then
'    Command10.Picture = imgWhite
'    ElseIf Command10.Picture = imgWhite Then
'    Command10.Picture = imgBlack
'    Else
'    Command10.Picture = imgBlank
'    End If

'End Sub
Private Sub Command21_Click()
    If Command21.Picture = imgBlank Then
    Command21.Picture = imgWhite
    ElseIf Command21.Picture = imgWhite Then
    Command21.Picture = imgBlack
    Else
    Command21.Picture = imgBlank
    End If
End Sub

Private Sub Command22_Click()
    If Command22.Picture = imgBlank Then
    Command22.Picture = imgWhite
    ElseIf Command22.Picture = imgWhite Then
    Command22.Picture = imgBlack
    Else
    Command22.Picture = imgBlank
    End If

End Sub

Private Sub Command23_Click()
    If Command23.Picture = imgBlank Then
    Command23.Picture = imgWhite
    ElseIf Command23.Picture = imgWhite Then
    Command23.Picture = imgBlack
    Else
    Command23.Picture = imgBlank
    End If

End Sub

Private Sub Command24_Click()
    If Command24.Picture = imgBlank Then
    Command24.Picture = imgWhite
    ElseIf Command24.Picture = imgWhite Then
    Command24.Picture = imgBlack
    Else
    Command24.Picture = imgBlank
    End If

End Sub

Private Sub Command25_Click()
    If Command25.Picture = imgBlank Then
    Command25.Picture = imgWhite
    ElseIf Command25.Picture = imgWhite Then
    Command25.Picture = imgBlack
    Else
    Command25.Picture = imgBlank
    End If

End Sub

Private Sub Command26_Click()
    If Command26.Picture = imgBlank Then
    Command26.Picture = imgWhite
    ElseIf Command26.Picture = imgWhite Then
    Command26.Picture = imgBlack
    Else
    Command26.Picture = imgBlank
    End If

End Sub

Private Sub Command27_Click()
    If Command27.Picture = imgBlank Then
    Command27.Picture = imgWhite
    ElseIf Command27.Picture = imgWhite Then
    Command27.Picture = imgBlack
    Else
    Command27.Picture = imgBlank
    End If

End Sub

Private Sub Command28_Click()
    If Command28.Picture = imgBlank Then
    Command28.Picture = imgWhite
    ElseIf Command28.Picture = imgWhite Then
    Command28.Picture = imgBlack
    Else
    Command28.Picture = imgBlank
    End If

End Sub

Private Sub Command29_Click()
    If Command29.Picture = imgBlank Then
    Command29.Picture = imgWhite
    ElseIf Command29.Picture = imgWhite Then
    Command29.Picture = imgBlack
    Else
    Command29.Picture = imgBlank
    End If

End Sub

Private Sub Command30_Click()
    If Command30.Picture = imgBlank Then
    Command30.Picture = imgWhite
    ElseIf Command30.Picture = imgWhite Then
    Command30.Picture = imgBlack
    Else
    Command30.Picture = imgBlank
    End If

End Sub
Private Sub Command31_Click()
    If Command31.Picture = imgBlank Then
    Command31.Picture = imgWhite
    ElseIf Command31.Picture = imgWhite Then
    Command31.Picture = imgBlack
    Else
    Command31.Picture = imgBlank
    End If
End Sub

Private Sub Command32_Click()
    If Command32.Picture = imgBlank Then
    Command32.Picture = imgWhite
    ElseIf Command32.Picture = imgWhite Then
    Command32.Picture = imgBlack
    Else
    Command32.Picture = imgBlank
    End If

End Sub

Private Sub Command33_Click()
    If Command33.Picture = imgBlank Then
    Command33.Picture = imgWhite
    ElseIf Command33.Picture = imgWhite Then
    Command33.Picture = imgBlack
    Else
    Command33.Picture = imgBlank
    End If

End Sub

Private Sub Command34_Click()
    If Command34.Picture = imgBlank Then
    Command34.Picture = imgWhite
    ElseIf Command34.Picture = imgWhite Then
    Command34.Picture = imgBlack
    Else
    Command34.Picture = imgBlank
    End If

End Sub

Private Sub Command35_Click()
    If Command35.Picture = imgBlank Then
    Command35.Picture = imgWhite
    ElseIf Command35.Picture = imgWhite Then
    Command35.Picture = imgBlack
    Else
    Command35.Picture = imgBlank
    End If

End Sub

Private Sub Command36_Click()
    If Command36.Picture = imgBlank Then
    Command36.Picture = imgWhite
    ElseIf Command36.Picture = imgWhite Then
    Command36.Picture = imgBlack
    Else
    Command36.Picture = imgBlank
    End If

End Sub

Private Sub Command37_Click()
    If Command37.Picture = imgBlank Then
    Command37.Picture = imgWhite
    ElseIf Command37.Picture = imgWhite Then
    Command37.Picture = imgBlack
    Else
    Command37.Picture = imgBlank
    End If

End Sub

Private Sub Command38_Click()
    If Command38.Picture = imgBlank Then
    Command38.Picture = imgWhite
    ElseIf Command38.Picture = imgWhite Then
    Command38.Picture = imgBlack
    Else
    Command38.Picture = imgBlank
    End If

End Sub

Private Sub Command39_Click()
    If Command39.Picture = imgBlank Then
    Command39.Picture = imgWhite
    ElseIf Command39.Picture = imgWhite Then
    Command39.Picture = imgBlack
    Else
    Command39.Picture = imgBlank
    End If

End Sub

Private Sub Command40_Click()
    If Command40.Picture = imgBlank Then
    Command40.Picture = imgWhite
    ElseIf Command40.Picture = imgWhite Then
    Command40.Picture = imgBlack
    Else
    Command40.Picture = imgBlank
    End If

End Sub
Private Sub Command41_Click()
    If Command41.Picture = imgBlank Then
    Command41.Picture = imgWhite
    ElseIf Command41.Picture = imgWhite Then
    Command41.Picture = imgBlack
    Else
    Command41.Picture = imgBlank
    End If
End Sub

Private Sub Command42_Click()
    If Command42.Picture = imgBlank Then
    Command42.Picture = imgWhite
    ElseIf Command42.Picture = imgWhite Then
    Command42.Picture = imgBlack
    Else
    Command42.Picture = imgBlank
    End If

End Sub

Private Sub Command43_Click()
    If Command43.Picture = imgBlank Then
    Command43.Picture = imgWhite
    ElseIf Command43.Picture = imgWhite Then
    Command43.Picture = imgBlack
    Else
    Command43.Picture = imgBlank
    End If

End Sub

Private Sub Command44_Click()
    If Command44.Picture = imgBlank Then
    Command44.Picture = imgWhite
    ElseIf Command44.Picture = imgWhite Then
    Command44.Picture = imgBlack
    Else
    Command44.Picture = imgBlank
    End If

End Sub

Private Sub Command45_Click()
    If Command45.Picture = imgBlank Then
    Command45.Picture = imgWhite
    ElseIf Command45.Picture = imgWhite Then
    Command45.Picture = imgBlack
    Else
    Command45.Picture = imgBlank
    End If

End Sub

Private Sub Command46_Click()
    If Command46.Picture = imgBlank Then
    Command46.Picture = imgWhite
    ElseIf Command46.Picture = imgWhite Then
    Command46.Picture = imgBlack
    Else
    Command46.Picture = imgBlank
    End If

End Sub

Private Sub Command47_Click()
    If Command47.Picture = imgBlank Then
    Command47.Picture = imgWhite
    ElseIf Command47.Picture = imgWhite Then
    Command47.Picture = imgBlack
    Else
    Command47.Picture = imgBlank
    End If

End Sub

Private Sub Command48_Click()
    If Command48.Picture = imgBlank Then
    Command48.Picture = imgWhite
    ElseIf Command48.Picture = imgWhite Then
    Command48.Picture = imgBlack
    Else
    Command48.Picture = imgBlank
    End If

End Sub

Private Sub Command49_Click()
    If Command49.Picture = imgBlank Then
    Command49.Picture = imgWhite
    ElseIf Command49.Picture = imgWhite Then
    Command49.Picture = imgBlack
    Else
    Command49.Picture = imgBlank
    End If

End Sub

Private Sub Command50_Click()
    If Command50.Picture = imgBlank Then
    Command50.Picture = imgWhite
    ElseIf Command50.Picture = imgWhite Then
    Command50.Picture = imgBlack
    Else
    Command50.Picture = imgBlank
    End If

End Sub
Private Sub Command51_Click()
    If Command51.Picture = imgBlank Then
    Command51.Picture = imgWhite
    ElseIf Command51.Picture = imgWhite Then
    Command51.Picture = imgBlack
    Else
    Command51.Picture = imgBlank
    End If
End Sub

Private Sub Command52_Click()
    If Command52.Picture = imgBlank Then
    Command52.Picture = imgWhite
    ElseIf Command52.Picture = imgWhite Then
    Command52.Picture = imgBlack
    Else
    Command52.Picture = imgBlank
    End If

End Sub

Private Sub Command53_Click()
    If Command53.Picture = imgBlank Then
    Command53.Picture = imgWhite
    ElseIf Command53.Picture = imgWhite Then
    Command53.Picture = imgBlack
    Else
    Command53.Picture = imgBlank
    End If

End Sub

Private Sub Command54_Click()
    If Command54.Picture = imgBlank Then
    Command54.Picture = imgWhite
    ElseIf Command54.Picture = imgWhite Then
    Command54.Picture = imgBlack
    Else
    Command54.Picture = imgBlank
    End If

End Sub

Private Sub Command55_Click()
    If Command55.Picture = imgBlank Then
    Command55.Picture = imgWhite
    ElseIf Command55.Picture = imgWhite Then
    Command55.Picture = imgBlack
    Else
    Command55.Picture = imgBlank
    End If

End Sub

Private Sub Command56_Click()
    If Command56.Picture = imgBlank Then
    Command56.Picture = imgWhite
    ElseIf Command56.Picture = imgWhite Then
    Command56.Picture = imgBlack
    Else
    Command56.Picture = imgBlank
    End If

End Sub

Private Sub Command57_Click()
    If Command57.Picture = imgBlank Then
    Command57.Picture = imgWhite
    ElseIf Command57.Picture = imgWhite Then
    Command57.Picture = imgBlack
    Else
    Command57.Picture = imgBlank
    End If

End Sub

Private Sub Command58_Click()
    If Command58.Picture = imgBlank Then
    Command58.Picture = imgWhite
    ElseIf Command58.Picture = imgWhite Then
    Command58.Picture = imgBlack
    Else
    Command58.Picture = imgBlank
    End If

End Sub

Private Sub Command59_Click()
    If Command59.Picture = imgBlank Then
    Command59.Picture = imgWhite
    ElseIf Command59.Picture = imgWhite Then
    Command59.Picture = imgBlack
    Else
    Command59.Picture = imgBlank
    End If

End Sub

Private Sub Command60_Click()
    If Command60.Picture = imgBlank Then
    Command60.Picture = imgWhite
    ElseIf Command60.Picture = imgWhite Then
    Command60.Picture = imgBlack
    Else
    Command60.Picture = imgBlank
    End If

End Sub
Private Sub Command61_Click()
    If Command61.Picture = imgBlank Then
    Command61.Picture = imgWhite
    ElseIf Command61.Picture = imgWhite Then
    Command61.Picture = imgBlack
    Else
    Command61.Picture = imgBlank
    End If
End Sub

Private Sub Command62_Click()
    If Command62.Picture = imgBlank Then
    Command62.Picture = imgWhite
    ElseIf Command62.Picture = imgWhite Then
    Command62.Picture = imgBlack
    Else
    Command62.Picture = imgBlank
    End If

End Sub

Private Sub Command63_Click()
    If Command63.Picture = imgBlank Then
    Command63.Picture = imgWhite
    ElseIf Command63.Picture = imgWhite Then
    Command63.Picture = imgBlack
    Else
    Command63.Picture = imgBlank
    End If

End Sub

Private Sub Command64_Click()
    If Command64.Picture = imgBlank Then
    Command64.Picture = imgWhite
    ElseIf Command64.Picture = imgWhite Then
    Command64.Picture = imgBlack
    Else
    Command64.Picture = imgBlank
    End If

End Sub

Private Sub Command65_Click()
    If Command65.Picture = imgBlank Then
    Command65.Picture = imgWhite
    ElseIf Command65.Picture = imgWhite Then
    Command65.Picture = imgBlack
    Else
    Command65.Picture = imgBlank
    End If

End Sub

Private Sub Command66_Click()
    If Command66.Picture = imgBlank Then
    Command66.Picture = imgWhite
    ElseIf Command66.Picture = imgWhite Then
    Command66.Picture = imgBlack
    Else
    Command66.Picture = imgBlank
    End If

End Sub

Private Sub Command67_Click()
    If Command67.Picture = imgBlank Then
    Command67.Picture = imgWhite
    ElseIf Command67.Picture = imgWhite Then
    Command67.Picture = imgBlack
    Else
    Command67.Picture = imgBlank
    End If

End Sub

Private Sub Command68_Click()
    If Command68.Picture = imgBlank Then
    Command68.Picture = imgWhite
    ElseIf Command68.Picture = imgWhite Then
    Command68.Picture = imgBlack
    Else
    Command68.Picture = imgBlank
    End If

End Sub

Private Sub Command69_Click()
    If Command69.Picture = imgBlank Then
    Command69.Picture = imgWhite
    ElseIf Command69.Picture = imgWhite Then
    Command69.Picture = imgBlack
    Else
    Command69.Picture = imgBlank
    End If

End Sub

Private Sub Command70_Click()
    If Command70.Picture = imgBlank Then
    Command70.Picture = imgWhite
    ElseIf Command70.Picture = imgWhite Then
    Command70.Picture = imgBlack
    Else
    Command70.Picture = imgBlank
    End If

End Sub
Private Sub Command71_Click()
    If Command71.Picture = imgBlank Then
    Command71.Picture = imgWhite
    ElseIf Command71.Picture = imgWhite Then
    Command71.Picture = imgBlack
    Else
    Command71.Picture = imgBlank
    End If
End Sub

Private Sub Command72_Click()
    If Command72.Picture = imgBlank Then
    Command72.Picture = imgWhite
    ElseIf Command72.Picture = imgWhite Then
    Command72.Picture = imgBlack
    Else
    Command72.Picture = imgBlank
    End If

End Sub

Private Sub Command73_Click()
    If Command73.Picture = imgBlank Then
    Command73.Picture = imgWhite
    ElseIf Command73.Picture = imgWhite Then
    Command73.Picture = imgBlack
    Else
    Command73.Picture = imgBlank
    End If

End Sub

Private Sub Command74_Click()
    If Command74.Picture = imgBlank Then
    Command74.Picture = imgWhite
    ElseIf Command74.Picture = imgWhite Then
    Command74.Picture = imgBlack
    Else
    Command74.Picture = imgBlank
    End If

End Sub

Private Sub Command75_Click()
    If Command75.Picture = imgBlank Then
    Command75.Picture = imgWhite
    ElseIf Command75.Picture = imgWhite Then
    Command75.Picture = imgBlack
    Else
    Command75.Picture = imgBlank
    End If

End Sub

Private Sub Command76_Click()
    If Command76.Picture = imgBlank Then
    Command76.Picture = imgWhite
    ElseIf Command76.Picture = imgWhite Then
    Command76.Picture = imgBlack
    Else
    Command76.Picture = imgBlank
    End If

End Sub

Private Sub Command77_Click()
    If Command77.Picture = imgBlank Then
    Command77.Picture = imgWhite
    ElseIf Command77.Picture = imgWhite Then
    Command77.Picture = imgBlack
    Else
    Command77.Picture = imgBlank
    End If

End Sub

Private Sub Command78_Click()
    If Command78.Picture = imgBlank Then
    Command78.Picture = imgWhite
    ElseIf Command78.Picture = imgWhite Then
    Command78.Picture = imgBlack
    Else
    Command78.Picture = imgBlank
    End If

End Sub

Private Sub Command79_Click()
    If Command79.Picture = imgBlank Then
    Command79.Picture = imgWhite
    ElseIf Command79.Picture = imgWhite Then
    Command79.Picture = imgBlack
    Else
    Command79.Picture = imgBlank
    End If

End Sub

Private Sub Command80_Click()
    If Command80.Picture = imgBlank Then
    Command80.Picture = imgWhite
    ElseIf Command80.Picture = imgWhite Then
    Command80.Picture = imgBlack
    Else
    Command80.Picture = imgBlank
    End If

End Sub
Private Sub Command81_Click()
    If Command81.Picture = imgBlank Then
    Command81.Picture = imgWhite
    ElseIf Command81.Picture = imgWhite Then
    Command81.Picture = imgBlack
    Else
    Command81.Picture = imgBlank
    End If
End Sub

Private Sub Command82_Click()
    If Command82.Picture = imgBlank Then
    Command82.Picture = imgWhite
    ElseIf Command82.Picture = imgWhite Then
    Command82.Picture = imgBlack
    Else
    Command82.Picture = imgBlank
    End If

End Sub

Private Sub Command83_Click()
    If Command83.Picture = imgBlank Then
    Command83.Picture = imgWhite
    ElseIf Command83.Picture = imgWhite Then
    Command83.Picture = imgBlack
    Else
    Command83.Picture = imgBlank
    End If

End Sub

Private Sub Command84_Click()
    If Command84.Picture = imgBlank Then
    Command84.Picture = imgWhite
    ElseIf Command84.Picture = imgWhite Then
    Command84.Picture = imgBlack
    Else
    Command84.Picture = imgBlank
    End If

End Sub

Private Sub Command85_Click()
    If Command85.Picture = imgBlank Then
    Command85.Picture = imgWhite
    ElseIf Command85.Picture = imgWhite Then
    Command85.Picture = imgBlack
    Else
    Command85.Picture = imgBlank
    End If

End Sub

Private Sub Command86_Click()
    If Command86.Picture = imgBlank Then
    Command86.Picture = imgWhite
    ElseIf Command86.Picture = imgWhite Then
    Command86.Picture = imgBlack
    Else
    Command86.Picture = imgBlank
    End If

End Sub

Private Sub Command87_Click()
    If Command87.Picture = imgBlank Then
    Command87.Picture = imgWhite
    ElseIf Command87.Picture = imgWhite Then
    Command87.Picture = imgBlack
    Else
    Command87.Picture = imgBlank
    End If

End Sub

Private Sub Command88_Click()
    If Command88.Picture = imgBlank Then
    Command88.Picture = imgWhite
    ElseIf Command88.Picture = imgWhite Then
    Command88.Picture = imgBlack
    Else
    Command88.Picture = imgBlank
    End If

End Sub

Private Sub Command89_Click()
    If Command89.Picture = imgBlank Then
    Command89.Picture = imgWhite
    ElseIf Command89.Picture = imgWhite Then
    Command89.Picture = imgBlack
    Else
    Command89.Picture = imgBlank
    End If

End Sub

Private Sub Command90_Click()
    If Command90.Picture = imgBlank Then
    Command90.Picture = imgWhite
    ElseIf Command90.Picture = imgWhite Then
    Command90.Picture = imgBlack
    Else
    Command90.Picture = imgBlank
    End If

End Sub
Private Sub Command91_Click()
    If Command91.Picture = imgBlank Then
    Command91.Picture = imgWhite
    ElseIf Command91.Picture = imgWhite Then
    Command91.Picture = imgBlack
    Else
    Command91.Picture = imgBlank
    End If
End Sub

Private Sub Command92_Click()
    If Command92.Picture = imgBlank Then
    Command92.Picture = imgWhite
    ElseIf Command92.Picture = imgWhite Then
    Command92.Picture = imgBlack
    Else
    Command92.Picture = imgBlank
    End If

End Sub

Private Sub Command93_Click()
    If Command93.Picture = imgBlank Then
    Command93.Picture = imgWhite
    ElseIf Command93.Picture = imgWhite Then
    Command93.Picture = imgBlack
    Else
    Command93.Picture = imgBlank
    End If

End Sub

Private Sub Command94_Click()
    If Command94.Picture = imgBlank Then
    Command94.Picture = imgWhite
    ElseIf Command94.Picture = imgWhite Then
    Command94.Picture = imgBlack
    Else
    Command94.Picture = imgBlank
    End If

End Sub

Private Sub Command95_Click()
    If Command95.Picture = imgBlank Then
    Command95.Picture = imgWhite
    ElseIf Command95.Picture = imgWhite Then
    Command95.Picture = imgBlack
    Else
    Command95.Picture = imgBlank
    End If

End Sub

Private Sub Command96_Click()
    If Command96.Picture = imgBlank Then
    Command96.Picture = imgWhite
    ElseIf Command96.Picture = imgWhite Then
    Command96.Picture = imgBlack
    Else
    Command96.Picture = imgBlank
    End If

End Sub

Private Sub Command97_Click()
    If Command97.Picture = imgBlank Then
    Command97.Picture = imgWhite
    ElseIf Command97.Picture = imgWhite Then
    Command97.Picture = imgBlack
    Else
    Command97.Picture = imgBlank
    End If

End Sub

Private Sub Command98_Click()
    If Command98.Picture = imgBlank Then
    Command98.Picture = imgWhite
    ElseIf Command98.Picture = imgWhite Then
    Command98.Picture = imgBlack
    Else
    Command98.Picture = imgBlank
    End If

End Sub

Private Sub Command99_Click()
    If Command99.Picture = imgBlank Then
    Command99.Picture = imgWhite
    ElseIf Command99.Picture = imgWhite Then
    Command99.Picture = imgBlack
    Else
    Command99.Picture = imgBlank
    End If

End Sub

Private Sub Command100_Click()
    If Command100.Picture = imgBlank Then
    Command100.Picture = imgWhite
    ElseIf Command100.Picture = imgWhite Then
    Command100.Picture = imgBlack
    Else
    Command100.Picture = imgBlank
    End If

End Sub

Private Sub Form_Load()
Command1.Picture = imgBlank
Command2.Picture = imgBlank
Command3.Picture = imgBlank
Command4.Picture = imgBlank
Command5.Picture = imgBlank
Command6.Picture = imgBlank
Command7.Picture = imgBlank
Command8.Picture = imgBlank
Command9.Picture = imgBlank
Command10.Picture = imgBlank
Command11.Picture = imgBlank
Command12.Picture = imgBlank
Command13.Picture = imgBlank
Command14.Picture = imgBlank
Command15.Picture = imgBlank
Command16.Picture = imgBlank
Command17.Picture = imgBlank
Command18.Picture = imgBlank
Command19.Picture = imgBlank
Command20.Picture = imgBlank
Command21.Picture = imgBlank
Command22.Picture = imgBlank
Command23.Picture = imgBlank
Command24.Picture = imgBlank
Command25.Picture = imgBlank
Command26.Picture = imgBlank
Command27.Picture = imgBlank
Command28.Picture = imgBlank
Command29.Picture = imgBlank
Command30.Picture = imgBlank
Command31.Picture = imgBlank
Command32.Picture = imgBlank
Command33.Picture = imgBlank
Command34.Picture = imgBlank
Command35.Picture = imgBlank
Command36.Picture = imgBlank
Command37.Picture = imgBlank
Command38.Picture = imgBlank
Command39.Picture = imgBlank
Command40.Picture = imgBlank
Command41.Picture = imgBlank
Command42.Picture = imgBlank
Command43.Picture = imgBlank
Command44.Picture = imgBlank
Command45.Picture = imgBlank
Command46.Picture = imgBlank
Command47.Picture = imgBlank
Command48.Picture = imgBlank
Command49.Picture = imgBlank
Command50.Picture = imgBlank
Command51.Picture = imgBlank
Command52.Picture = imgBlank
Command53.Picture = imgBlank
Command54.Picture = imgBlank
Command55.Picture = imgBlank
Command56.Picture = imgBlank
Command57.Picture = imgBlank
Command58.Picture = imgBlank
Command59.Picture = imgBlank
Command60.Picture = imgBlank
Command61.Picture = imgBlank
Command62.Picture = imgBlank
Command63.Picture = imgBlank
Command64.Picture = imgBlank
Command65.Picture = imgBlank
Command66.Picture = imgBlank
Command67.Picture = imgBlank
Command68.Picture = imgBlank
Command69.Picture = imgBlank
Command70.Picture = imgBlank
Command71.Picture = imgBlank
Command72.Picture = imgBlank
Command73.Picture = imgBlank
Command74.Picture = imgBlank
Command75.Picture = imgBlank
Command76.Picture = imgBlank
Command77.Picture = imgBlank
Command78.Picture = imgBlank
Command79.Picture = imgBlank
Command80.Picture = imgBlank
Command81.Picture = imgBlank
Command82.Picture = imgBlank
Command83.Picture = imgBlank
Command84.Picture = imgBlank
Command85.Picture = imgBlank
Command86.Picture = imgBlank
Command87.Picture = imgBlank
Command88.Picture = imgBlank
Command89.Picture = imgBlank
Command90.Picture = imgBlank
Command91.Picture = imgBlank
Command92.Picture = imgBlank
Command93.Picture = imgBlank
Command94.Picture = imgBlank
Command95.Picture = imgBlank
Command96.Picture = imgBlank
Command97.Picture = imgBlank
Command98.Picture = imgBlank
Command99.Picture = imgBlank
Command100.Picture = imgBlank
End Sub
