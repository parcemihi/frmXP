VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   7800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   6120
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Left            =   6480
      Top             =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   149
      Left            =   5400
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   148
      Left            =   5040
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   147
      Left            =   4680
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   146
      Left            =   4320
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   145
      Left            =   3960
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   144
      Left            =   3600
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   143
      Left            =   3240
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   142
      Left            =   2880
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   141
      Left            =   2520
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   140
      Left            =   2160
      Top             =   5400
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   139
      Left            =   5400
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   138
      Left            =   5040
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   137
      Left            =   4680
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   136
      Left            =   4320
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   135
      Left            =   3960
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   134
      Left            =   3600
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   133
      Left            =   3240
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   132
      Left            =   2880
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   131
      Left            =   2520
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   130
      Left            =   2160
      Top             =   5040
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   129
      Left            =   5400
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   128
      Left            =   5040
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   127
      Left            =   4680
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   126
      Left            =   4320
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   125
      Left            =   3960
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   124
      Left            =   3600
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   123
      Left            =   3240
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   122
      Left            =   2880
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   121
      Left            =   2520
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   120
      Left            =   2160
      Top             =   4680
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   119
      Left            =   5400
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   118
      Left            =   5040
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   117
      Left            =   4680
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   116
      Left            =   4320
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   115
      Left            =   3960
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   114
      Left            =   3600
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   113
      Left            =   3240
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   112
      Left            =   2880
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   111
      Left            =   2520
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   110
      Left            =   2160
      Top             =   4320
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   109
      Left            =   5400
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   108
      Left            =   5040
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   107
      Left            =   4680
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   106
      Left            =   4320
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   105
      Left            =   3960
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   104
      Left            =   3600
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   103
      Left            =   3240
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   102
      Left            =   2880
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   101
      Left            =   2520
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   100
      Left            =   2160
      Top             =   3960
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   99
      Left            =   5400
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   98
      Left            =   5040
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   97
      Left            =   4680
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   96
      Left            =   4320
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   95
      Left            =   3960
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   94
      Left            =   3600
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   93
      Left            =   3240
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   92
      Left            =   2880
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   91
      Left            =   2520
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   90
      Left            =   2160
      Top             =   3600
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   89
      Left            =   5400
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   88
      Left            =   5040
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   87
      Left            =   4680
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   86
      Left            =   4320
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   85
      Left            =   3960
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   84
      Left            =   3600
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   83
      Left            =   3240
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   82
      Left            =   2880
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   81
      Left            =   2520
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   80
      Left            =   2160
      Top             =   3240
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   79
      Left            =   5400
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   78
      Left            =   5040
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   77
      Left            =   4680
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   76
      Left            =   4320
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   75
      Left            =   3960
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   74
      Left            =   3600
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   73
      Left            =   3240
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   72
      Left            =   2880
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   71
      Left            =   2520
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   70
      Left            =   2160
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   69
      Left            =   5400
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   68
      Left            =   5040
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   67
      Left            =   4680
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   66
      Left            =   4320
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   65
      Left            =   3960
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   64
      Left            =   3600
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   63
      Left            =   3240
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   62
      Left            =   2880
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   61
      Left            =   2520
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   60
      Left            =   2160
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   59
      Left            =   5400
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   58
      Left            =   5040
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   57
      Left            =   4680
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   56
      Left            =   4320
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   55
      Left            =   3960
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   54
      Left            =   3600
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   53
      Left            =   3240
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   52
      Left            =   2880
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   51
      Left            =   2520
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   50
      Left            =   2160
      Top             =   2160
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   49
      Left            =   5400
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   48
      Left            =   5040
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   47
      Left            =   4680
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   46
      Left            =   4320
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   45
      Left            =   3960
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   44
      Left            =   3600
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   43
      Left            =   3240
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   42
      Left            =   2880
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   41
      Left            =   2520
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   40
      Left            =   2160
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   39
      Left            =   5400
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   38
      Left            =   5040
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   37
      Left            =   4680
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   36
      Left            =   4320
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   35
      Left            =   3960
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   34
      Left            =   3600
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   33
      Left            =   3240
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   32
      Left            =   2880
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   31
      Left            =   2520
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   30
      Left            =   2160
      Top             =   1440
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   29
      Left            =   5400
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   28
      Left            =   5040
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   27
      Left            =   4680
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   26
      Left            =   4320
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   25
      Left            =   3960
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   24
      Left            =   3600
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   23
      Left            =   3240
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   22
      Left            =   2880
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   21
      Left            =   2520
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   20
      Left            =   2160
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   19
      Left            =   5400
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   18
      Left            =   5040
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   17
      Left            =   4680
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   16
      Left            =   4320
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   15
      Left            =   3960
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   14
      Left            =   3600
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   13
      Left            =   3240
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   12
      Left            =   2880
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   11
      Left            =   2520
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   10
      Left            =   2160
      Top             =   720
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   9
      Left            =   5400
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   8
      Left            =   5040
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   7
      Left            =   4680
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   6
      Left            =   4320
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   5
      Left            =   3960
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   4
      Left            =   3600
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   3
      Left            =   3240
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   2
      Left            =   2880
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   1
      Left            =   2520
      Top             =   360
      Width           =   375
   End
   Begin VB.Image imgBoard 
      Height          =   375
      Index           =   0
      Left            =   2160
      Top             =   360
      Width           =   375
   End
   Begin VB.Image Image7 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":0000
      Top             =   2880
      Width           =   360
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":0702
      Top             =   2160
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":0E04
      Top             =   2520
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":1506
      Top             =   1800
      Width           =   360
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":1C08
      Top             =   1440
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":230A
      Top             =   1080
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   360
      Picture         =   "frmTetris2.frx":2A0C
      Top             =   720
      Width           =   360
   End
   Begin VB.Shape Shape1 
      Height          =   5415
      Left            =   2160
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i1 As Integer, i2 As Integer, _
    i3 As Integer, i4 As Integer
Dim rndom As Single

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyLeft Then
    If i4 = 0 Or i4 = 10 Or i4 = 20 Or i4 = 30 _
   Or i4 = 40 Or i4 = 50 Or i4 = 60 Or i4 = 70 _
   Or i4 = 80 Or i4 = 90 Or i4 = 100 Or i4 = 110 _
   Or i4 = 120 Or i4 = 130 Or i4 = 140 Then
   Else
   imgBoard(i1).Visible = False
        i1 = i1 - 1
        imgBoard(i1).Visible = True
        imgBoard(i2).Visible = False
        i2 = i2 - 1
        imgBoard(i2).Visible = True
        imgBoard(i3).Visible = False
        i3 = i3 - 1
        imgBoard(i3).Visible = True
        imgBoard(i4).Visible = False
        i4 = i4 - 1
        imgBoard(i4).Visible = True
    
    End If
    End If
    
    If KeyCode = vbKeyRight Then
    If i1 = 9 Or i1 = 19 Or i1 = 29 Or i1 = 39 _
    Or i1 = 49 Or i1 = 59 Or i1 = 69 Or i1 = 79 _
    Or i1 = 89 Or i1 = 99 Or i1 = 109 Or i1 = 119 _
    Or i1 = 129 Or i1 = 139 Or i1 = 149 Then
    Else
        imgBoard(i1).Visible = False
        i1 = i1 + 1
        imgBoard(i1).Visible = True
        imgBoard(i2).Visible = False
        i2 = i2 + 1
        imgBoard(i2).Visible = True
        imgBoard(i3).Visible = False
        i3 = i3 + 1
        imgBoard(i3).Visible = True
        imgBoard(i4).Visible = False
        i4 = i4 + 1
        imgBoard(i4).Visible = True
    End If
    End If
    
    If KeyCode = vbKeyUp Then
    If i1 + 10 = i2 Then
        imgBoard(i1).Visible = False
        i1 = i1 + 11
        imgBoard(i1).Visible = True
        imgBoard(i2).Visible = False
        i2 = i2
        imgBoard(i2).Visible = True
        imgBoard(i3).Visible = False
        i3 = i3 - 11
        imgBoard(i3).Visible = True
        imgBoard(i4).Visible = False
        i4 = i4 - 22
        imgBoard(i4).Visible = True
    ElseIf i1 - 1 = i2 Then
        imgBoard(i1).Visible = False
        i1 = i1 - 11
        imgBoard(i1).Visible = True
        
        imgBoard(i2).Visible = False
        i2 = i2
        imgBoard(i2).Visible = True
        
        imgBoard(i3).Visible = False
        i3 = i3 + 11
        imgBoard(i3).Visible = True
        
        imgBoard(i4).Visible = False
        i4 = i4 + 22
        imgBoard(i4).Visible = True
    End If
    End If
    
    If KeyCode = vbKeyF2 Then
        genColor
        Timer1.Enabled = True
    End If
    
    If KeyCode = vbKeyF3 Then
        MsgBox "Game paused"
        If vbOKOnly = 0 Then
        Timer1.Enabled = True
        End If
    End If
End Sub
Private Sub Form_Load()
Command1.Enabled = False
    Timer1.Enabled = True
    Timer1.Interval = 1000
    genColor
End Sub
Public Sub genColor()
Randomize

rndom = Rnd
rndom = Int(rndom * 700)

Select Case rndom
    Case 0 To 700
    i1 = 5
    i2 = 15
    i3 = 25
    i4 = 35
    Timer1_Timer
    Case 100 To 199
    Case 200 To 299
    Case 300 To 399
    Case 400 To 499
    Case 500 To 599
    Case 600 To 700
End Select

End Sub

Private Sub Timer1_Timer()

    
If i1 + 10 = i2 Then
    imgBoard(i1).Visible = False
    i1 = i1 + 10
    imgBoard(i1).Visible = True
    imgBoard(i1) = Image1.Picture

'    imgBoard(i2).Visible = False
    i2 = i2 + 10
    imgBoard(i2).Visible = True
    imgBoard(i2) = Image1.Picture
    
'    imgBoard(i3).Visible = False
    i3 = i3 + 10
    imgBoard(i3).Visible = True
    imgBoard(i3) = Image1.Picture
    
'    imgBoard(i4).Visible = False
    i4 = i4 + 10
    imgBoard(i4).Visible = True
    imgBoard(i4) = Image1.Picture

    If i1 > 139 Or i2 > 139 Or i3 > 139 Or i4 > 139 Then
    genColor
    End If
    
End If




If i1 - 1 = i2 Then
    If imgBoard(i1 + 10).Visible = True Or _
       imgBoard(i2 + 10).Visible = True Or _
       imgBoard(i3 + 10).Visible = True Or _
       imgBoard(i4 + 10).Visible = True Then
'    genColor
    Else
'    imgBoard(i2).Visible = False
'    imgBoard(i3).Visible = False
'    imgBoard(i4).Visible = False
    End If
    End If
    
If i1 - 1 = i2 Then
    imgBoard(i1).Visible = False
    i1 = i1 + 10
    imgBoard(i1).Visible = True
    imgBoard(i1) = Image1.Picture
    
    
    imgBoard(i2).Visible = False
    i2 = i2 + 10
    imgBoard(i2).Visible = True
    imgBoard(i2) = Image1.Picture
    
    
    imgBoard(i3).Visible = False
    i3 = i3 + 10
    imgBoard(i3).Visible = True
    imgBoard(i3) = Image1.Picture
    
    
    imgBoard(i4).Visible = False
    i4 = i4 + 10
    imgBoard(i4).Visible = True
    imgBoard(i4) = Image1.Picture

    If i1 > 139 Or i2 > 139 Or i3 > 139 Or i4 > 139 Then
    genColor
    End If
End If
End Sub
