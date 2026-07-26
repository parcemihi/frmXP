VERSION 5.00
Begin VB.Form frmRate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate for the Region"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   2925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "East Region"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "West Region"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "South Region"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "North Region"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
perNorth = Text1.Text
perSouth = Text2.Text
perWest = Text3.Text
perEast = Text4.Text
frmRMsales.Show
frmRate.Hide

End Sub

Private Sub Form_Load()
    Text1.Text = 0
    Text2.Text = 0
    Text3.Text = 0
    Text4.Text = 0
    Text1.Text = Format(Text1.Text, "percent")
    Text2.Text = Format(Text2.Text, "percent")
    Text3.Text = Format(Text3.Text, "percent")
    Text4.Text = Format(Text4.Text, "percent")
End Sub



Private Sub Text1_GotFocus()
    Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text2_GotFocus()
    Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text3_GotFocus()
    Text3.SelLength = Len(Text3.Text)
End Sub
Private Sub Text4_GotFocus()
    Text4.SelLength = Len(Text4.Text)
End Sub


Private Sub Text1_LostFocus()
    Text1.Text = Format(Text1.Text, "percent")
End Sub
Private Sub Text2_LostFocus()
    Text2.Text = Format(Text2.Text, "percent")
End Sub
Private Sub Text3_LostFocus()
    Text3.Text = Format(Text3.Text, "percent")
End Sub
Private Sub Text4_LostFocus()
    Text4.Text = Format(Text4.Text, "percent")
End Sub
