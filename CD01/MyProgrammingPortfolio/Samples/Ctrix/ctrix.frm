VERSION 5.00
Object = "{6DE6E6DD-C656-11D2-B052-444553540000}#3.0#0"; "VBCARDS.OCX"
Begin VB.Form Form1 
   Caption         =   "Ctrix"
   ClientHeight    =   7215
   ClientLeft      =   2115
   ClientTop       =   1320
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   6495
   Begin VBCards.Deck Deck1 
      Left            =   0
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   1032
      Picture         =   "ctrix.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COL 5"
      Height          =   375
      Index           =   4
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COL 4"
      Height          =   375
      Index           =   3
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COL 3"
      Height          =   375
      Index           =   2
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COL 2"
      Height          =   375
      Index           =   1
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "COL 1"
      Height          =   375
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "2) Click COL button under the column with your card."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   6720
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Is this the card you were thinking of? . . ."
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "1)  Think of a card ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6360
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   24
      Left            =   5160
      Top             =   4200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   23
      Left            =   3960
      Top             =   4200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   22
      Left            =   2760
      Top             =   4200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   21
      Left            =   1560
      Top             =   4200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   20
      Left            =   360
      Top             =   4200
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   19
      Left            =   5160
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   18
      Left            =   3960
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   17
      Left            =   2760
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   16
      Left            =   1560
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   15
      Left            =   360
      Top             =   3120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   14
      Left            =   5160
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   13
      Left            =   3960
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   12
      Left            =   2760
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   11
      Left            =   1560
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   10
      Left            =   360
      Top             =   2040
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   9
      Left            =   5160
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   8
      Left            =   3960
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   7
      Left            =   2760
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   6
      Left            =   1560
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1455
      Index           =   5
      Left            =   360
      Top             =   960
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   4
      Left            =   5160
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   3
      Left            =   3960
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   2
      Left            =   2760
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   1
      Left            =   1560
      Top             =   0
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Index           =   0
      Left            =   360
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'     Card Trix - A basic card trick
'
'     Program By:  Max Seim  mlseim@mmm.com
'
'     An old card trick adapted for VB ...
'     been done before, but I just thought it
'     would be fun to try it myself.
'
Dim p As Integer
Dim r As Integer
Dim x As Integer
Dim y As Integer
Dim row As Integer
Dim row1 As Integer
Dim row2 As Integer
Dim table(24) As Integer
Dim deck(54) As Integer
Dim fill(54) As Integer

Private Sub Command1_Click(Index As Integer)
   If row = 2 Then
   row2 = Index
   Label1.Caption = ""
   Label3.Caption = ""
   Call finish
   Exit Sub
   End If
Randomize
r = Int(Rnd(1) * 5 + 1)
st = (r - 1)
row1 = 5 - st
t = 0
Randomize
p = Int(Rnd(1) * 2 + 1)
If p = 1 Then
For x = 0 To 4
   For y = st To (st + 20) Step 5
      If Index = st Then
      row1 = x
      End If
   Deck1.ChangeCard = deck(y + 1)
   Image1(t).Picture = Deck1.Picture
   t = t + 1
   Next y
      st = st + 1
         If st = 5 Then
         st = 0
         End If
Next x
End If

If p = 2 Then
For x = 0 To 4
   For y = (st + 20) To st Step -5
      If Index = st Then
      row1 = x
      End If
   Deck1.ChangeCard = deck(y + 1)
   Image1(t).Picture = Deck1.Picture
   t = t + 1
   Next y
      st = st + 1
         If st = 5 Then
         st = 0
         End If
Next x
End If
Label1.Caption = "3) Now, once again ..."
Label3.Caption = "   Click COL button under the column with your card."
row = 2
End Sub

Private Sub Command2_Click()
Call clear
Call shuffle
End Sub

Private Sub Form_Load()
Call shuffle
row = 1
End Sub
Private Sub shuffle()
' -------------------------------------------------
'  This is a pretty good shuffle routine ...
'  Copy and Paste into your own card program.
'
Erase deck
Erase fill
Randomize
r = Int(Rnd(1) * 52 + 1)
For x = 1 To 52
   Do While fill(r) = 1
   r = Int(Rnd(1) * 52 + 1)
   Loop
fill(r) = 1
deck(x) = r
Next x
'
' -------------------------------------------------
'
For x = 0 To 24
Deck1.ChangeCard = deck(x + 1)
Image1(x).Picture = Deck1.Picture
Next

End Sub
Private Sub finish()
For x = 0 To 24
Image1(x).Visible = False
Next x
For x = 0 To 4
Command1(x).Visible = False
Next x
Label2.Visible = True
Command2.Visible = True
Image1(12).Picture = Image1((row1 * 5) + row2).Picture
Image1(12).Visible = True
End Sub
Private Sub clear()
For x = 0 To 24
Image1(x).Visible = True
Next x
For x = 0 To 4
Command1(x).Visible = True
Next x
Label2.Visible = False
Command2.Visible = False
row = 1
Label1.Caption = "1) Think of a card ..."
Label3.Caption = "2) Click COL button under the column with your card."
End Sub
