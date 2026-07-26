VERSION 5.00
Begin VB.Form BookLoan 
   Caption         =   "Registered Nursing - Book Loan Database"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ending 
      Caption         =   "Move to &Last"
      Height          =   495
      Left            =   4320
      TabIndex        =   29
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton begining 
      Caption         =   "Move to F&irst"
      Height          =   495
      Left            =   360
      TabIndex        =   28
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   27
      Top             =   210
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   11
      Left            =   1920
      TabIndex        =   26
      Top             =   3720
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   10
      Left            =   1920
      TabIndex        =   25
      Top             =   3345
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1920
      TabIndex        =   24
      Top             =   2985
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   23
      Top             =   2610
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   22
      Top             =   2250
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   21
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton ExitRec 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4320
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton UpdateRec 
      Caption         =   "&Update"
      Height          =   495
      Left            =   3000
      TabIndex        =   18
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton NextRec 
      Caption         =   "&Next"
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton PreviousRec 
      Caption         =   "&Previous"
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton DeleteRec 
      Caption         =   "&Delete Record"
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton AddRec 
      Caption         =   "&Add Record"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   13
      Top             =   570
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   11
      Top             =   1305
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Top             =   945
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   20
      Top             =   1305
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Top             =   1305
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Book Author:"
      Height          =   195
      Index           =   1
      Left            =   840
      TabIndex        =   12
      Top             =   2700
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Telephone:"
      Height          =   195
      Index           =   10
      Left            =   960
      TabIndex        =   8
      Top             =   1770
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "City/State/Zip:"
      Height          =   195
      Index           =   9
      Left            =   720
      TabIndex        =   7
      Top             =   1395
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address 2:"
      Height          =   195
      Index           =   8
      Left            =   1020
      TabIndex        =   6
      Top             =   1035
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Address 1:"
      Height          =   195
      Index           =   7
      Left            =   1020
      TabIndex        =   5
      Top             =   660
      Width           =   750
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Owner of Book:"
      Height          =   195
      Index           =   5
      Left            =   660
      TabIndex        =   4
      Top             =   300
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Return date:"
      Height          =   195
      Index           =   4
      Left            =   885
      TabIndex        =   3
      Top             =   3810
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Reciept date:"
      Height          =   195
      Index           =   3
      Left            =   810
      TabIndex        =   2
      Top             =   3435
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ISBN:"
      Height          =   195
      Index           =   2
      Left            =   1350
      TabIndex        =   1
      Top             =   3075
      Width           =   420
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Title of Book:"
      Height          =   195
      Index           =   0
      Left            =   825
      TabIndex        =   0
      Top             =   2340
      Width           =   945
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuopenData 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "BookLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim LastRec As Integer
Dim BDB(0 To 499) As RNbookloan
Private Type RNbookloan
    Owner As String
    Add1 As String
    Add2 As String
    City As String
    State As String
    Zip As String
    Phone As String
    Title As String
    Author As String
    ISBN As String
    DateReciept As String
    DateReturn As String
End Type

Private Sub begining_Click()
    Text1(0).Text = BDB(0).Owner
    Text1(1).Text = BDB(0).Add1
    Text1(2).Text = BDB(0).Add2
    Text1(3).Text = BDB(0).City
    Text1(4).Text = BDB(0).State
    Text1(5).Text = BDB(0).Zip
    Text1(6).Text = BDB(0).Phone
    Text1(7).Text = BDB(0).Title
    Text1(8).Text = BDB(0).Author
    Text1(9).Text = BDB(0).ISBN
    Text1(10).Text = BDB(0).DateReciept
    Text1(11).Text = BDB(0).DateReturn
End Sub
Private Sub DeleteRec_Click()
Dim msg As String
    msg = MsgBox("               Delete record?", vbYesNo, "Delete Record")
End Sub

Private Sub ending_Click()
    Text1(0).Text = BDB(LastRec).Owner
    Text1(1).Text = BDB(LastRec).Add1
    Text1(2).Text = BDB(LastRec).Add2
    Text1(3).Text = BDB(LastRec).City
    Text1(4).Text = BDB(LastRec).State
    Text1(5).Text = BDB(LastRec).Zip
    Text1(6).Text = BDB(LastRec).Phone
    Text1(7).Text = BDB(LastRec).Title
    Text1(8).Text = BDB(LastRec).Author
    Text1(9).Text = BDB(LastRec).ISBN
    Text1(10).Text = BDB(LastRec).DateReciept
    Text1(11).Text = BDB(LastRec).DateReturn
End Sub

Private Sub ExitRec_Click()
    End
End Sub

Private Sub Form_Load()
LastRec = 0
Open App.Path & "\bookloan.dat" For Input As #1
    Do Until EOF(1)
        Input #1, BDB(LastRec).Owner, BDB(LastRec).Add1, BDB(LastRec).Add2, BDB(LastRec).City, BDB(LastRec).State, BDB(LastRec).Zip, BDB(LastRec).Phone, BDB(LastRec).Title, BDB(LastRec).Author, BDB(LastRec).ISBN, BDB(LastRec).DateReciept, BDB(LastRec).DateReturn
    LastRec = LastRec + 1
    Loop
Close #1
begining_Click
End Sub

