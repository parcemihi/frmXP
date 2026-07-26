VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoadRecord 
      Caption         =   "Load Records"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeleteRecord 
      Caption         =   "Delete Record"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton cmdSortRecord 
      Caption         =   "Sort Record"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   720
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   7575
   End
   Begin VB.CommandButton cmdAddRecord 
      Caption         =   "Add Record"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type Record   ' Define user-defined type.
   ID As Integer
   Name As String * 20
End Type



Private Sub cmdAddRecord_Click()


Dim MyRecord As Record   ' Declare variable.

 ID = Text1.Text
'Name = Text2
Open App.Path & "\testpat.dat" For Random As #1 Len = Len(MyRecord)

    Write #1, ID, Name




Close #1

End Sub

