VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FF0000&
      Caption         =   "Change &Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FF0000&
      Caption         =   "Change &Height"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FF0000&
      Caption         =   "Change &Bold"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdChange 
      BackColor       =   &H00FF0000&
      Caption         =   "Change &Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click(Index As Integer)
  ' Changes ONLY the button clicked!
  Select Case Index
      Case 0:
      If cmdChange.Item(Index).BackColor = vbRed Then
         cmdChange.Item(Index).BackColor = vbBlue
      Else
         cmdChange.Item(Index).BackColor = vbRed
      End If
    Case 1:
      If cmdChange.Item(Index).FontBold = True Then
         cmdChange.Item(Index).FontBold = False
      Else
         cmdChange.Item(Index).FontBold = True
      End If
    Case 2:
      If cmdChange.Item(Index).Height = 495 Then
         cmdChange.Item(Index).Height = 595
      Else
         cmdChange.Item(Index).Height = 495
      End If
    Case 3:
      If cmdChange.Item(Index).Width = 1215 Then
         cmdChange.Item(Index).Width = 2000
      Else
         cmdChange.Item(Index).Width = 1215
      End If
    End Select
   
End Sub
