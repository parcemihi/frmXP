VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EmailAddress(0 To 1000) As String * 40

Private Sub Command1_Click()
Dim EmailField As String * 14
Dim TotalEmails As Integer
Dim EmailLine As String
Dim TheAT As Integer
ChDir App.Path
TotalEmails = -1
EmailField = "E-mail Address"
CommonDialog1.ShowOpen

Open CommonDialog1.FileName For Input As #1
    Do Until EOF(1)
        TotalEmails = TotalEmails + 1
        
        Input #1, EmailAddress(TotalEmails)
        
        
        
        
    Loop
Close #1

Open App.Path & "\test.dat" For Output As #1
    
    
    
    
    
Close #1





CommonDialog1.ShowSave

End Sub

