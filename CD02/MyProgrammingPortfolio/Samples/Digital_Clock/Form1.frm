VERSION 5.00
Object = "{B66B514D-9AED-11D4-89D8-F0E1FDCFB023}#3.0#0"; "Clock.ocx"
Begin VB.Form Form1 
   Caption         =   "Clock Demo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Clock.ctlClock ctlClock1 
      Height          =   330
      Left            =   1140
      TabIndex        =   0
      Top             =   1290
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   582
      SetAlarmHr      =   8
      SetAlarmMin     =   17
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ctlClock1_Alarm(ThisTime As Date)
    MsgBox "Now it is " & Format$(ThisTime, "hh:mm"), vbOKOnly, "Alarm"
End Sub

Private Sub Form_Load()
    ctlClock1.Enabled = True
End Sub
