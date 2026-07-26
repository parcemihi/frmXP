VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmFiles 
   Caption         =   "File Name View"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "See &Files"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgFiles 
      Left            =   2160
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      DialogTitle     =   "File View"
      FileName        =   "Autoexec.bat"
      Filter          =   "*.bat;*.txt"
      FontSize        =   4.44119e-38
      InitDir         =   "c:\"
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  dlgFiles.ShowOpen
End Sub
