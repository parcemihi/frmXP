VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "TCP Client"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtOutput 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtSendData 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   960
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
' The name of the Winsock control is tcpClient.
' Note: to specify a remote host, you can use
' either the IP address (ex: "121.111.1.1") or
' the computer's "friendly" name, as shown here.
tcpClient.RemoteHost = "Gateway"
tcpClient.RemotePort = 1001
End Sub

Private Sub cmdConnect_Click()
' Invoke the Connect method to initiate a
' connection.
tcpClient.Connect
End Sub

Private Sub txtSendData_Change()
tcpClient.SendData txtSendData.Text
End Sub

Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
Dim strData As String
tcpClient.GetData strData
txtOutput.Text = strData
End Sub

