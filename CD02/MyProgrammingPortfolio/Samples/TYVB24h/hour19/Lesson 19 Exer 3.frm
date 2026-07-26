VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form frmHappy 
   Caption         =   "Happy Sad Face"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      BeginProperty Buttons {7791BA41-E020-11CF-8E74-00A0C90F26F8} 
         NumButtons      =   2
         BeginProperty Button1 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Smile"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {7791BA43-E020-11CF-8E74-00A0C90F26F8} 
            Key             =   "Sad"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Shape shpTear 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   4200
      Shape           =   2  'Oval
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpWink 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4080
      Top             =   1800
      Width           =   255
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      BeginProperty Images {8556BCD1-E01E-11CF-8E74-00A0C90F26F8} 
         NumListImages   =   2
         BeginProperty ListImage1 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Lesson 19 Exer 3.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {8556BCD3-E01E-11CF-8E74-00A0C90F26F8} 
            Picture         =   "Lesson 19 Exer 3.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line lneSmile 
      BorderWidth     =   10
      Index           =   1
      X1              =   3000
      X2              =   3480
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line lneSmile 
      BorderWidth     =   10
      Index           =   0
      X1              =   4080
      X2              =   4560
      Y1              =   2760
      Y2              =   2640
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Shape shpEye 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   4080
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape shpEye 
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2655
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   960
      Width           =   2655
   End
End
Attribute VB_Name = "frmHappy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
  If Button.Key = "Smile" Then
         ' Add corners of mouth
         ' Wink eye
    lneSmile(0).Visible = True
    lneSmile(1).Visible = True
    shpWink.Visible = True
    shpEye(1).Visible = False
    shpTear.Visible = False
  Else   ' Sad was clicked
         ' Remove corners of mouth
         ' Add tear
    lneSmile(0).Visible = False
    lneSmile(1).Visible = False
    shpWink.Visible = False
    shpEye(1).Visible = True
    shpTear.Visible = True
    
  End If
  
End Sub
