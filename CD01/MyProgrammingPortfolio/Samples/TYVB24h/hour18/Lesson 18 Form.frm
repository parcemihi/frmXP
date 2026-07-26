VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Animated Cartoon"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAni 
      Interval        =   500
      Left            =   2280
      Top             =   3000
   End
   Begin VB.Image imgHappy 
      Height          =   1685
      Left            =   2280
      Picture         =   "Lesson 18 Form.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  ' Adjust the image's location
  imgHappy.Left = 0   ' Number of twips from
                      ' left of Form window
                      
  imgHappy.Top = 3820 ' Number of twips from
                      ' top of Form window
    
End Sub

Private Sub tmrAni_Timer()
  ' Adjust the Left and Top properties
  ' as well as the happy face shown so
  ' that the face appears to float up
  ' and across the Form window.
  
  ' The first time you declare a Satic Bollean
  ' variable, VB initializes it to false
  Static blnFace As Boolean
  
    ' Add to Left and Top only if room is left
  If (imgHappy.Left < 4800) And _
     (imgHappy.Top > 500) Then
       imgHappy.Left = imgHappy.Left + 100
       imgHappy.Top = imgHappy.Top - 50
  Else
    imgHappy.Left = 0    ' Restore image's first
    imgHappy.Top = 3820  ' position.
  End If
  
  ' Change the image displayed
  
  ' You may need to edit the graphic paths you see below to
  ' match the graphic file locations of your VB installation.
  ' In most cases, simply changing the drive letter from K:\...
  ' to C:\... (or whatever drive VB is installed) will do.
  
  If blnFace = True Then
     imgHappy.Picture = LoadPicture("K:\Program Files\Microsoft Visual Studio\" & _
                                    "Common\Graphics\Icons\Misc\Face03.ico")
     blnFace = False
  Else
     imgHappy.Picture = LoadPicture("K:\Program Files\Microsoft Visual Studio\" & _
                                    "Common\Graphics\Icons\Misc\Face02.ico")
     blnFace = True
  End If
     

End Sub
