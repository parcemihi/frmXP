VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu Application"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuWrite 
      Caption         =   "Write"
      Begin VB.Menu mnuWriteKeyboard 
         Caption         =   "Keyboard"
      End
      Begin VB.Menu mnuWritePencil 
         Caption         =   "Pencil"
      End
      Begin VB.Menu mnuWritePen 
         Caption         =   "Pen"
      End
   End
   Begin VB.Menu mnuRead 
      Caption         =   "Read"
      Begin VB.Menu mnuReadScreen 
         Caption         =   "Screen"
      End
      Begin VB.Menu mnuReadBook 
         Caption         =   "Book"
      End
      Begin VB.Menu mnuReadMagazine 
         Caption         =   "Magazine"
      End
   End
   Begin VB.Menu mnuListen 
      Caption         =   "Listen"
      Begin VB.Menu mnuListenRadio 
         Caption         =   "Radio"
      End
      Begin VB.Menu mnuListenTelevision 
         Caption         =   "Television"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
