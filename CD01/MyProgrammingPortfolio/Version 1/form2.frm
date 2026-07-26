VERSION 5.00
Begin VB.Form form2 
   Caption         =   "Form2"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type CharRange
    firstChar As Long         ' First character of range (0 for start of doc)
    lastChar As Long          ' Last character of range (-1 for end of doc)
End Type

Private Type FormatRange
    hdc As Long               ' Actual DC to draw on
    hdcTarget As Long         ' Target DC for determining text formatting
    rectRegion As Rect        ' Region of the DC to draw to (in twips)
    rectPage As Rect          ' Page size of the entire DC (in twips)
    mCharRange As CharRange   ' Range of text to draw (see above user type)
End Type


Private Sub Form_Load()
Print ""
CharRange(firstChar) = 1

Print ""
Print ""




End Sub
