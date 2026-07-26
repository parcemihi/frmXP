Attribute VB_Name = "GetShape"
Option Explicit

Public Sub genColor()
Randomize

rndom = Rnd
rndom = Int(rndom * 700)

Select Case rndom
    Case 0 To 700
    i1 = 5
    i2 = 15
    i3 = 25
    i4 = 35
'    Timer1_Timer
    Case 100 To 199
    Case 200 To 299
    Case 300 To 399
    Case 400 To 499
    Case 500 To 599
    Case 600 To 700
End Select

End Sub


