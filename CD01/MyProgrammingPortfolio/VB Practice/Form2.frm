VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function BubbleSortArray(ByVal NumericArray As Variant) As Variant
'RETURNS Array, or vbEmpty if there's an error
'e.g., passed array contains elements that
'can't be compared to each other, such as
'objects
'will work when elements are all numbers or single characters
'of the same case.
Dim vAns As Variant
Dim vTemp As Variant
Dim bSorted As Boolean
Dim lCtr As Long
Dim lCount As Long
Dim lStart As Long
vAns = NumericArray
If Not IsArray(vAns) Then
    BubbleSortArray = vbEmpty
    Exit Function
End If
On Error GoTo ErrorHandler
lStart = LBound(vAns)
lCount = UBound(vAns)
    bSorted = False
    Do While Not bSorted
      bSorted = True
      For lCtr = lCount - 1 To lStart Step -1
        If vAns(lCtr + 1) < vAns(lCtr) Then
          DoEvents
          bSorted = False
           vTemp = vAns(lCtr)
           vAns(lCtr) = vAns(lCtr + 1)
           vAns(lCtr + 1) = vTemp
         End If
      Next lCtr
    Loop
BubbleSortArray = vAns
Exit Function
ErrorHandler:
BubbleSortArray = vbEmpty
Exit Function
End Function
