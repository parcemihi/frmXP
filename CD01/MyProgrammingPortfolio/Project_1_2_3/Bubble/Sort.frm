VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim uniqueNum() As Long


Private Sub Form_Load()
ReDim uniqueNum(0 To 100000)
Dim i As Integer
    Open App.Path & "\unique.num" For Input As #1
        Do Until EOF(1)
        i = i + 1
            Input #1, uniqueNum(i)
        Loop
    Close #1
ReDim Preserve uniqueNum(0 To i)
    

BubbleSortArray (uniqueNum())
    
i = i
    
    Open App.Path & "\bubble.num" For Output As #1
    Do While i > 0
    
    Write #1, uniqueNum(i)
    i = i - 1
    Loop
    Close #1






End Sub

'RETURNS Array, or vbEmpty if there's an error
'e.g., passed array contains elements that
'can't be compared to each other, such as
'objects
'will work when elements are all numbers or single characters
'of the same case.
Public Function BubbleSortArray(ByVal NumericArray As Variant) _
    As Variant
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

