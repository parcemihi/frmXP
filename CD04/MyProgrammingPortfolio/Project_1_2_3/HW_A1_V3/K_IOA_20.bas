Attribute VB_Name = "IOA"
Option Explicit
Global BN As Integer
Dim x As Integer
Global kdc As KDataCrdntr20
Type KDataCrdntr20
    Fields(0 To 2500) As Integer
    Index_(0 To 2500) As String
    bank(0 To 9, 0 To 2500) As Integer
    price(0 To 9) As Currency
End Type
Public Function Fins(Path_Filename As String)
BN = -1
On Error GoTo ErrorHandler
Open Path_Filename For Input As #1
    Do Until EOF(1)
    BN = BN + 1
    Input #1, kdc.Fields(BN)
    Select Case kdc.Fields(BN)
        Case 1:  Input #1, kdc.Index_(BN), kdc.bank(0, BN)
        Case 2:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN)
        Case 3:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN)
        Case 4:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN)
        Case 5:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN)
        Case 6:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN)
        Case 7:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN)
        Case 8:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN), kdc.bank(7, BN)
        Case 9:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN), kdc.bank(7, BN), kdc.bank(8, BN)
        Case 10: Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN), kdc.bank(7, BN), kdc.bank(8, BN), kdc.bank(9, BN)
    End Select
Loop
Close #1
Exit Function
ErrorHandler:
    errNumber.ErrorHandler (Err.Number)
End Function
Public Function Fouts(Path_Filename As String)
Open App.Path & "\" & Path_Filename For Output As #2
For x = 0 To Form1.Combo1.ListCount - 1
Select Case kdc.Fields(x)
    Case 0:  Write #2, msg62
    Case 1:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x)
    Case 2:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x)
    Case 3:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x)
    Case 4:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x)
    Case 5:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x)
    Case 6:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x)
    Case 7:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x)
    Case 8:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x), kdc.bank(7, x)
    Case 9:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x), kdc.bank(7, x), kdc.bank(8, x)
    Case 10: Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x), kdc.bank(7, x), kdc.bank(8, x), kdc.bank(9, x)
End Select
Next x
Close #2
End Function
Public Function Fpurch(Path_Filename As String)
x = Form1.Combo1.ListIndex
Open App.Path & "\" & Path_Filename & ".pur" For Output As #2
Select Case kdc.Fields(x)
    Case 0:  Write #2, msg62
    Case 1:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x)
    Case 2:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x)
    Case 3:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x)
    Case 4:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x)
    Case 5:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x)
    Case 6:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x)
    Case 7:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x)
    Case 8:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x), kdc.bank(7, x)
    Case 9:  Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x), kdc.bank(7, x), kdc.bank(8, x)
    Case 10: Write #2, kdc.Fields(x), kdc.Index_(x), kdc.bank(0, x), kdc.bank(1, x), kdc.bank(2, x), kdc.bank(3, x), kdc.bank(4, x), kdc.bank(5, x), kdc.bank(6, x), kdc.bank(7, x), kdc.bank(8, x), kdc.bank(9, x)
End Select
Close #2
End Function

