Attribute VB_Name = "banking"
Option Explicit
Dim Transaction As trans

Type trans
    ID As Integer
    Debit As Boolean
    Credit As Boolean



End Type







Function LoadAccountInfo()
Open App.Path & "\account.dat" For Input As #1


Close #1
End Function
