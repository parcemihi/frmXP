Attribute VB_Name = "basBanking"
Option Explicit
Dim AccountNumber As String
Global TotalChecks As Integer, CurrentCheck As Integer
Global TotalDeposits As Integer, CurrentDeposit As Integer
Global Checking(0 To 5000) As chkChecking
Public Type chkChecking
    Number As Integer
    Offset As Integer
    Date As String * 8
    Description As String * 40
    Amount As Single
End Type

Global Deposits(0 To 5000) As depChecking
Public Type depChecking
    DepID As Integer
    Date As String * 8
    Description As String * 40
    Amount As String * 11
End Type


Public Function SortDataArray(TheArray As String)
Select Case TheArray
Case "Checking"
        SortChecking
Case "Deposits"
End Select
End Function

Public Sub SortChecking()
Dim FileIsSorted As Boolean
FileIsSorted = False
Dim Lcnt As Integer
Dim Ucnt As Integer

Do Until FileIsSorted = True
    FileIsSorted = True
Loop
End Sub
Public Sub Input_Data()

AccountNumber = "03580946"
TotalChecks = -1
Open App.Path & "\" & AccountNumber & ".chk" For Input As #1
    Do Until EOF(1)
        TotalChecks = TotalChecks + 1
        With Checking(TotalChecks)
            Input #1, .Number, _
                      .Offset, _
                      .Date, _
                      .Description, _
                      .Amount
        End With
    Loop
Close #1


UpdateList1
End Sub

Public Sub UpdateList1()
Dim i As Integer, j As Integer
Dim ChkDollar As String
Dim TotalDollars As Double
Banking.List1.Clear
Banking.List1.AddItem "Number  Date      Description                             Amount       "
Banking.List1.AddItem "———————————————————————————————————————————————————————————————————————"

For i = 0 To TotalChecks
    ChkDollar = Format(Checking(i).Amount, "##,##0.00")

    For j = 1 To 10 - Len(ChkDollar)
        ChkDollar = " " & ChkDollar
    Next j
    ChkDollar = "$" & ChkDollar

With Checking(i)
Banking.List1.AddItem .Number & "-" & .Offset & "  " & .Date & "  " & .Description & ChkDollar
TotalDollars = TotalDollars + ChkDollar
End With
Next i
ChkDollar = Format(TotalDollars, "##,##0.00")
    For j = 1 To 10 - Len(ChkDollar)
        ChkDollar = " " & ChkDollar
    Next j

Banking.List1.AddItem "                                                          —————————————"
Banking.List1.AddItem "                                                    Total $" & ChkDollar

End Sub


Public Sub Output_Data()
AccountNumber = "03580946"
Dim i As Integer
Open App.Path & "\" & AccountNumber & ".chk" For Output As #1
    For i = 0 To TotalChecks
        With Checking(i)
            Write #1, .Number, _
                      .Offset, _
                      .Date, _
                      .Description, _
                      .Amount
        End With
    Next i
Close #1
UpdateList1
End Sub

