Attribute VB_Name = "sort"
Option Explicit
Global Total
Global ArrayPointer As Integer
Global EndOf_Array As Integer
Global EndOfCredit_Array As Integer
Global DataBoxDirect As Boolean
Global Entry() As Bank_Account
Type Bank_Account
    Number As Integer
    Date As Date
    Description As String
    Amount As Currency
    cleared As Integer
    taxable As Integer
    EntType As Integer
End Type
Function SORT_RECORD(Index As Integer, Optional SortUp As Boolean)
Dim i As Integer, tmpLo As Integer, tmpHi As Integer, Hi As Integer, Lo As Integer
Dim IsSorted As Boolean
Lo = 0: Hi = EndOf_Array
Do While Lo <= Hi
IsSorted = True: tmpLo = Lo: tmpHi = Hi
If SortUp = True Then
For i = Lo To Hi - 1
Select Case Index
    Case Is = 0
        If Entry(i).Number > Entry(tmpHi).Number Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 1
        If Entry(i).Date > Entry(tmpHi).Date Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 2
        If Entry(i).Description > Entry(tmpHi).Description Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 3
        If Entry(i).Amount > Entry(tmpHi).Amount Then
            tmpHi = i: IsSorted = False
        End If
End Select
Next i
If IsSorted = False Then
    SWAP tmpHi, Hi
End If
Hi = Hi - 1
For i = Lo To Hi
Select Case Index
    Case Is = 0
        If Entry(i).Number < Entry(tmpLo).Number Then
            tmpLo = i: IsSorted = False
        End If
    Case Is = 1
        If Entry(i).Date < Entry(tmpLo).Date Then
            tmpLo = i: IsSorted = False
        End If
    Case Is = 2
        If Entry(i).Description < Entry(tmpLo).Description Then
            tmpLo = i: IsSorted = False
        End If
    Case Is = 3
        If Entry(i).Amount < Entry(tmpLo).Amount Then
            tmpLo = i: IsSorted = False
        End If
End Select
Next i
If IsSorted = False Then
    SWAP tmpLo, Lo
End If
Lo = Lo + 1

Else
For i = Lo To Hi - 1
Select Case Index
    Case Is = 0
        If Entry(i).Number < Entry(tmpHi).Number Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 1
        If Entry(i).Date < Entry(tmpHi).Date Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 2
        If Entry(i).Description < Entry(tmpHi).Description Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 3
        If Entry(i).Amount < Entry(tmpHi).Amount Then
            tmpHi = i: IsSorted = False
        End If
    Case Is = 4
        If Entry(i).EntType = 0 Then
            tmpHi = i: IsSorted = False
        End If

End Select
Next i
If IsSorted = False Then
    SWAP tmpHi, Hi
End If
Hi = Hi - 1
For i = Lo To Hi
Select Case Index
    Case Is = 0
        If Entry(i).Number > Entry(tmpLo).Number Then
            tmpLo = i: IsSorted = False
        End If
    Case Is = 1
        If Entry(i).Date > Entry(tmpLo).Date Then
            tmpLo = i: IsSorted = False
        End If
    Case Is = 2
        If Entry(i).Description > Entry(tmpLo).Description Then
            tmpLo = i: IsSorted = False
        End If
    Case Is = 3
        If Entry(i).Amount > Entry(tmpLo).Amount Then
            tmpLo = i: IsSorted = False
        End If
End Select
Next i
If IsSorted = False Then
    SWAP tmpLo, Lo
End If
Lo = Lo + 1
End If
Loop
End Function
Function SWAP(First As Integer, Last As Integer)
Dim tmp1 As Variant, tmp2 As Variant, tmp3 As Variant, tmp4 As Variant, tmp5 As Variant, tmp6 As Variant
tmp1 = Entry(First).Amount
tmp2 = Entry(First).cleared
tmp3 = Entry(First).Date
tmp4 = Entry(First).Description
tmp5 = Entry(First).Number
tmp6 = Entry(First).taxable

     Entry(First).Amount = Entry(Last).Amount
    Entry(First).cleared = Entry(Last).cleared
       Entry(First).Date = Entry(Last).Date
Entry(First).Description = Entry(Last).Description
     Entry(First).Number = Entry(Last).Number
    Entry(First).taxable = Entry(Last).taxable

Entry(Last).Amount = tmp1
Entry(Last).cleared = tmp2
Entry(Last).Date = tmp3
Entry(Last).Description = tmp4
Entry(Last).Number = tmp5
Entry(Last).taxable = tmp6

End Function

Function DELETE_RECORD()
With Entry(ArrayPointer)
     .Amount = Entry(EndOf_Array).Amount
    .cleared = Entry(EndOf_Array).cleared
       .Date = Entry(EndOf_Array).Date
.Description = Entry(EndOf_Array).Description
     .Number = Entry(EndOf_Array).Number
    .taxable = Entry(EndOf_Array).taxable
End With
Entry(EndOf_Array).Amount = 0
Entry(EndOf_Array).cleared = 0
Entry(EndOf_Array).Date = 0
Entry(EndOf_Array).Description = ""
Entry(EndOf_Array).Number = 0
Entry(EndOf_Array).taxable = 0
EndOf_Array = EndOf_Array - 1
SORT_RECORD 0, True
End Function

