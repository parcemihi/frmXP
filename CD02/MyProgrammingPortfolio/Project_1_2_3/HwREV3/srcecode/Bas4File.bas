Attribute VB_Name = "File"
Option Explicit
Global filename As String
Global dod As DeclareOutputData
Global DID As DeclareInputData
    Public Type DeclareInputData
    Fins(0 To 10) As Variant
    End Type
    Public Type DeclareOutputData
    Fouts(0 To 10) As Variant
    End Type

'//Create up to ten fields automatically
'//Valid field names are Fouts(1 to 10). Fouts(0) = Number of fields.
Public Function Create_Entry(Fouts As Integer)
Open App.Path & filename For Output As #1
    Select Case dod.Fouts(0)
    Case 0
    Write #1,
    Case 1
    Write #1, dod.Fouts(1)
    Case 2
    Write #1, dod.Fouts(1), dod.Fouts(2)
    Case 3
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3)
    Case 4
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4)
    Case 5
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5)
    Case 6
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6)
    Case 7
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7)
    Case 8
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8)
    Case 9
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8), dod.Fouts(9)
    Case 10
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8), dod.Fouts(9), dod.Fouts(10)
    Close #1

End Select
End Function

Public Function Append_Entry()
Open App.Path & filename For Append As #1
    Select Case dod.Fouts(0)
    Case 1
    Write #1, dod.Fouts(1)
    Case 2
    Write #1, dod.Fouts(1), dod.Fouts(2)
    Case 3
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3)
    Case 4
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4)
    Case 5
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5)
    Case 6
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6)
    Case 7
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7)
    Case 8
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8)
    Case 9
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8), dod.Fouts(9)
    Case 10
    Write #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8), dod.Fouts(9), dod.Fouts(10)
    Close #1
End Select
End Function



Public Function Input_Entry()
Open App.Path & filename For Input As #1
    Select Case dod.Fouts(0)
    Case 1
    Input #1, dod.Fouts(1)
    Case 2
    Input #1, dod.Fouts(1), dod.Fouts(2)
    Case 3
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3)
    Case 4
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4)
    Case 5
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5)
    Case 6
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6)
    Case 7
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7)
    Case 8
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8)
    Case 9
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8), dod.Fouts(9)
    Case 10
    Input #1, dod.Fouts(1), dod.Fouts(2), dod.Fouts(3), dod.Fouts(4), dod.Fouts(5), dod.Fouts(6), dod.Fouts(7), dod.Fouts(8), dod.Fouts(9), dod.Fouts(10)
    Close #1
End Select
End Function

