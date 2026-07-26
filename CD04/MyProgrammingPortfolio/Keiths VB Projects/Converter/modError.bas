Attribute VB_Name = "modError"
Option Explicit

Public Function errorCheck()
Dim msgOct As String, msgBin As String
Dim msgDec As String, msgHex As String
Dim error As String, message As String
Dim check As String, length As Integer
Dim alpha, i
alpha = Array("a", "b", "c", "d", "e", "f", _
              "A", "B", "C", "D", "E", "F", _
              "1", "2", "3", "4", "5", "6", _
              "7", "8", "9", "0")
error = "Error check results"
msgBin = "That is not a valid Bianary number"
msgHex = "That is not a valid Hexidecimal number"
msgOct = "That is not a valid Octal number"
msgDec = "That is not a valid Decimal number"

If frmConvert.Opt1Dec = True Then
    length = Len(pass)
    Do Until length = 0
    check = Mid(pass, length, 1)
    length = length - 1
        Select Case check
        Case 0 To 9
        errorCheck = 1
        Case Else
        errorCheck = 0
        message = MsgBox(msgDec, vbCritical, error)
        Exit Do
        End Select
    Loop
End If

If frmConvert.Opt2Hex = True Then
    length = Len(pass)
    Do Until length = 0
    errorCheck = 0
    check = Mid(pass, length, 1)
    length = length - 1
        For i = 0 To 21
            If alpha(i) = check Then
                errorCheck = 1
            Exit For
            End If
        Next i
        If errorCheck = 0 Then
            message = MsgBox(msgHex, vbCritical, error)
            Exit Do
        End If
    Loop
End If

If frmConvert.Opt3Bin = True Then
    length = Len(pass)
    Do Until length = 0
    check = Mid(pass, length, 1)
    length = length - 1
        Select Case check
        Case 0 To 1
        errorCheck = 1
        Case Else
        errorCheck = 0
        message = MsgBox(msgBin, vbCritical, error)
        Exit Do
        End Select
    Loop
End If

If frmConvert.Opt4Oct = True Then
    length = Len(pass)
    Do Until length = 0
    check = Mid(pass, length, 1)
    length = length - 1
        Select Case check
        Case 0 To 7
        errorCheck = 1
        Case Else
        errorCheck = 0
        message = MsgBox(msgOct, vbCritical, error)
        Exit Do
        End Select
    Loop
End If
End Function
