Attribute VB_Name = "mod_ErrorCheck"
Public Function ErrorCheckBin() As Integer
Dim length, start
length = Len(frmConvertBDHO.Text1.Text)
start = 0
Do
    start = start + 1
        If Mid(frmConvertBDHO.Text1, start, 1) > 1 Then
        intPress = MsgBox("That is not a bianary number", vbCritical)
        frmConvertBDHO.Text1.SetFocus
        ErrorCheck = 1
        Exit Do
        End If
Loop Until start = length
'If Len(frmConvertBDHO.Text1.Text) > 8 Then
'    intPress = MsgBox("BDHO only converts up to 8 digets bianary", vbCritical)
'    frmConvertBDHO.Text1.SetFocus
'    ErrorCheck = 1
'    Exit Function
'End If
ErrorCheck = 0

End Function
Public Function ErrorCheckHex()

Dim length, start
length = Len(frmConvertBDHO.Text1.Text)
start = 0
Do
    start = start + 1
        If Mid(frmConvertBDHO.Text1, start, 1) > 1 Then
        intPress = MsgBox("That is not a bianary number", vbCritical)
        frmConvertBDHO.Text1.SetFocus
        ErrorCheck = 1
        Exit Do
        End If
Loop Until start = length
If Len(frmConvertBDHO.Text1.Text) > 2 Then
    intPress = MsgBox("BDHO only converts up to 8 digets bianary", vbCritical)
    frmConvertBDHO.Text1.SetFocus
    ErrorCheck = 1
    Exit Function
End If
ErrorCheck = 0


End Function
Public Function ErrorCheckOct()

End Function
Public Function ErrorCheckDec()

End Function
