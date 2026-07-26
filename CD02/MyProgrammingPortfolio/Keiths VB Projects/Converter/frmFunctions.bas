Attribute VB_Name = "frmFunctions"
Public pass As String

'----------------Decimal Conversion Section----------------
Public Function DEC2HEX(x)
    If Val(pass) = 0 Then
        pass = Hex(Val(pass))
        pass = 0
    Else
        pass = Hex(Val(pass))
    End If
End Function
Public Function DEC2BIN(x)
    Dim num, power, Num1, Num2
    num = Val(pass)
    If num = 0 Then
        pass = 0
        pass = 0
    Else
    power = Fix(Log2(num))  'sets highest power of 2 < num
    Num2 = ""
        Do While power >= 0
            If num - 2 ^ power >= 0 Then
                Num1 = 1
                num = num - 2 ^ power
            Else
                Num1 = 0
            End If
            power = power - 1
            Num2 = Num2 & Num1
        Loop
    pass = Num2
    End If
End Function
Public Function DEC2OCT(x)
    If Val(pass) = 0 Then
        pass = Oct(Val(pass))
        pass = 0
    Else
        pass = Oct(Val(pass))
    End If
End Function
Public Function DEC2DEC(x)
    MsgBin = MsgBox("--------------- DEC is DEC ---------------", vbOKOnly)
End Function

'--------------Hexidecimal Conversion Section--------------
Public Function HEX2HEX(x)
    MsgBin = MsgBox("--------------- HEX is HEX ---------------", vbOKOnly)
End Function
Public Function HEX2BIN(x)
    DEC2BIN (HEX2DEC(pass))
End Function
Public Function HEX2OCT(x)
    DEC2OCT (HEX2DEC(pass))
End Function
Public Function HEX2DEC(x)
    Dim alpha, power, Num1, Num2, hexi

    If pass = "" Then
    pass = 0
    End If
    
    power = 0
    alpha = Len(frmConvert.Text1)
    hexi = 0
    Num1 = 0
    Num2 = 0
    Do
        Do
            If Mid(pass, alpha, 1) = "a" Then
            hexi = 10
            ElseIf Mid(pass, alpha, 1) = "b" Then
            hexi = 11
            ElseIf Mid(pass, alpha, 1) = "c" Then
            hexi = 12
            ElseIf Mid(pass, alpha, 1) = "d" Then
            hexi = 13
            ElseIf Mid(pass, alpha, 1) = "e" Then
            hexi = 14
            ElseIf Mid(pass, alpha, 1) = "f" Then
            hexi = 15
            ElseIf Mid(pass, alpha, 1) = "A" Then
            hexi = 10
            ElseIf Mid(pass, alpha, 1) = "B" Then
            hexi = 11
            ElseIf Mid(pass, alpha, 1) = "C" Then
            hexi = 12
            ElseIf Mid(pass, alpha, 1) = "D" Then
            hexi = 13
            ElseIf Mid(pass, alpha, 1) = "E" Then
            hexi = 14
            ElseIf Mid(pass, alpha, 1) = "F" Then
            hexi = 15
            Else
            hexi = Mid(pass, alpha, 1)
            End If
            Num1 = hexi * 16 ^ power
            power = power + 1
            alpha = alpha - 1
            Exit Do
        Loop
    Num2 = Num1 + Num2
    Loop Until alpha = 0
    pass = Num2
End Function

'-----------------Octal Conversion Section-----------------
Public Function OCT2HEX(x)
    DEC2HEX (OCT2DEC(pass))
End Function
Public Function OCT2BIN(x)
    DEC2BIN (OCT2DEC(pass))
End Function
Public Function OCT2OCT(x)
MsgBin = MsgBox("--------------- OCT is OCT ---------------", vbOKOnly)
End Function
Public Function OCT2DEC(x)
    Dim count As Integer, length As Integer, num As Integer, eight As Integer, power As Integer
    length = Len(pass)
    eight = 0
    power = -1
        Do
        power = power + 1
        num = Mid(pass, length, 1)
        eight = 8 ^ power
        count = count + num * eight
        length = length - 1
        Loop Until length = 0
    pass = count
End Function

'----------------Bianary Conversion Section----------------
Public Function BIN2HEX(x)
    DEC2HEX (BIN2DEC(pass))
End Function
Public Function BIN2BIN(x)
    MsgBin = MsgBox("--------------- BIN is BIN ---------------", vbOKOnly)
End Function
Public Function BIN2OCT(x)
    DEC2OCT (BIN2DEC(pass))
End Function
Public Function BIN2DEC(x)
    Dim power, posi, count, num
    power = Len(pass) - 1
    num = 0
    Do
        count = count + 1
        If Mid(pass, count, 1) = 1 Then
        num = 2 ^ power + num
        power = power - 1
        Else
        power = power - 1
        End If
    Loop Until power = -1
    pass = num
End Function

'------------Miscellaneous Subs and Functions------------

Public Function Log2(x)
    Log2 = Log(x) / Log(2#)
End Function
Public Function TXT1(x)
    TXT1 = frmConvert.Text1
End Function
Public Function TXT2(x)
    TXT2 = frmConvert.Text2
End Function
