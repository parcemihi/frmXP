Attribute VB_Name = "UserPass"
Option Explicit
Dim UserNPP(3) As String
Dim UName As String
Dim UPass As String
Dim UPermission
Dim x As Single, y As String, z As String
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim LoginUser As String
Dim LoginPass As String

Public Function Clear_UserNPP()
For i = 0 To 3
UserNPP(i) = ""
Next i
i = 0
End Function

Public Function Login_UserPass(username As String, userpass As String)
Clear_UserNPP
y = ""
j = 1
Open App.Path & "\userpass.dat" For Input As #1
    Do Until EOF(1)
    Input #1, y
        For i = 1 To Len(y)
        UserNPP(0) = UserNPP(0) & Chr((Asc(Mid(y, i, 1))) / 4)
        Next i
        x = x
        
        For i = 1 To Len(UserNPP(0)) Step 7
        
        If Chr((((Mid(UserNPP(0), i, 7)) / 100000) - 10) ^ 2) = "!" Then
        j = j + 1
        Else
        UserNPP(j) = UserNPP(j) & Chr((((Mid(UserNPP(0), i, 7)) / 100000) - 10) ^ 2)
        End If
        Next i
    
    If username = UserNPP(1) And userpass = UserNPP(2) Then
        Exit Do
    End If
    
    Loop
Close #1
End Function
Public Function CreateUser(username As String, userpass As String, UserPriv As Integer)
'//Encryption
z = ""
UserNPP(0) = username & "!" & userpass & "!" & UserPriv

Open App.Path & "\userpass.dat" For Append As #1
    For i = 1 To Len(UserNPP(0))
    x = (Sqr(Asc(Mid(UserNPP(0), i, 1))) + 10) * 100000
        For j = 1 To 7
        z = z & Chr((Asc(Mid(x, j, 1))) * 4)
        l = Len(z)
        
        Next j
    Next i
Write #1, z
Close #1
End Function






'    Input #1, UName, UPass, UPermission
'        For i = 1 To Len(UName)
'        y = y & Chr((Asc(Mid(UName, i, 1))) / 4)
'        Next i
'    UName = ""
'        For i = 1 To Len(y) Step 7
'        UName = UName & Chr((((Mid(y, i, 7)) / 100000) - 10) ^ 2)
'        Next i
'
'        For i = 1 To Len(UPass)
'        y = y & Chr((Asc(Mid(UPass, i, 1))) / 4)
'        Next i
'    UPass = ""
'        For i = 1 To Len(y) Step 7
'        UPass = UPass & Chr((((Mid(y, i, 7)) / 100000) - 10) ^ 2)
'        Next i
'
'        For i = 1 To Len(UPermission)
'        y = y & Chr((Asc(Mid(UPermission, i, 1))) / 4)
'        Next i
'    UPermission = ""
'        For i = 1 To Len(y) Step 7
'        UPermission = UPermission & Chr((((Mid(y, i, 7)) / 100000) - 10) ^ 2)
'        Next i








'Write #1, UName
'Close #1

'Open App.Path & "\checkthis.dat" For Append As #1
'    For i = 1 To Len(userpass)
'    x = (Sqr(Asc(Mid(userpass, i, 1))) + 10) * 100000
'        For j = 1 To 7
'        UPass = UPass & Chr((Asc(Mid(x, j, 1))) * 4)
'        Next j
'    Next i



'z = UName & "!" & UPass







'//Decryption
'y = ""
'j = 0
'Open App.Path & "\checkthis.dat" For Input As #1
'Do Until EOF(1)
'Input #1, y
'    For i = 1 To Len(y)
'        If Mid(y, i, 1) = "!" Then
'            z = ""
'            For k = 1 To Len(UserNPP(j)) Step 7
'            z = z & Chr((((Mid(UserNPP(j), k, 7)) / 100000) - 10) ^ 2)
'            Next k
'        UserNPP(j) = z
'        j = j + 1
'
'        ElseIf Mid(y, i, 1) = "#" Then
'        UserNPP(j) = UserNPP(j) & Chr((Asc(Mid(y, i, 1))) / 4)
'        Else
'
'        UserNPP(j) = UserNPP(j) & Chr((Asc(Mid(y, i, 1))) / 4)
'
'        End If
'    Next i
'    UName = ""
'    For i = 1 To Len(y) Step 7
'    UName = UName & Chr((((Mid(y, i, 7)) / 100000) - 10) ^ 2)
'    Next i
'Loop
'Close #1
'


