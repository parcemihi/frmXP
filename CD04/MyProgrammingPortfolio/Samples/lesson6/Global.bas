Attribute VB_Name = "basQuicken"
Option Explicit
Public objAccessConnection As ADODB.Connection
Public gsFileName As String
Public gErrFormName As String
Public gsPath As String
Public gsDrive As String
Public gsFHD_APPL_User As String
Public gsFHD_APPL_Password As String
Public gsSA_User As String
Public gsSA_Password As String


Sub HandleError(strLoc As String, strError$, lError As Long, varModule As Variant)

    Dim nCursorType As Integer

    nCursorType = Screen.MousePointer

    Screen.MousePointer = vbNormal
    MsgBox strLoc & ": " & strError & " (" & lError & ")", vbExclamation, varModule
    Screen.MousePointer = nCursorType

End Sub

Public Function FixApostrophy(ByVal sSQL As String) As String
Dim sFront$, sBack$, nParamLen%
Dim sPhrase As String
Dim wLength As Integer
Dim i As Integer
On Error GoTo FixApostrophy_Error

    wLength = Len(sSQL)
    For i = 1 To wLength
        If Mid$(sSQL, i, 1) = "'" Then
            sPhrase = sPhrase + "''"
        Else
            sPhrase = sPhrase + Mid$(sSQL, i, 1)
        End If
    Next
    FixApostrophy = sPhrase
    
Exit_FixApostrophy:

Exit Function
FixApostrophy_Error:
    #If gnDebug Then
        Stop
        Resume
    #End If
    'Standard error handling statement
    HandleError "FixApostrophy", Err.Description, Err.Number, gErrFormName
    Resume Exit_FixApostrophy

End Function

Public Function IsNvlYesNo(vValue As Variant) As Variant
    If IsNull(vValue) Then
        IsNvlYesNo = Null
    ElseIf vValue = False Then
        IsNvlYesNo = 0
    ElseIf vValue = True Then
        IsNvlYesNo = -1
    End If
    
End Function

Public Function IsNVLString(parameter As Variant) As Variant
    On Error GoTo IsNVLString_Error
    
    If IsNull(parameter) Or parameter = "" Then
        IsNVLString = "Null"
        GoTo Exit_IsNVLString
    End If
            
    IsNVLString = "'" & FixApostrophy(parameter) & "'"
    
    
Exit_IsNVLString:
   Exit Function
IsNVLString_Error:
    
    #If gnDebug Then
        Stop
        Resume
    #End If

    HandleError "IsNVLString", Err.Description, Err.Number, gErrFormName
    Resume Exit_IsNVLString
    
End Function
Public Function IsNVLNumber(parameter As Variant) As Variant
    On Error GoTo IsNVLNumber_Error
    
    If IsNull(parameter) Then
        IsNVLNumber = "Null"
        GoTo Exit_IsNVLNumber
    End If
            
    IsNVLNumber = parameter
    
    
Exit_IsNVLNumber:
   Exit Function
IsNVLNumber_Error:
    
    #If gnDebug Then
        Stop
        Resume
    #End If

    HandleError "IsNVLNumber", Err.Description, Err.Number, gErrFormName
    Resume Exit_IsNVLNumber
    
End Function
Public Function AppPath() As String
    
    Dim sAns As String
    sAns = App.Path
    If Right(App.Path, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns

End Function
