Attribute VB_Name = "acc_load"
Option Explicit
Dim accounts_ini(100) As AI
Private Type AI
    iProperty As String
    iValue As String
End Type
Dim INI_Entries As Integer
Dim Entry_Point As Integer

Public Sub LOAD_ACCOUNT_OPTIONS()
INI_Entries = -1
Open App.Path & "\accounts.ini" For Input As #1
    Do Until EOF(1)
        INI_Entries = INI_Entries + 1
        Input #1, accounts_ini(INI_Entries).iProperty
    Loop
Close #1
Format_Accounts_ini
Process_Accounts
End Sub

Private Sub Format_Accounts_ini()
Dim i As Integer, j As Integer
For i = 0 To INI_Entries
    Select Case Mid(accounts_ini(i).iProperty, 1, 1)
        Case "[" Or ";"
        Case Else
        For j = 1 To Len(accounts_ini(i).iProperty)
            If Mid(accounts_ini(i).iProperty, j, 1) = "=" Then
accounts_ini(i).iValue = Mid(accounts_ini(i).iProperty, j, Len(accounts_ini(i).iProperty))
accounts_ini(i).iProperty = Mid(accounts_ini(i).iProperty, 1, j)
                
            End If
        Next j
    End Select
Next i
End Sub
Private Sub Process_Accounts()
iniPointer = -1
Do Until iniPointer = iniEntries
iniPointer = iniPointer + 1
    Select Case accounts_ini(iniPointer)
        Case "[DATABASE]"
    
        Case "[OPTIONS]"
    
    End Select

Loop
End Sub




