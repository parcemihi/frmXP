VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AddCheck()
    Dim stSQL As String
    Dim stDescription As String
    Dim dDebit As Double
    Dim stTransactionDate As String
    Dim stCheckNumber As String
    
    With frmCheck
        stDescription = .txtDescription
        dDebit = Val(.txtAmount)
        If Len(.txtTransactionDate) = 0 Then
            stTransactionDate = "NULL"
        Else
            stTransactionDate = "#" & .txtTransactionDate & "#"
        End If
        stCheckNumber = .txtCheckNumber
        
        stSQL = "insert into ledger"
        stSQL = stSQL & "("
        stSQL = stSQL & "Description"
        stSQL = stSQL & ",Debit"
        stSQL = stSQL & ",TransactionDate"
        stSQL = stSQL & ",CheckNumber"
        stSQL = stSQL & ")"
        stSQL = stSQL & " values"
        stSQL = stSQL & "("
        stSQL = stSQL & IsNVLString(stDescription)
        stSQL = stSQL & "," & IsNVLNumber(dDebit)
        stSQL = stSQL & "," & stTransactionDate
        stSQL = stSQL & "," & IsNVLString(stCheckNumber)
        stSQL = stSQL & ")"
        
        objAccessConnection.Execute stSQL
        
        
    End With

    MsgBox "Worked"
End Sub
Public Sub AddDeposit()
    Dim stSQL As String
    Dim stDescription As String
    Dim dCredit As Double
    Dim stTransactionDate As String
    
    With frmDeposit
        stDescription = .txtDescription
        dCredit = Val(.txtCredit)
        If Len(.txtTransactionDate) = 0 Then
            stTransactionDate = "NULL"
        Else
            stTransactionDate = "#" & .txtTransactionDate & "#"
        End If
        
        stSQL = "insert into ledger"
        stSQL = stSQL & "("
        stSQL = stSQL & "Description"
        stSQL = stSQL & ",Credit"
        stSQL = stSQL & ",TransactionDate"
        stSQL = stSQL & ")"
        stSQL = stSQL & " values"
        stSQL = stSQL & "("
        stSQL = stSQL & IsNVLString(stDescription)
        stSQL = stSQL & "," & IsNVLNumber(dCredit)
        stSQL = stSQL & "," & stTransactionDate
        stSQL = stSQL & ")"
        
        objAccessConnection.Execute stSQL
        
        
    End With

    MsgBox "Worked"
End Sub
