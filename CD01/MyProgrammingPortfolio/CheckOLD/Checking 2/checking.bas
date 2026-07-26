Attribute VB_Name = "checking"
Option Explicit
Global CheckIndx As Integer
Global DepositIndx As Integer
Global totChks As Integer
Global totDeps As Integer

Global Checks(0 To 5000) As WChk
Public Type WChk
    c1Numb As Integer
    c2Spec As Integer
    c3Date As Date
    c4Desr As String
    c5Amnt As Currency
End Type
Global Deposits(0 To 5000) As DChk
Public Type DChk
    d1Indx As String
    d2Date As String
    d3Desr As String
    d4Amnt As String
End Type

