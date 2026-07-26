Attribute VB_Name = "checking"
Global EntryPointer As Integer
Global TotalEntries As Integer
Global account(0 To 5000) As Checking_Account
Global DBstart As String

Public Type Checking_Account
    ENT0_CVD_TYPE As Integer
    ENT1_CLEARED As Integer
    ENT2_NUM As Integer
    ENT3_FAMILY As Integer
    ENT4_MONTH As Integer
    ENT5_DAY As Integer
    ENT6_YEAR As Integer
    ENT7_DESCRIPTION As String * 20
    ENT8_AMOUNT As Single
End Type
