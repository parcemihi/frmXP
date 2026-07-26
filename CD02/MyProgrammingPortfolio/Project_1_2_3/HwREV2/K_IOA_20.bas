Attribute VB_Name = "IOA"
Option Explicit
Global BN As Integer
Global kdc As KDataCrdntr20
Type KDataCrdntr20
    Fields(0 To 1487) As Integer
    Index_(0 To 1487) As Integer
    bank(0 To 9, 0 To 1487) As Single
End Type

Public Function Fins(Path_Filename As String)
BN = -1
Open Path_Filename For Input As #1
Do Until EOF(1)
BN = BN + 1
Input #1, kdc.Fields(BN)
Select Case kdc.Fields(BN)
    Case 1:  Input #1, kdc.Index_(BN), kdc.bank(0, BN)
    Case 2:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN)
    Case 3:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN)
    Case 4:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN)
    Case 5:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN)
    Case 6:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN)
    Case 7:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN)
    Case 8:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN), kdc.bank(7, BN)
    Case 9:  Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN), kdc.bank(7, BN), kdc.bank(8, BN)
    Case 10: Input #1, kdc.Index_(BN), kdc.bank(0, BN), kdc.bank(1, BN), kdc.bank(2, BN), kdc.bank(3, BN), kdc.bank(4, BN), kdc.bank(5, BN), kdc.bank(6, BN), kdc.bank(7, BN), kdc.bank(8, BN), kdc.bank(9, BN)
End Select
Loop
Close #1
End Function

Function Fouts(Path_Filename As String, Fields As Integer, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16, bank17, bank18, bank19, bank20)
Open Path_Filename For Output As #2
Select Case Fields
    Case 0: Write #2, msg62
'    Case 1: Write #2, Fields, bank01
'    Case 2: Write #2, Fields, bank01, bank02
'    Case 3: Write #2, Fields, bank01, bank02, bank03
'    Case 4: Write #2, Fields, bank01, bank02, bank03, bank04
'    Case 5: Write #2, Fields, bank01, bank02, bank03, bank04, bank05
'    Case 6: Write #2, Fields, bank01, bank02, bank03, bank04, bank05, bank06
'    Case 7: Write #2, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07
'    Case 8: Write #2, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08
'    Case 9: Write #2, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09
'    Case 10: Write #2, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10
'    Case 11: Write #2, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11
End Select
Close #2
End Function

Public Function Fappnds(Path_Filename As String, Fields As Integer, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16, bank17, bank18, bank19, bank20)
Open Path_Filename For Append As #428571
Select Case Fields
    Case 1: Write #428571, Fields, bank01
    Case 2: Write #428571, Fields, bank01, bank02
    Case 3: Write #428571, Fields, bank01, bank02, bank03
    Case 4: Write #428571, Fields, bank01, bank02, bank03, bank04
    Case 5: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05
    Case 6: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06
    Case 7: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07
    Case 8: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08
    Case 9: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09
    Case 10: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10
    Case 11: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11
    Case 12: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12
    Case 13: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13
    Case 14: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14
    Case 15: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15
    Case 16: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16
    Case 17: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16, bank17
    Case 18: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16, bank17, bank18
    Case 19: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16, bank17, bank18, bank19
    Case 20: Write #428571, Fields, bank01, bank02, bank03, bank04, bank05, bank06, bank07, bank08, bank09, bank10, bank11, bank12, bank13, bank14, bank15, bank16, bank17, bank18, bank19, bank20
End Select
Close #428571
End Function






'    bank02(0 To 1487) As Single
'    bank03(0 To 1487) As Single
'    bank04(0 To 1487) As Single
'    bank05(0 To 1487) As Single
'    bank06(0 To 1487) As Single
'    bank07(0 To 1487) As Single
'    bank08(0 To 1487) As Single
'    bank09(0 To 1487) As Single
'    bank10(0 To 1487) As Single

