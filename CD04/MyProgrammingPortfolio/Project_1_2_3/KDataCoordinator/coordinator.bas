Attribute VB_Name = "DataCoor"
Option Explicit
Global kdc As KDataCrdntr20
Type KDataCrdntr20
FFields As Integer
bank01(0 To 203) As Variant
bank02(0 To 203) As Variant
bank03(0 To 203) As Variant
bank04(0 To 203) As Variant
bank05(0 To 203) As Variant
bank06(0 To 203) As Variant
bank07(0 To 203) As Variant
bank08(0 To 203) As Variant
bank09(0 To 203) As Variant
bank10(0 To 203) As Variant
bank11(0 To 203) As Variant
bank12(0 To 203) As Variant
bank13(0 To 203) As Variant
bank14(0 To 203) As Variant
bank15(0 To 203) As Variant
bank16(0 To 203) As Variant
bank17(0 To 203) As Variant
bank18(0 To 203) As Variant
bank19(0 To 203) As Variant
bank20(0 To 203) As Variant
End Type

Function Fins(Path_Filename As String, Fields As Integer, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20)
Open Path_Filename For Input As #142857
Do Until EOF(142857)

Select Case Fields
Case 1
Input #142857, B01
Case 2
Input #142857, B01, B02
Case 3
Input #142857, B01, B02, B03
Case 4
Input #142857, B01, B02, B03, B04
Case 5
Input #142857, B01, B02, B03, B04, B05
Case 6
Input #142857, B01, B02, B03, B04, B05, B06
Case 7
Input #142857, B01, B02, B03, B04, B05, B06, B07
Case 8
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08
Case 9
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09
Case 10
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10
Case 11
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11
Case 12
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12
Case 13
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13
Case 14
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14
Case 15
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15
Case 16
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16
Case 17
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17
Case 18
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18
Case 19
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19
Case 20
Input #142857, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20
End Select
Loop
Close #142857
End Function

Function Fouts(Path_Filename As String, Fields As Integer, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20)
Open Path_Filename For Output As #285714
Select Case Fields
Case 1
Write #285714, B01
Case 2
Write #285714, B01, B02
Case 3
Write #285714, B01, B02, B03
Case 4
Write #285714, B01, B02, B03, B04
Case 5
Write #285714, B01, B02, B03, B04, B05
Case 6
Write #285714, B01, B02, B03, B04, B05, B06
Case 7
Write #285714, B01, B02, B03, B04, B05, B06, B07
Case 8
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08
Case 9
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09
Case 10
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10
Case 11
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11
Case 12
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12
Case 13
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13
Case 14
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14
Case 15
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15
Case 16
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16
Case 17
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17
Case 18
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18
Case 19
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19
Case 20
Write #285714, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20
End Select


Close #285714
End Function

Function Fappnds(Path_Filename As String, Fields As Integer, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20)
Open Path_Filename For Append As #428571
Select Case Fields
Case 1
Write #428571, B01
Case 2
Write #428571, B01, B02
Case 3
Write #428571, B01, B02, B03
Case 4
Write #428571, B01, B02, B03, B04
Case 5
Write #428571, B01, B02, B03, B04, B05
Case 6
Write #428571, B01, B02, B03, B04, B05, B06
Case 7
Write #428571, B01, B02, B03, B04, B05, B06, B07
Case 8
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08
Case 9
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09
Case 10
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10
Case 11
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11
Case 12
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12
Case 13
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13
Case 14
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14
Case 15
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15
Case 16
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16
Case 17
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17
Case 18
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18
Case 19
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19
Case 20
Write #428571, B01, B02, B03, B04, B05, B06, B07, B08, B09, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20
End Select
Close #428571
End Function

