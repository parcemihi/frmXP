VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List4 
      Height          =   7860
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   7860
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   7860
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   7860
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


   Sub Command1_Click()
     Dim lMyArray(0 To 29999) As Long
     Dim vTemp1 As Variant
     Dim vTemp2 As Variant
     Dim vTemp3 As Variant
     Dim iLoop As Integer
     Randomize
     For iLoop = LBound(lMyArray) To UBound(lMyArray)
       lMyArray(iLoop) = Int(Rnd * 30000) + 1
        List1.AddItem lMyArray(iLoop)
     Next iLoop
     vTemp1 = lMyArray
     vTemp2 = lMyArray
     vTemp3 = lMyArray
     Call BubbleSortNumbers(vTemp1)
     Call SelectionSortNumbers(vTemp2)
     Call ShellSortNumbers(vTemp3)
   End Sub

   Sub BubbleSortNumbers(iArray As Variant)
     Dim lLoop1 As Long
     Dim lLoop2 As Long
     Dim lTemp As Long
     For lLoop1 = UBound(iArray) To LBound(iArray) Step -1
       For lLoop2 = LBound(iArray) + 1 To lLoop1

         If iArray(lLoop2 - 1) > iArray(lLoop2) Then
           lTemp = iArray(lLoop2 - 1)
           iArray(lLoop2 - 1) = iArray(lLoop2)
           iArray(lLoop2) = lTemp
         End If
       Next lLoop2
       Form1.Caption = lLoop1
     Next lLoop1
   
     For lLoop1 = LBound(iArray) To UBound(iArray)
       List2.AddItem iArray(lLoop1)
     Next lLoop1
   
   End Sub

   Sub SelectionSortNumbers(vArray As Variant)
     Dim lLoop1 As Long
     Dim lLoop2 As Long
     Dim lMin As Long
     Dim lTemp As Long
     For lLoop1 = LBound(vArray) To UBound(vArray) - 1
       lMin = lLoop1
         For lLoop2 = lLoop1 + 1 To UBound(vArray)
           If vArray(lLoop2) < vArray(lMin) Then lMin = lLoop2
         Next lLoop2
         lTemp = vArray(lMin)
         vArray(lMin) = vArray(lLoop1)
         vArray(lLoop1) = lTemp
       Form1.Caption = lLoop1
     Next lLoop1
   
     For lLoop1 = LBound(vArray) To UBound(vArray)
       List3.AddItem vArray(lLoop1)
     Next lLoop1
   
   
   
   
   End Sub

   Sub ShellSortNumbers(vArray As Variant)
     Dim lLoop1 As Long
     Dim lHold As Long
     Dim lHValue As Long
     Dim lTemp As Long
     lHValue = LBound(vArray)
     Do
       lHValue = 3 * lHValue + 1
     Loop Until lHValue > UBound(vArray)
     Do
       lHValue = lHValue / 3
       For lLoop1 = lHValue + LBound(vArray) To UBound(vArray)
         lTemp = vArray(lLoop1)
         lHold = lLoop1
         Do While vArray(lHold - lHValue) > lTemp
           vArray(lHold) = vArray(lHold - lHValue)
           lHold = lHold - lHValue
           If lHold < lHValue Then Exit Do
         Loop
         vArray(lHold) = lTemp
       Form1.Caption = lLoop1
       Next lLoop1
     Loop Until lHValue = LBound(vArray)
   
   
     For lLoop1 = LBound(vArray) To UBound(vArray)
       List4.AddItem vArray(lLoop1)
     Next lLoop1
   
   
   
   
   End Sub



