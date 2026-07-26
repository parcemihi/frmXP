VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sort N^2"
      Default         =   -1  'True
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Filter"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Load random numbers and sort using N ^ 2."
      Height          =   555
      Left            =   3120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Load and Filter duplicates."
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command5_Click()
Dim DataIn(0 To 30000) As Integer
Dim i As Integer
i = -1
Open App.Path & "\unique.num" For Input As #1
    Do Until EOF(1)
        i = i + 1
        Input #1, DataIn(i)
    Loop
Close #1

KEG_RecursiveSort (DataIn)

Open App.Path & "\inorder.dat" For Append As #1
For i = LBound(DataIn) To UBound(DataIn)
Write #1, DataIn(i)
Next i
Close #1

End Sub

Public Function KEG_RecursiveSort(ByVal YourArray As Variant) As Variant
Dim i As Integer
For i = LBound(YourArray) To UBound(YourArray)
    YourArray(i) = i
Next i

Open App.Path & "\inorder.dat" For Append As #1
For i = LBound(YourArray) To UBound(YourArray)
Write #1, YourArray(i)
Next i
Close #1


End Function




























Private Sub Command1_Click()
'//Generate random number list and write to file named random.num
Dim i As Integer
Dim RN As Integer

Randomize
Open App.Path & "\random.num" For Append As #1
    For i = 1 To 30000
    RN = Int((30000 * Rnd) + 1)
        Write #1, RN
    Next i
Close #1

End Sub

Private Sub Command2_Click()
'Filter Random Numbers
Dim DataIn(0 To 30000) As Integer
Dim i As Integer, n As Integer
Dim CheckNum As Integer
Dim EndOfArray As Integer
Dim IsNumOK As Boolean


Open App.Path & "\random.num" For Input As #1
    Do Until EOF(1)
        Input #1, DataIn(i)
        i = i + 1
    Loop
Close #1

EndOfArray = i - 1
i = 0
n = 1
IsNumOK = True
Do Until i = EndOfArray
    CheckNum = DataIn(i)
    Do Until n = EndOfArray
        If CheckNum = DataIn(n) Then
        
        MsgBox ("Duplicate Number " & CheckNum)
        
            IsNumOK = False
            Do Until n = EndOfArray
                DataIn(n) = DataIn(n + 1)
                n = n + 1
            Loop
            EndOfArray = EndOfArray - 1
        End If
        
        If IsNumOK = True Then
            n = n + 1
        Else
            n = i + 1
            IsNumOK = True
        End If
    Loop
i = i + 1
n = i + 1
Form1.Caption = i
Loop


Open App.Path & "\test.k" For Output As #1
    For i = 0 To EndOfArray - 1
        Write #1, DataIn(i)
    Next i
Close #1
End Sub

Private Sub Command3_Click()
'Sort Random Numbers
Dim i As Integer
Dim NumOfPasses As Integer
Dim SortArray() As Integer
ReDim SortArray(0 To 30000)

Dim FileIsSorted As Boolean
Dim NeedToPassAgain As Boolean

Dim Swapper As Integer
Dim EndOfArray As Integer
EndOfArray = 0

Open App.Path & "\unique.num" For Input As #1
    Do Until EOF(1)
        Input #1, SortArray(EndOfArray)
        EndOfArray = EndOfArray + 1
    Loop
Close #1


ReDim Preserve SortArray(0 To EndOfArray)

FileIsSorted = False
Do Until FileIsSorted = True
    If (SortArray(i) > SortArray(i + 1)) And Not (i = EndOfArray - 1) Then
        Swapper = SortArray(i)
        SortArray(i) = SortArray(i + 1)
        SortArray(i + 1) = Swapper
        i = i + 1
        NeedToPassAgain = True
        
    Else
        If i = EndOfArray - 1 And NeedToPassAgain = True Then
            NeedToPassAgain = False
            i = 0
            NumOfPasses = NumOfPasses + 1
            Command3.Caption = NumOfPasses
        Else
            If NeedToPassAgain = False And i = EndOfArray - 1 Then
                FileIsSorted = True
            Else
                i = i + 1
            End If
        End If
    
    End If

Form1.Caption = Swapper

Loop

Open App.Path & "\sorted.num" For Output As #1
    For i = 0 To EndOfArray
        Write #1, SortArray(i)
    Next i
Close #1

End Sub

Private Sub Command4_Click()
'// Bubble Sort, logarithmic properties
'// QUESTION is this example (the photostat given)
'// placing the random numbers into an array?
Dim SortArray() As Integer
ReDim SortArray(0 To 30000)
Dim EndOfArray As Integer
EndOfArray = 0
Open App.Path & "\unique.num" For Input As #1
    Do Until EOF(1)
        Input #1, SortArray(EndOfArray)
        EndOfArray = EndOfArray + 1
    Loop
Close #1

MyBubbleSort SortArray(0), SortArray(EndOfArray - 1)


End Sub

Public Function MyBubbleSort(First As Integer, Last As Integer)
Dim hi As Integer
Dim lo As Integer
If First > Last Then









hi = First
lo = Last

Else







End If
End Function

