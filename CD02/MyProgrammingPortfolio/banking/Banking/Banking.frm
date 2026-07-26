VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Taxable"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Top             =   5280
      Width           =   1215
   End
   Begin VB.OptionButton optEntryType 
      Caption         =   "&Credit"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   26
      Top             =   3480
      Width           =   735
   End
   Begin VB.OptionButton optEntryType 
      Caption         =   "&Debit"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   25
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Debits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Cleared"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Hide All"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CheckBox taxable 
      Caption         =   "&Taxable"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.CheckBox cleared 
      Caption         =   "Clea&red"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox DataBoxes 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6960
      TabIndex        =   7
      Text            =   "TextData"
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox DataBoxes 
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   6
      Text            =   "TextData"
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox DataBoxes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Text            =   "TextData"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox DataBoxes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Text            =   "TextData"
      Top             =   720
      Width           =   735
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   1560
      TabIndex        =   10
      Top             =   1200
      Width           =   6975
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Amount"
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   14
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   13
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   12
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSort 
      Caption         =   "Number"
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   ">>"
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "&>"
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdNavi 
      Caption         =   "&<"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdWriteDataFile 
      Caption         =   "&Save"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddEntry 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Current Balance"
      Height          =   195
      Left            =   5745
      TabIndex        =   29
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Available"
      Height          =   195
      Left            =   6240
      TabIndex        =   28
      Top             =   120
      Width           =   645
   End
   Begin VB.Label lblAvailable 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   24
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      Top             =   360
      Width           =   1440
   End
   Begin VB.Label DUMP 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4200
      TabIndex        =   17
      Top             =   2280
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Available As Double
Dim SortOrder(0 To 3) As Boolean
Dim IsSorted As Boolean
Private Sub cmdFilter_Click(Index As Integer)
Dim crd As Integer, i As Integer
If Index = 0 Then
    Cap_Swap cmdFilter(Index), "Hide All", "Show All"
    FontAttrib_Swap cmdFilter(Index), True, False
End If
If Index = 1 Then
End If
If Index = 2 Then
End If
If Index = 3 Then
End If
If Index = 4 Then
    Cap_Swap cmdFilter(Index), "All", "Taxable", "Non-Taxable"
Else
    FontAttrib_Swap cmdFilter(Index), True, False
End If
End Sub
Private Function Cap_Swap(ButtonName As Variant, Name1 As String, Name2 As String, Optional Name3 As String)
With ButtonName
    If .Caption = Name1 Then
        .Caption = Name2
    ElseIf .Caption = Name2 And Not Len(Name3) = 0 Then
        .Caption = Name3
    Else
        .Caption = Name1
    End If
End With
End Function
Private Function FontAttrib_Swap(ButtonName As Variant, Bold As Boolean, Italic As Boolean)
With ButtonName
    If Bold = True Then
        If .FontBold = True Then
            .FontBold = False
        Else
            .FontBold = True
        End If
    End If
    
    If Italic = True Then
        If .FontItalic = True Then
            .FontItalic = False
        Else
            .FontItalic = True
        End If
    End If

End With
End Function
Private Sub Form_Load()
ReDim Entry(0 To 5000)
Form1.Width = 8670
ArrayPointer = 0
EndOf_Array = -1
Open App.Path & "\format.dat" For Input As #1
    Do Until EOF(1)
        EndOf_Array = EndOf_Array + 1
        Input #1, Entry(EndOf_Array).Date, _
                  Entry(EndOf_Array).Number, _
                  Entry(EndOf_Array).Description, _
                  Entry(EndOf_Array).Amount, _
                  Entry(EndOf_Array).cleared, _
                  Entry(EndOf_Array).taxable, _
                  Entry(EndOf_Array).EntType
    Loop
Close #1
FillDataboxes
Load_List
Get_Total
End Sub
Private Sub FillDataboxes()
DataBoxDirect = False
DataBoxes(0).Text = Entry(ArrayPointer).Number
DataBoxes(1).Text = Entry(ArrayPointer).Date
DataBoxes(2).Text = Entry(ArrayPointer).Description
DataBoxes(3).Text = Format(Entry(ArrayPointer).Amount, "Currency")
optEntryType(Entry(ArrayPointer).EntType).Value = True
cleared.Value = Entry(ArrayPointer).cleared
taxable.Value = Entry(ArrayPointer).taxable
DataBoxDirect = True
End Sub
Private Sub Load_List()
Dim tmp As String
Dim i As Integer
List1.Clear
Dim FixNumber As String, FixDate As String, FixDescription As String * 30, FixAmount As String
For i = 0 To EndOf_Array
FixNumber = Entry(i).Number: FixDate = Entry(i).Date: FixDescription = Entry(i).Description
FixAmount = Format(Abs(Entry(i).Amount), "#,##0.00")
Select Case Len(FixAmount)
    Case Is = 4: FixAmount = "    " & FixAmount
    Case Is = 5: FixAmount = "   " & FixAmount
    Case Is = 6: FixAmount = "  " & FixAmount
    Case Is = 7: FixAmount = " " & FixAmount
End Select
If Val(Entry(i).Amount) < 0 Then
    FixAmount = "$(" & FixAmount & ")"
Else
    FixAmount = "$ " & FixAmount & " "
End If
Select Case Len(FixNumber)
    Case Is = 1: FixNumber = " 000" & FixNumber
    Case Is = 2: FixNumber = " 00" & FixNumber
    Case Is = 3: FixNumber = " 0" & FixNumber
    Case Is = 4: FixNumber = " " & FixNumber
End Select
Select Case Len(FixDate)
    Case Is = 6
        FixDate = "0" & FixDate
        FixDate = Mid(FixDate, 1, 3) & "0" & Mid(FixDate, 4, 4)
    Case Is = 7
        If Val(Mid(FixDate, 1, 2)) < 10 Then
            FixDate = "0" & FixDate
        Else
            FixDate = Mid(FixDate, 1, 3) & "0" & Mid(FixDate, 4, 4)
        End If
End Select
List1.AddItem FixNumber & "   " & FixDate & "   " & FixDescription & FixAmount
Next i
End Sub
Private Sub cmdNavi_Click(Index As Integer) 'Trap scroll lock. Timer loop through records. Command Button order toggle.
Select Case Index
Case 0
        ArrayPointer = 0
Case 1
        If ArrayPointer > 0 Then
            ArrayPointer = ArrayPointer - 1
        Else
            ArrayPointer = EndOf_Array
        End If
Case 2
        If ArrayPointer < EndOf_Array Then
            ArrayPointer = ArrayPointer + 1
        Else
            ArrayPointer = 0
        End If
Case 3
        ArrayPointer = EndOf_Array
End Select
FillDataboxes
End Sub
Private Sub cmdDelete_Click()
Dim msg As Integer
'msg = MsgBox("Are you sure you want to delete this record?", vbYesNo)
'If msg = vbYes Then
    sort.DELETE_RECORD
    Load_List
    FillDataboxes
'End If
End Sub
Private Sub cleared_Click()

If DataBoxDirect = True Then
Entry(ArrayPointer).cleared = cleared.Value
Get_Total
End If
End Sub
Private Sub List1_Click()
ArrayPointer = List1.ListIndex
FillDataboxes
End Sub

Private Sub optEntryType_Click(Index As Integer)
Entry(ArrayPointer).EntType = Index
End Sub

Private Sub taxable_Click()
If DataBoxDirect = True Then
Entry(ArrayPointer).taxable = taxable.Value
End If
End Sub
Private Sub DataBoxes_LostFocus(Index As Integer)
On Error GoTo ErrHandler
If Index = 0 Then
    Entry(ArrayPointer).Number = DataBoxes(0).Text
End If
If Index = 1 Then
    Entry(ArrayPointer).Date = DataBoxes(1).Text
End If
If Index = 2 Then
    Entry(ArrayPointer).Description = DataBoxes(2).Text
End If
If Index = 3 Then
    Entry(ArrayPointer).Amount = DataBoxes(3).Text
    Get_Total
End If

Exit Sub
ErrHandler:
If Index = 0 Then
MsgBox "Numerics only.", vbOKOnly
ElseIf Index = 1 Then
MsgBox "Invalid Date.", vbOKOnly
ElseIf Index = 2 Then
MsgBox "Don't know the cause of error!.", vbOKOnly
ElseIf Index = 3 Then
MsgBox "Not a number ;(", vbOKOnly
End If
DataBoxes_GotFocus (Index)
End Sub
Private Function Get_Total()
Dim i As Integer
Total = 0
Available = 0
For i = 0 To EndOf_Array
    If Entry(i).cleared = 1 Then
        Total = Total + Entry(i).Amount
    End If
    Available = Available + Entry(i).Amount
Next i
Load_List
FillDataboxes
lblTotal.Caption = Format(Total, "Currency")
lblAvailable.Caption = Format(Available, "Currency")
End Function
Private Sub cmdAddEntry_Click()
EndOf_Array = EndOf_Array + 1
ArrayPointer = EndOf_Array
FillDataboxes
DataBoxes(0).SetFocus
End Sub
Private Sub cmdWriteDataFile_Click()
Dim i As Integer
Open App.Path & "\format.dat" For Output As #1
For i = 0 To EndOf_Array
'Entry(i).cleared = 1
        Write #1, Entry(i).Date, _
                  Entry(i).Number, _
                  Entry(i).Description, _
                  Entry(i).Amount, _
                  Entry(i).cleared, _
                  Entry(i).taxable, _
                  Entry(i).EntType
Next i
Close #1
Form_Load
End Sub
Private Sub DataBoxes_GotFocus(Index As Integer)
DataBoxes(Index).SelStart = 0
DataBoxes(Index).SelLength = Len(DataBoxes(Index).Text)
End Sub
Private Sub cmdSort_Click(Index As Integer)
If SortOrder(Index) = True Then
SortOrder(Index) = False
Else
SortOrder(Index) = True
End If

Select Case Index
    Case Is = 0
    SORT_RECORD Index, SortOrder(Index)
    Case Is = 1
    SORT_RECORD Index, SortOrder(Index)
    Case Is = 2
    SORT_RECORD Index, SortOrder(Index)
    Case Is = 3
    SORT_RECORD Index, SortOrder(Index)
End Select
Load_List
End Sub

'    Do Until Entry(ArrayPointer).Delete = False Or _
'                           ArrayPointer = EndOf_Array
'        If Entry(ArrayPointer).Delete = True Then
'            ArrayPointer = ArrayPointer + 1
'        End If
'    Loop
'    Do Until Entry(ArrayPointer).Delete = False Or _
'                           ArrayPointer = 0
'        If Entry(ArrayPointer).Delete = True Then
'            ArrayPointer = ArrayPointer - 1
'        End If
'    Loop

'Load_List
'If Not ArrayPointer = EndOf_Array Then
'    Do Until Entry(ArrayPointer).Delete = False
'    ArrayPointer = ArrayPointer + 1
'    Loop
'    FillDataboxes
'End If
